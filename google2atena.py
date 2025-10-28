# google2atena.py
import io, re, csv, os
from typing import Dict, List, Tuple
from flask import Flask, render_template, request, send_file, abort, Response

app = Flask(__name__)

# =========================
# 文字種変換ユーティリティ
# =========================

FULLWIDTH_OFFSET = ord("！") - ord("!")
ASCII_MIN, ASCII_MAX = 33, 126

def to_zenkaku(s: str) -> str:
    """ASCII英数記号とスペースを全角化。その他はそのまま。"""
    if not isinstance(s, str): return ""
    out = []
    for ch in s:
        code = ord(ch)
        if ch == " ":
            out.append("　")  # 全角スペース
        elif 33 <= code <= 126:
            out.append(chr(code + FULLWIDTH_OFFSET))
        else:
            out.append(ch)
    return "".join(out)

def to_hankaku_simple(s: str) -> str:
    """最小限の半角化（英字/数字/記号/スペースを半角へ）。"""
    if not isinstance(s, str): return ""
    out = []
    for ch in s:
        if ch == "　":
            out.append(" ")
            continue
        code = ord(ch)
        # 全角英数記号の範囲
        if 65281 <= code <= 65374:
            out.append(chr(code - FULLWIDTH_OFFSET))
        else:
            out.append(ch)
    return "".join(out)

def normalize_for_phone_email_postal(s: str) -> str:
    """電話・メール・郵便番号は半角に統一し、前後空白を詰める。"""
    s = (s or "")
    s = to_hankaku_simple(s)
    return s.strip()

# =========================
# 会社名の先頭語処理
# =========================

CORP_PREFIXES = [
    "株式会社", "有限会社", "合同会社",
    "一般社団法人", "一般財団法人",
    "公益社団法人", "公益財団法人",
    "社団法人", "財団法人",
]

def insert_space_after_corp_prefix(name: str) -> str:
    """会社名が法人格で始まるとき、全角スペースを1つ挿入（重複は避ける）。"""
    s = name or ""
    for p in CORP_PREFIXES:
        if s.startswith(p):
            rest = s[len(p):]
            # すでに全角スペースが続いていればそのまま
            if rest.startswith("　"):
                return s
            # 半角スペースやスペースなしで続いていれば全角スペースに正規化
            rest = rest.lstrip(" ")
            return p + "　" + rest
    return s

# =========================
# 住所正規化
# =========================
# 方針：
# 1) 全体をまず全角化
# 2) 「丁目/番/番地/号」を "N−N−N" へ正規化（Nは全角数字）
# 3) 建物ワード（ビル/マンション/…）検出で建物名＋部屋番号/階を分離
# 4) 建物名が無い場合でも、住所に「丁目/番/号」が含まれ、末尾 "−数字{3,}" なら部屋番号と判定
# 5) 地番のみ（例：海士944 等）は部屋番号と誤判定しない

BUILDING_KEYWORDS = [
    "ビル", "マンション", "ハイツ", "アパート", "コーポ",
    "タワー", "ヒルズ", "荘", "レジデンス", "ハウス", "テラス", "メゾン",
]

def unify_hyphen(s: str) -> str:
    """ASCIIハイフン・長音などを全角ハイフン（U+FF0D '－'）へ寄せる。"""
    return re.sub(r"[‐-–—−-]", "－", s)

def zenkaku_digits(s: str) -> str:
    """ASCII数字を全角数字へ。"""
    tbl = str.maketrans("0123456789", "０１２３４５６７８９")
    return s.translate(tbl)

def normalize_address(addr: str) -> Tuple[str, str]:
    """
    住所文字列を正規化して (base, building) を返す。
    base: 「…Ｎ−Ｎ−Ｎ」まで
    building: 「　ネコノスビル　＃２０３」や「　ネコノスビル　２Ｆ」等（先頭には全角スペースを付けて返す）
    """
    if not addr: return "", ""
    s = addr.strip()

    # まず全角化（英数記号/スペース）
    s = to_zenkaku(s)
    # ハイフン類正規化 → 全角ハイフン（U+FF0D）
    s = unify_hyphen(s)
    # 数字は全角
    s = zenkaku_digits(s)

    # (1) 丁目/番/番地/号 → "N－N－N" 化
    #   * "丁目" を "－" に
    #   * "番地"/"番" を "－" に
    #   * "号" は削除（直前の数字にぶら下がっている想定）
    s = re.sub(r"([０-９]+)\s*丁目", r"\1－", s)
    s = re.sub(r"([０-９]+)\s*番地?", r"\1－", s)
    s = re.sub(r"([０-９]+)\s*号(室)?", r"\1", s)

    # (2) 連続ハイフン「－－」→「－」
    s = re.sub(r"－{2,}", "－", s)

    # 建物名の検出
    has_building_word = any(k in s for k in BUILDING_KEYWORDS)

    # (3) 建物名＋部屋/階の抽出（例：ネコノスビル２階 / サンライトビル３０２）
    building = ""
    # パターンA： <建物名><数字>階
    m = re.search(rf"(.+?({'|'.join(BUILDING_KEYWORDS)}))\s*([０-９]+)\s*階", s)
    if m:
        base = s[:m.start()].rstrip()
        bname = m.group(1).strip()
        floor = m.group(3)
        building = f"　{bname}　{floor}Ｆ"
        return base, building

    # パターンB： <建物名><数字{1,4}>（号室省略）
    m = re.search(rf"(.+?({'|'.join(BUILDING_KEYWORDS)}))\s*([０-９]{{1,4}})$", s)
    if m:
        base = s[:m.start()].rstrip()
        bname = m.group(1).strip()
        room = m.group(3)
        building = f"　{bname}　＃{room}"
        return base, building

    # パターンC： <建物名><数字{1,4}>号室
    m = re.search(rf"(.+?({'|'.join(BUILDING_KEYWORDS)}))\s*([０-９]{{1,4}})\s*号室", s)
    if m:
        base = s[:m.start()].rstrip()
        bname = m.group(1).strip()
        room = m.group(3)
        building = f"　{bname}　＃{room}"
        return base, building

    # (4) 建物名が無い場合でも、丁目/番/号が出ていて末尾が「－数字{3,}」なら部屋番号扱い
    if ("丁目" in addr) or ("番" in addr) or ("号" in addr):
        m = re.search(r"(.+?)－([０-９]{3,})$", s)
        if m:
            base = m.group(1).rstrip()
            room = m.group(2)
            building = f"　＃{room}"
            return base, building

    # (5) 上記に当てはまらなければ全体を base として返す（地番3桁などはそのまま）
    return s, ""

# =========================
# Google CSV パース
# =========================

def parse_google_row(row: Dict[str, str]) -> Dict[str, str]:
    """Google連絡先の1行を、必要な辞書に集約（名前・会社・連絡先・住所など）。"""
    get = lambda k: (row.get(k) or "").strip()

    # 名前系
    first = get("First Name")
    middle = get("Middle Name")
    last = get("Last Name")
    pf_first = get("Phonetic First Name")
    pf_middle = get("Phonetic Middle Name")
    pf_last = get("Phonetic Last Name")
    nickname = get("Nickname")
    notes = get("Notes")
    birthday = get("Birthday")

    # 組織
    org = get("Organization Name")
    dept = get("Organization Department")
    title = get("Organization Title")

    # E-mail 集約（work/home/other）
    emails = {"work": [], "home": [], "other": []}
    phones = {"work": [], "home": [], "other": []}

    # 可変列: "E-mail n - Type" / "E-mail n - Value"
    for n in range(1, 50):
        tkey = f"E-mail {n} - Type"
        vkey = f"E-mail {n} - Value"
        if tkey in row or vkey in row:
            typ = get(tkey).lower()
            val = get(vkey)
            if val:
                if "work" in typ:
                    emails["work"].append(val)
                elif "home" in typ:
                    emails["home"].append(val)
                else:
                    emails["other"].append(val)
        else:
            break

    # 可変列: "Phone n - Type" / "Phone n - Value"
    for n in range(1, 50):
        tkey = f"Phone {n} - Type"
        vkey = f"Phone {n} - Value"
        if tkey in row or vkey in row:
            typ = get(tkey).lower()
            val = get(vkey)
            if val:
                if "work" in typ:
                    phones["work"].append(val)
                elif "home" in typ:
                    phones["home"].append(val)
                else:
                    phones["other"].append(val)
        else:
            break

    # 住所（work/home/other）: Region + City + Street / Postal Code
    addr = {
        "work": {"postal": "", "line": ""},
        "home": {"postal": "", "line": ""},
        "other": {"postal": "", "line": ""},
    }
    for n in range(1, 50):
        tkey = f"Address {n} - Type"
        reg = f"Address {n} - Region"
        city = f"Address {n} - City"
        street = f"Address {n} - Street"
        postal = f"Address {n} - Postal Code"
        if tkey in row or reg in row or city in row or street in row or postal in row:
            typ = get(tkey).lower()
            line = " ".join([get(reg), get(city), get(street)]).strip()
            pcd = get(postal)
            if "work" in typ:
                addr["work"]["postal"] = pcd or addr["work"]["postal"]
                addr["work"]["line"] = line or addr["work"]["line"]
            elif "home" in typ:
                addr["home"]["postal"] = pcd or addr["home"]["postal"]
                addr["home"]["line"] = line or addr["home"]["line"]
            else:
                addr["other"]["postal"] = pcd or addr["other"]["postal"]
                addr["other"]["line"] = line or addr["other"]["line"]
        else:
            break

    return {
        "first": first, "middle": middle, "last": last,
        "pf_first": pf_first, "pf_middle": pf_middle, "pf_last": pf_last,
        "nickname": nickname, "notes": notes, "birthday": birthday,
        "org": org, "dept": dept, "title": title,
        "emails": emails, "phones": phones, "addrs": addr
    }

# =========================
# 宛名職人 出力列（アップロードの result.csv と同一）
# =========================

TARGET_COLUMNS = [
    "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称","ニックネーム","旧姓","宛先",
    "自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話1〜10","自宅IM ID1〜10","自宅E-mail1〜5","自宅URL1〜5","自宅Social1〜10",
    "会社〒","会社住所1","会社住所2","会社住所3","会社電話1〜10","会社IM ID1〜10","会社E-mail1〜5","会社URL1〜5","会社Social1〜10",
    "その他〒","その他住所1","その他住所2","その他住所3","その他電話1〜10","その他IM ID1〜10","その他E-mail1〜5","その他URL1〜5","その他Social1〜10",
    "会社名かな","会社名","部署名1","部署名2","役職名","連名1〜20","連名ふりがな1〜20","連名敬称1〜20","連名誕生日1〜20",
    "メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
]

# =========================
# 1レコード → 宛名職人行
# =========================

def build_atena_row(g: Dict[str,str]) -> Dict[str,str]:
    """Google行（集約後辞書）→ 宛名職人フォーマット1行辞書"""

    # 氏名
    last = g["last"]; first = g["first"]; middle = g["middle"]
    pf_last = g["pf_last"]; pf_first = g["pf_first"]; pf_middle = g["pf_middle"]
    nickname = g["nickname"]
    notes = g["notes"]
    birthday = g["birthday"]

    # 会社
    org = insert_space_after_corp_prefix(g["org"])
    dept = g["dept"]
    title = g["title"]

    # 住所（work/home/other）
    aw = g["addrs"]["work"]; ah = g["addrs"]["home"]; ao = g["addrs"]["other"]

    # 住所正規化（line → (base, building)）
    def pack_addr(line: str):
        base, bld = normalize_address(line)
        return base, bld

    w_base, w_bld = pack_addr(aw["line"])
    h_base, h_bld = pack_addr(ah["line"])
    o_base, o_bld = pack_addr(ao["line"])

    # 文字種ルール適用：
    #  - 電話/メール/郵便番号：半角
    #  - その他：全角
    def join_and_hankaku(vals: List[str]) -> str:
        vals = [normalize_for_phone_email_postal(v) for v in vals if v]
        return ";".join(v for v in vals if v)

    # phone/email
    phone_work = join_and_hankaku(g["phones"]["work"])
    phone_home = join_and_hankaku(g["phones"]["home"])
    phone_other = join_and_hankaku(g["phones"]["other"])

    email_work = join_and_hankaku(g["emails"]["work"])
    email_home = join_and_hankaku(g["emails"]["home"])
    email_other = join_and_hankaku(g["emails"]["other"])

    # postal
    postal_work = normalize_for_phone_email_postal(aw["postal"])
    postal_home = normalize_for_phone_email_postal(ah["postal"])
    postal_other = normalize_for_phone_email_postal(ao["postal"])

    # ふりがな補完
    last_k = pf_last or ""
    first_k = pf_first or ""
    middle_k = pf_middle or ""

    # 宛先：会社情報があれば「会社」、なければ「自宅」
    atesaki = "会社" if org else "自宅"

    # フルネーム
    sei = last or ""
    mei = first or ""
    sei_mei = (first + last) if (first or last) else ""  # 宛名職人の「姓名」（順序は要件に合わせ調整可）
    sei_mei_kana = (first_k + last_k) if (first_k or last_k) else ""

    # 全角化すべきフィールド群
    def Z(x: str) -> str:
        return to_zenkaku(x or "")

    row = {k: "" for k in TARGET_COLUMNS}

    # --- 基本項目 ---
    row["姓"] = Z(sei)
    row["名"] = Z(mei)
    row["姓かな"] = Z(last_k)
    row["名かな"] = Z(first_k)
    row["姓名"] = Z(sei_mei)
    row["姓名かな"] = Z(sei_mei_kana)
    row["ミドルネーム"] = Z(middle)
    row["ミドルネームかな"] = Z(middle_k)
    row["敬称"] = ""  # 不明
    row["ニックネーム"] = Z(nickname)
    row["旧姓"] = ""
    row["宛先"] = Z(atesaki)

    # --- 自宅 ---
    row["自宅〒"] = postal_home
    hb, hl2 = h_base, h_bld
    row["自宅住所1"] = Z(hb)
    row["自宅住所2"] = Z(hl2)
    row["自宅住所3"] = ""
    row["自宅電話1〜10"] = phone_home
    row["自宅IM ID1〜10"] = ""
    row["自宅E-mail1〜5"] = email_home
    row["自宅URL1〜5"] = ""
    row["自宅Social1〜10"] = ""

    # --- 会社 ---
    row["会社〒"] = postal_work
    wb, wl2 = w_base, w_bld
    row["会社住所1"] = Z(wb)
    row["会社住所2"] = Z(wl2)
    row["会社住所3"] = ""
    row["会社電話1〜10"] = phone_work
    row["会社IM ID1〜10"] = ""
    row["会社E-mail1〜5"] = email_work
    row["会社URL1〜5"] = ""
    row["会社Social1〜10"] = ""

    # --- その他 ---
    row["その他〒"] = postal_other
    ob, ol2 = o_base, o_bld
    row["その他住所1"] = Z(ob)
    row["その他住所2"] = Z(ol2)
    row["その他住所3"] = ""
    row["その他電話1〜10"] = phone_other
    row["その他IM ID1〜10"] = ""
    row["その他E-mail1〜5"] = email_other
    row["その他URL1〜5"] = ""
    row["その他Social1〜10"] = ""

    # --- 会社情報 ---
    row["会社名かな"] = ""  # 情報なし
    row["会社名"] = Z(org)
    row["部署名1"] = Z(dept)
    row["部署名2"] = ""
    row["役職名"] = Z(title)

    # --- 連名など（未使用） ---
    row["連名1〜20"] = ""
    row["連名ふりがな1〜20"] = ""
    row["連名敬称1〜20"] = ""
    row["連名誕生日1〜20"] = ""

    # --- メモ/備考/誕生日ほか ---
    # Google「Notes」を備考1へ
    row["メモ1"] = ""
    row["メモ2"] = ""
    row["メモ3"] = ""
    row["メモ4"] = ""
    row["メモ5"] = ""
    row["備考1"] = Z(notes)
    row["備考2"] = ""
    row["備考3"] = ""
    row["誕生日"] = Z(birthday)  # 例: 1980-01-23 → 全角数字へ
    row["性別"] = ""
    row["血液型"] = ""
    row["趣味"] = ""
    row["性格"] = ""

    return row

# =========================
# CSV 読み込み/書き出し
# =========================

def read_google_csv(file_bytes: bytes) -> List[Dict[str,str]]:
    encodings = ["utf-8", "utf-8-sig", "cp932", "shift_jis"]
    last_err = None
    for enc in encodings:
        try:
            f = io.StringIO(file_bytes.decode(enc))
            r = csv.DictReader(f)
            rows = []
            for row in r:
                # None → "" に
                clean = { (k or "").strip(): (v if v is not None else "") for k,v in row.items() }
                rows.append(clean)
            return rows
        except Exception as e:
            last_err = e
            continue
    raise ValueError(f"CSV読み込み失敗（文字コード）: {last_err}")

def write_atena_csv(rows: List[Dict[str,str]]) -> bytes:
    buf = io.StringIO(newline="")
    w = csv.DictWriter(buf, fieldnames=TARGET_COLUMNS)
    w.writeheader()
    for r in rows:
        w.writerow(r)
    return buf.getvalue().encode("utf-8")

# =========================
# Basic認証（環境変数でON/OFF）
# =========================

@app.before_request
def require_basic_auth():
    if request.path.startswith("/static/") or request.path.startswith("/healthz"):
        return
    user = os.environ.get("BASIC_AUTH_USER", "")
    pw   = os.environ.get("BASIC_AUTH_PASS", "")
    if not (user and pw):
        return
    auth = request.authorization
    if not (auth and auth.username == user and auth.password == pw):
        return Response("Authentication required", 401,
                        {"WWW-Authenticate": 'Basic realm="google2atena"'})

# =========================
# ルーティング
# =========================

@app.get("/")
def index():
    return render_template("index.html")

@app.post("/convert")
def convert():
    if "file" not in request.files:
        abort(400, "file is required")
    data = request.files["file"].read()
    try:
        grows = read_google_csv(data)
    except Exception as e:
        abort(400, f"CSVの読み込みに失敗しました: {e}")

    out_rows = []
    for row in grows:
        g = parse_google_row(row)
        out_rows.append(build_atena_row(g))

    payload = write_atena_csv(out_rows)
    return send_file(
        io.BytesIO(payload),
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name="google_converted.csv",
    )

@app.get("/healthz")
def healthz():
    return {"ok": True}

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5001))  # ← PORT環境変数を利用
    app.run(host="0.0.0.0", port=port)
