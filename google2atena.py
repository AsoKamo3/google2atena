# google2atena.py (修正版)
import io, re, csv, os
from flask import Flask, render_template, request, send_file, abort, Response

app = Flask(__name__)

# =========================
# 文字種変換ユーティリティ
# =========================

FULLWIDTH_OFFSET = ord("！") - ord("!")
ASCII_MIN, ASCII_MAX = 33, 126

def to_zenkaku(s: str) -> str:
    if not isinstance(s, str): return ""
    out = []
    for ch in s:
        code = ord(ch)
        if ch == " ":
            out.append("　")
        elif 33 <= code <= 126:
            out.append(chr(code + FULLWIDTH_OFFSET))
        else:
            out.append(ch)
    return "".join(out)

def to_hankaku_simple(s: str) -> str:
    if not isinstance(s, str): return ""
    out = []
    for ch in s:
        if ch == "　":
            out.append(" ")
            continue
        code = ord(ch)
        if 65281 <= code <= 65374:
            out.append(chr(code - FULLWIDTH_OFFSET))
        else:
            out.append(ch)
    return "".join(out)

def normalize_for_phone_email_postal(s: str) -> str:
    s = (s or "")
    s = to_hankaku_simple(s)
    return s.strip()

# =========================
# 会社名の法人格スペース処理
# =========================

CORP_PREFIXES = [
    "株式会社", "有限会社", "合同会社",
    "一般社団法人", "一般財団法人",
    "公益社団法人", "公益財団法人",
    "社団法人", "財団法人",
]

def insert_space_after_corp_prefix(name: str) -> str:
    s = name or ""
    for p in CORP_PREFIXES:
        if s.startswith(p):
            rest = s[len(p):]
            if rest.startswith("　"):
                return s
            rest = rest.lstrip(" ")
            return p + "　" + rest
    return s

# =========================
# 住所正規化
# =========================

BUILDING_KEYWORDS = [
    "ビル", "マンション", "ハイツ", "アパート", "コーポ",
    "タワー", "ヒルズ", "荘", "レジデンス", "ハウス", "テラス", "メゾン",
]

def unify_hyphen(s: str) -> str:
    return re.sub(r"[‐-–—−-]", "－", s)

def zenkaku_digits(s: str) -> str:
    tbl = str.maketrans("0123456789", "０１２３４５６７８９")
    return s.translate(tbl)

def normalize_address(addr: str):
    if not addr: return "", ""
    s = addr.strip()
    s = to_zenkaku(s)
    s = unify_hyphen(s)
    s = zenkaku_digits(s)
    s = re.sub(r"([０-９]+)\s*丁目", r"\1－", s)
    s = re.sub(r"([０-９]+)\s*番地?", r"\1－", s)
    s = re.sub(r"([０-９]+)\s*号(室)?", r"\1", s)
    s = re.sub(r"－{2,}", "－", s)

    has_building_word = any(k in s for k in BUILDING_KEYWORDS)
    building = ""

    m = re.search(rf"(.+?({'|'.join(BUILDING_KEYWORDS)}))\s*([０-９]+)\s*階", s)
    if m:
        base = s[:m.start()].rstrip()
        bname = m.group(1).strip()
        floor = m.group(3)
        building = f"　{bname}　{floor}Ｆ"
        return base, building

    m = re.search(rf"(.+?({'|'.join(BUILDING_KEYWORDS)}))\s*([０-９]{{1,4}})$", s)
    if m:
        base = s[:m.start()].rstrip()
        bname = m.group(1).strip()
        room = m.group(3)
        building = f"　{bname}　＃{room}"
        return base, building

    m = re.search(rf"(.+?({'|'.join(BUILDING_KEYWORDS)}))\s*([０-９]{{1,4}})\s*号室", s)
    if m:
        base = s[:m.start()].rstrip()
        bname = m.group(1).strip()
        room = m.group(3)
        building = f"　{bname}　＃{room}"
        return base, building

    if ("丁目" in addr) or ("番" in addr) or ("号" in addr):
        m = re.search(r"(.+?)－([０-９]{3,})$", s)
        if m:
            base = m.group(1).rstrip()
            room = m.group(2)
            building = f"　＃{room}"
            return base, building

    return s, ""

# =========================
# Google CSV パース
# =========================

def parse_google_row(row):
    get = lambda k: (row.get(k) or "").strip()
    first = get("First Name")
    middle = get("Middle Name")
    last = get("Last Name")
    pf_first = get("Phonetic First Name")
    pf_middle = get("Phonetic Middle Name")
    pf_last = get("Phonetic Last Name")
    nickname = get("Nickname")
    notes = get("Notes")
    birthday = get("Birthday")

    org = get("Organization Name")
    dept = get("Organization Department")
    title = get("Organization Title")

    emails = {"work": [], "home": [], "other": []}
    phones = {"work": [], "home": [], "other": []}

    for n in range(1, 50):
        tkey = f"E-mail {n} - Type"
        vkey = f"E-mail {n} - Value"
        if tkey in row or vkey in row:
            typ = get(tkey).strip().lower()
            val = get(vkey)
            if val:
                if typ in ("work", "勤務先", "会社", "職場"):
                    emails["work"].append(val)
                elif typ in ("home", "自宅", "ホーム"):
                    emails["home"].append(val)
                else:
                    emails["other"].append(val)
        else:
            break

    for n in range(1, 50):
        tkey = f"Phone {n} - Type"
        vkey = f"Phone {n} - Value"
        if tkey in row or vkey in row:
            typ = get(tkey).strip().lower()
            val = get(vkey)
            if val:
                if typ in ("work", "勤務先", "会社", "職場"):
                    phones["work"].append(val)
                elif typ in ("home", "自宅", "ホーム"):
                    phones["home"].append(val)
                else:
                    phones["other"].append(val)
        else:
            break

    addr = {"work": {"postal": "", "line": ""},
            "home": {"postal": "", "line": ""},
            "other": {"postal": "", "line": ""}}
    for n in range(1, 50):
        tkey = f"Address {n} - Type"
        reg = f"Address {n} - Region"
        city = f"Address {n} - City"
        street = f"Address {n} - Street"
        postal = f"Address {n} - Postal Code"
        if tkey in row or reg in row or city in row or street in row or postal in row:
            typ = get(tkey).strip().lower()
            line = " ".join([get(reg), get(city), get(street)]).strip()
            pcd = get(postal)
            if typ in ("work", "勤務先", "会社", "職場"):
                addr["work"]["postal"] = pcd or addr["work"]["postal"]
                addr["work"]["line"] = line or addr["work"]["line"]
            elif typ in ("home", "自宅", "ホーム"):
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
# 宛名職人出力処理（修正版）
# =========================

def build_atena_row(g):
    last = g["last"]; first = g["first"]; middle = g["middle"]
    pf_last = g["pf_last"]; pf_first = g["pf_first"]; pf_middle = g["pf_middle"]
    nickname = g["nickname"]; notes = g["notes"]; birthday = g["birthday"]
    org = insert_space_after_corp_prefix(g["org"])
    dept = g["dept"]; title = g["title"]
    aw, ah, ao = g["addrs"]["work"], g["addrs"]["home"], g["addrs"]["other"]

    def pack_addr(line):
        base, bld = normalize_address(line)
        return base, bld

    w_base, w_bld = pack_addr(aw["line"])
    h_base, h_bld = pack_addr(ah["line"])
    o_base, o_bld = pack_addr(ao["line"])

    def join_and_hankaku(vals):
        vals = [normalize_for_phone_email_postal(v) for v in vals if v]
        return ";".join(v for v in vals if v)

    phone_work = join_and_hankaku(g["phones"]["work"])
    phone_home = join_and_hankaku(g["phones"]["home"])
    phone_other = join_and_hankaku(g["phones"]["other"])
    email_work = join_and_hankaku(g["emails"]["work"])
    email_home = join_and_hankaku(g["emails"]["home"])
    email_other = join_and_hankaku(g["emails"]["other"])

    postal_work = normalize_for_phone_email_postal(aw["postal"])
    postal_home = normalize_for_phone_email_postal(ah["postal"])
    postal_other = normalize_for_phone_email_postal(ao["postal"])

    last_k, first_k, middle_k = g["pf_last"], g["pf_first"], g["pf_middle"]
    atesaki = "会社" if org else "自宅"

    # 👇 修正版：姓名・姓名かなを「姓＋名」順にする
    sei_mei = (last + "　" + first).strip()
    sei_mei_kana = (last_k + "　" + first_k).strip()

    def Z(x): return to_zenkaku(x or "")
    row = {"姓": Z(last), "名": Z(first), "姓かな": Z(last_k), "名かな": Z(first_k),
           "姓名": Z(sei_mei), "姓名かな": Z(sei_mei_kana),
           "会社名": Z(org), "部署名1": Z(dept), "役職名": Z(title),
           "会社住所1": Z(w_base), "会社住所2": Z(w_bld),
           "会社〒": postal_work, "会社電話1〜10": phone_work,
           "自宅住所1": Z(h_base), "自宅住所2": Z(h_bld),
           "自宅〒": postal_home, "自宅電話1〜10": phone_home,
           "その他住所1": Z(o_base), "その他住所2": Z(o_bld),
           "その他〒": postal_other, "その他電話1〜10": phone_other,
           "備考1": Z(notes), "誕生日": Z(birthday)}
    return row

# =========================
# CSV I/O と Flaskエンドポイント
# =========================

def read_google_csv(file_bytes):
    for enc in ("utf-8", "utf-8-sig", "cp932", "shift_jis"):
        try:
            f = io.StringIO(file_bytes.decode(enc))
            r = csv.DictReader(f)
            rows = [{(k or "").strip(): (v or "") for k, v in row.items()} for row in r]
            return rows
        except Exception:
            continue
    raise ValueError("CSV読み込みに失敗しました")

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        abort(400, "file is required")
    data = request.files["file"].read()
    grows = read_google_csv(data)
    out_rows = [build_atena_row(parse_google_row(r)) for r in grows]

    buf = io.StringIO(newline="")
    writer = csv.DictWriter(buf, fieldnames=list(out_rows[0].keys()))
    writer.writeheader()
    writer.writerows(out_rows)
    return send_file(io.BytesIO(buf.getvalue().encode("utf-8-sig")),
                     mimetype="text/csv",
                     as_attachment=True,
                     download_name="google_converted.csv")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port)
