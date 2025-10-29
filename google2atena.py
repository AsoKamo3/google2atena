# google2atena.py  v3.9.18r3  (Render最終安定版 / no-pandas)
# - 文字コード自動判定（chardet）
# - 区切り自動検出（csv.Sniffer）
# - キー正規化
# - 住所整形：郵便番号=半角ハイフン、住所本文=全角（数字/英字/記号）、建物分離（先頭スペース/空白でスプリット）
# - 電話整形：全国市外局番プリセット + 050/0120/0800/0570、;連結、重複排除
# - メール整形：;連結、重複排除、Home/Work/Otherへ
# - メモ抽出：Relation「メモ1〜5」、Notes→備考
# - 会社かな：外部辞書（company_dicts / kanji_word_map / corp_terms）があれば使用。無ければ最小デフォルトでフェールセーフ。
# - 出力：CSV (UTF-8-SIG) ダウンロード

from flask import Flask, request, render_template_string, send_file, abort
import csv
import io
import re
from collections import OrderedDict
from datetime import datetime

# ---- optional dictionaries (fail-safe fallback) ----
try:
    from company_dicts import COMPANY_EXCEPT  # dict
except Exception:
    COMPANY_EXCEPT = {
        "ＮＨＫエデュケーショナル": "エヌエイチケーエデュケーショナル",
        "博報堂ＤＹメディアパートナーズ": "ハクホウドウディーワイメディアパートナーズ",
    }

try:
    from kanji_word_map import KANJI_WORD_MAP  # dict
except Exception:
    KANJI_WORD_MAP = {
        "日本放送協会": "ニッポンホウソウキョウカイ",
        "日本": "ニホン",
        "出版": "シュッパン",
        "新聞": "シンブン",
        "社": "シャ",
        "大学": "ダイガク",
        "病院": "ビョウイン",
        "東京": "トウキョウ",
        "札幌": "サッポロ",
        "慈恵会医科": "ジケイイカ",
        "厚生": "コウセイ",
        "文庫": "ブンコ",
        "テレビ": "テレビ",
        # 英字→読み（足りない分は COMPANY_EXCEPT で個別化）
        "ＮＨＫ": "エヌエイチケー",
        "ＤＹ": "ディーワイ",
    }

try:
    from corp_terms import CORP_TERMS  # list[str]
except Exception:
    CORP_TERMS = [
        "株式会社", "有限会社", "合同会社", "合資会社", "相互会社",
        "一般社団法人", "一般財団法人", "公益社団法人", "公益財団法人",
        "特定非営利活動法人", "ＮＰＯ法人", "学校法人", "医療法人",
        "宗教法人", "社会福祉法人", "公立大学法人", "独立行政法人",
        "地方独立行政法人", "国立研究開発法人", "公益法人", "協同組合",
        "協業組合", "信用金庫", "信用組合", "農業協同組合", "漁業協同組合",
        "労働金庫", "商工会議所", "商工会", "互助会", "同友会", "後援会",
        "振興会", "委員会", "連盟", "協会", "財団", "社団",
    ]

# ---- 市外局番＆特番帯（全国プリセット：要約版） ----
try:
    from jp_area_codes import AREA_CODES
except Exception:
    # 要約プリセット（必要帯域＋代表例のみ。実運用は jp_area_codes.py を置くことを推奨）
    AREA_CODES = set((
        # 3桁
        "011","013","014","015","017","018","019","022","023","024","025","026","027","028","029",
        "03","04","042","043","044","045","046","047","048","049",
        "052","053","054","055","058","059",
        "06","072","073","075","076","077","078","079",
        "082","083","084","086","087","088",
        "092","093","094","095","096","097","098",
        # 4桁 代表例（全網羅は別ファイルで）
        "0134","0138","0154","0178","0185","0195","0225","0237","0242","0254","0263","0276","0287","0299",
        "0422","0428","0436","0438","0448","0465","0475","0479","0480","0493","0495",
        "0532","0533","0544","0551","0555","0586","0594","0595",
        "0725","0735","0758","0761","0766","0770","0772","0774","0776","0789","0796","0797","0798",
        "0823","0824","0833","0847","0863","0875","0880","0884","0887",
        "0923","0930","0940","0955","0965","0972","0973","0982","0985","0987",
    ))

SPECIAL_PREFIX = ("0120","0800","0570","050")  # フリーダイヤル・ナビ・VoIP

# ---- chardet（文字コード自動判定） ----
try:
    import chardet
except Exception:
    chardet = None  # 無くてもUTF-8前提で動作

app = Flask(__name__)

# -------------------- ユーティリティ --------------------

def strip_bom(s: str) -> str:
    return s.lstrip("\ufeff").strip()

def normalize_key(k: str) -> str:
    k = strip_bom(k)
    k = re.sub(r"\s+", " ", k)
    return k

def detect_encoding(b: bytes) -> str:
    if chardet:
        res = chardet.detect(b) or {}
        enc = (res.get("encoding") or "utf-8")
        return enc
    return "utf-8"

# 全角/半角変換（住所本文は全角、郵便/電話は半角）
FW_TABLE = str.maketrans(
    {**{chr(ord('0')+i): chr(ord('０')+i) for i in range(10)},
     **{chr(ord('a')+i): chr(ord('ａ')+i) for i in range(26)},
     **{chr(ord('A')+i): chr(ord('Ａ')+i) for i in range(26)},
     **{
         '-':'－', '#':'＃', '/':'／', ' ':'　', ',':'，', '.':'．',
         ':':'：', ';':'；', '&':'＆', '@':'＠', '(': '（', ')':'）'
     }}
)

def to_zenkaku_address(s: str) -> str:
    if not s:
        return ""
    return s.translate(FW_TABLE)

def normalize_postal(s: str) -> str:
    if not s:
        return ""
    digits = re.sub(r"\D", "", s)
    if len(digits) == 7:
        return f"{digits[0:3]}-{digits[3:7]}"
    return ""

def split_building(addr_line: str) -> tuple[str, str]:
    """
    住所本文→建物分離：'最初に出てくる空白で二分割' を第一優先。
    失敗時、よくある建物語を拾って分割。
    """
    if not addr_line:
        return "", ""
    s = addr_line.strip()
    s = re.sub(r"[ \t]+", " ", s.replace("　", " "))
    # 1) 最初の空白で分割
    m = re.search(r" ", s)
    if m:
        i = m.start()
        return s[:i], s[i+1:]
    # 2) よくある建物ワード
    m2 = re.search(r"(ビル|マンション|ハイツ|アパート|タワー|レジデンス|コーポ|団地)", s)
    if m2:
        i = m2.start()
        return s[:i], s[i:]
    return s, ""

def cleanup_emails(raw_list) -> str:
    vals = []
    for x in raw_list:
        if not x:
            continue
        # ::: 区切りも分割
        for y in re.split(r"\s*[:;|／／,、]|:::+\s*", x):
            y = y.strip()
            if y and "@" in y:
                vals.append(y)
    # 重複排除・順序維持
    seen = set()
    out = []
    for v in vals:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return ";".join(out)

def hyphenate_phone_by_area(digits: str) -> str:
    """
    与えられた数字列（先頭0含む）を、AREA_CODES/SPECIAL_PREFIX から最長一致でハイフン挿入。
    """
    if not digits or not digits.startswith("0"):
        return digits
    # 特番帯
    for sp in SPECIAL_PREFIX:
        if digits.startswith(sp):
            # 0120-xxx-xxx / 0800-xxx-xxx / 0570-xxx-xxx / 050-xxxx-xxxx
            if sp in ("0120","0800","0570"):
                rest = digits[len(sp):]
                if len(rest) in (6,7):
                    return f"{sp}-{rest[:-3]}-{rest[-3:]}"
                # デフォルト 3-3
                return f"{sp}-{rest[:3]}-{rest[3:]}" if len(rest) > 3 else digits
            if sp == "050":
                rest = digits[3:]
                if len(rest) == 8:
                    return f"050-{rest[:4]}-{rest[4:]}"
                return f"050-{rest}"
    # 市外局番（最長一致）
    # 5桁候補→4→3→2
    for k in sorted(AREA_CODES, key=len, reverse=True):
        if digits.startswith(k):
            local = digits[len(k):]
            # 標準  k - (3 or 4) - 4
            if len(local) >= 7:
                return f"{k}-{local[:-4]}-{local[-4:]}"
            elif len(local) >= 5:
                return f"{k}-{local[:-4]}-{local[-4:]}"
            elif len(local) >= 4:
                return f"{k}-{local[:-4]}-{local[-4:]}"
            else:
                return f"{k}-{local}"
    # 携帯(PHS含む) 070/080/090 は 3-4-4
    if digits.startswith(("070","080","090")) and len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    # デフォルト：後方4桁で分割
    if len(digits) > 4:
        return f"{digits[:-4]}-{digits[-4:]}"
    return digits

def normalize_phones(raw_list) -> str:
    """
    入力の様々な区切りを受け付け、数字抽出→整形→;で連結。
    """
    tokens = []
    for x in raw_list:
        if not x:
            continue
        # :::やスペースなどで分割
        parts = re.split(r"\s*(?:::+|;|,|、|\||／|/|\s)\s*", x)
        for p in parts:
            p = p.strip()
            if not p:
                continue
            d = re.sub(r"\D", "", p)
            # 先頭0が抜けた国番号形式「81...」を 0 始まりに戻す（簡易）
            if d.startswith("81") and len(d) >= 10 and not d.startswith("810"):
                d = "0" + d[2:]
            # 妥当長のみ採用（6桁以上を電話候補とする）
            if len(d) >= 6:
                tokens.append(d)

    out = []
    seen = set()
    for d in tokens:
        f = hyphenate_phone_by_area(d)
        if f not in seen:
            seen.add(f)
            out.append(f)
    return ";".join(out)

def zenkaku_clean_company_kana(s: str) -> str:
    # 会社かなは 全角のみ・「・ . ，」除去
    if not s:
        return ""
    s = re.sub(r"[・\.\,，､]", "", s)
    return to_zenkaku_address(s)

def remove_corp_terms(name: str) -> str:
    s = name or ""
    s = s.strip()
    for t in sorted(CORP_TERMS, key=len, reverse=True):
        s = s.replace(t, "")
    # 全角/半角混在スペース除去
    s = re.sub(r"\s+", "", s.replace("　", ""))
    return s

def company_to_kana(name: str) -> str:
    if not name:
        return ""
    base = name.strip()
    # 1) 例外表
    for k, v in COMPANY_EXCEPT.items():
        if k in base:
            return zenkaku_clean_company_kana(v)
    # 2) 法人格を除去
    s = remove_corp_terms(base)
    # 3) KANJI_WORD_MAP を後方一致優先で置換
    if KANJI_WORD_MAP:
        for k in sorted(KANJI_WORD_MAP.keys(), key=len, reverse=True):
            s = s.replace(k, KANJI_WORD_MAP[k])
    # 4) 英字を読み上げていない場合の粗救済（大文字英字→読みを伸ばすのは辞書側で管理）
    #    ここでは単純に全角化＋記号除去のみにとどめる
    s = zenkaku_clean_company_kana(s)
    return s

# -------------------- 入出力カラム --------------------

HEADERS = [
    "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称","ニックネーム",
    "旧姓","宛先","自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
    "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social",
    "その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
    "会社名かな","会社名","部署名1","部署名2","役職名","連名","連名ふりがな","連名敬称","連名誕生日",
    "メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
]

# -------------------- 行変換 --------------------

def pick(row, *keys):
    for k in keys:
        if k in row and row[k]:
            return row[k]
    return ""

def collect_memos(row):
    memos = ["","","","",""]
    # Relation i - Label / Value
    for i in range(1, 10):
        lbl = row.get(f"Relation {i} - Label","").strip()
        val = row.get(f"Relation {i} - Value","").strip()
        if not lbl and not val:
            continue
        m = re.match(r"メモ\s*(\d+)", lbl)
        if m:
            idx = int(m.group(1))
            if 1 <= idx <= 5:
                memos[idx-1] = val
    return memos

def split_google_address(row):
    """
    Google CSVの Address 1 を Work/Home/Other いずれにも対応
    - 郵便番号は半角 123-4567
    - 住所本文は全角化
    - 建物分離は Address 1 - Street を先頭空白で分割（fallback で建物語）
    - city/region 等も連結して住所1へ
    """
    label = row.get("Address 1 - Label","").strip()
    street = (row.get("Address 1 - Street","") or "").replace("\r","").replace("\n"," ")
    city = row.get("Address 1 - City","") or ""
    region = row.get("Address 1 - Region","") or ""
    postal = row.get("Address 1 - Postal Code","") or ""
    # 国/拡張は無視（日本前提）
    # 住所1（市区町村＋番地等）を組み立て（半角→全角）
    line = " ".join([region, city, street]).strip()
    # 半角→全角
    line = to_zenkaku_address(line)
    # 建物分離
    addr2, bldg = split_building(line)
    # 郵便番号
    zip_ = normalize_postal(postal)
    # 住所3（建物/号室）
    return label, zip_, addr2, bldg

def assign_address_to_cols(label, zip_, addr2, bldg, slots):
    """
    slots: dict 出力フィールド辞書
    """
    # 住所本文は既に全角化されている前提
    if label.lower() == "home" or label == "自宅":
        slots["自宅〒"] = zip_
        slots["自宅住所1"] = addr2
        slots["自宅住所2"] = bldg
        slots["自宅住所3"] = ""
    elif label.lower() == "other" or label == "その他":
        slots["その他〒"] = zip_
        slots["その他住所1"] = addr2
        slots["その他住所2"] = bldg
        slots["その他住所3"] = ""
    else:
        # Work or default
        slots["会社〒"] = zip_
        slots["会社住所1"] = addr2
        slots["会社住所2"] = bldg
        slots["会社住所3"] = ""

def convert_row(row):
    out = OrderedDict((h,"") for h in HEADERS)

    # --- 基本名 ---
    last = pick(row, "Last Name","姓")
    first = pick(row, "First Name","名")
    last_k = pick(row, "Phonetic Last Name","姓（かな）","姓かな")
    first_k = pick(row, "Phonetic First Name","名（かな）","名かな")
    middle = pick(row, "Middle Name","ミドルネーム")
    middle_k = pick(row, "Phonetic Middle Name","ミドルネームかな")
    nick = pick(row, "Nickname","ニックネーム")
    honor = "様"
    full = f"{last}　{first}".strip()
    full_k = f"{last_k}　{first_k}".strip()

    out["姓"] = last
    out["名"] = first
    out["姓かな"] = last_k
    out["名かな"] = first_k
    out["姓名"] = full
    out["姓名かな"] = full_k
    out["ミドルネーム"] = middle
    out["ミドルネームかな"] = middle_k
    out["敬称"] = honor
    out["ニックネーム"] = nick
    out["旧姓"] = ""
    out["宛先"] = "会社"

    # --- 会社/部署/役職 ---
    org = pick(row, "Organization Name","会社名")
    dept = pick(row, "Organization Department","部署")
    title = pick(row, "Organization Title","役職")

    out["会社名"] = org
    out["部署名1"] = dept
    out["部署名2"] = ""
    out["役職名"] = title

    # 会社かな
    out["会社名かな"] = company_to_kana(org)

    # --- メール ---
    home_emails = []
    work_emails = []
    other_emails = []
    for i in range(1, 10):
        el = row.get(f"E-mail {i} - Label","") or ""
        ev = row.get(f"E-mail {i} - Value","") or ""
        if not ev:
            continue
        if el.lower() == "home":
            home_emails.append(ev)
        elif el.lower() == "work":
            work_emails.append(ev)
        else:
            other_emails.append(ev)
    out["自宅E-mail"] = cleanup_emails(home_emails)
    out["会社E-mail"] = cleanup_emails(work_emails)
    out["その他E-mail"] = cleanup_emails(other_emails)

    # --- 電話 ---
    home_phones = []
    work_phones = []
    other_phones = []
    for i in range(1, 10):
        pl = row.get(f"Phone {i} - Label","") or ""
        pv = row.get(f"Phone {i} - Value","") or ""
        if not pv:
            continue
        if pl.lower() == "home":
            home_phones.append(pv)
        elif pl.lower() == "work":
            work_phones.append(pv)
        elif pl.lower() == "mobile":
            # 方針：モバイルは会社電話に寄せるのか/自宅に寄せるのかはサンプルに倣い会社側へ併記
            work_phones.append(pv)
        else:
            other_phones.append(pv)

    out["自宅電話"] = normalize_phones(home_phones)  # 半角ハイフン
    out["会社電話"] = normalize_phones(work_phones)
    out["その他電話"] = normalize_phones(other_phones)

    # --- 住所（Address 1 のみ対応 / Google側は通常1枠）---
    label, zip_, addr2, bldg = split_google_address(row)
    assign_address_to_cols(label, zip_, addr2, bldg, out)

    # --- IM/URL/Social（ダミー：入力に無ければ空）---
    out["自宅IM ID"] = ""
    out["自宅URL"] = ""
    out["自宅Social"] = ""
    out["会社IM ID"] = ""
    out["会社URL"] = ""
    out["会社Social"] = ""
    out["その他IM ID"] = ""
    out["その他URL"] = ""
    out["その他Social"] = ""

    # --- 連名系 ---
    out["連名"] = ""
    out["連名ふりがな"] = ""
    out["連名敬称"] = ""
    out["連名誕生日"] = ""

    # --- メモ ---
    m1,m2,m3,m4,m5 = collect_memos(row)
    out["メモ1"] = m1
    out["メモ2"] = m2
    out["メモ3"] = m3
    out["メモ4"] = m4
    out["メモ5"] = m5

    # 備考：Notesを第一備考へ
    notes = pick(row, "Notes","メモ","備考")
    out["備考1"] = notes
    out["備考2"] = ""
    out["備考3"] = ""

    # --- 誕生日 ---
    bday = pick(row, "Birthday","誕生日")
    # 受け取り形式に関わらず y/m/d へ（失敗時は原文）
    bd_norm = ""
    if bday:
        for fmt in ("%Y-%m-%d","%Y/%m/%d","%Y.%m.%d","%Y%m%d","%m/%d/%Y"):
            try:
                dt = datetime.strptime(bday.strip(), fmt)
                bd_norm = dt.strftime("%Y/%m/%d")
                break
            except Exception:
                continue
        if not bd_norm:
            bd_norm = bday.strip()
    out["誕生日"] = bd_norm

    # --- ダミー属性 ---
    out["性別"] = "選択なし"
    out["血液型"] = "選択なし"
    out["趣味"] = ""
    out["性格"] = ""

    return out

def convert_text_to_rows(text: str):
    # 区切り自動検出
    sniffer = csv.Sniffer()
    sample = text[:4096]
    try:
        dialect = sniffer.sniff(sample)
        delimiter = dialect.delimiter
    except Exception:
        # Google/Atena 標準はカンマ
        delimiter = ","

    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)
    # キー正規化
    reader.fieldnames = [normalize_key(h) for h in reader.fieldnames or []]

    rows = []
    for r in reader:
        norm = {normalize_key(k): (v or "") for k, v in r.items()}
        rows.append(convert_row(norm))
    return rows

# -------------------- Web --------------------

INDEX_HTML = """<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>google2atena v3.9.18r3</title>
<style>
body{font-family:system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Hiragino Kaku Gothic ProN","Noto Sans JP",sans-serif;margin:24px;}
.card{max-width:900px;margin:0 auto;border:1px solid #ddd;border-radius:12px;padding:20px;box-shadow:0 2px 8px rgba(0,0,0,.04);}
h1{font-size:20px;margin:0 0 16px;}
.btn{display:inline-block;padding:10px 16px;border-radius:8px;border:1px solid #333;cursor:pointer;background:#111;color:#fff}
.note{color:#555;font-size:13px;margin-top:8px}
.status{margin-top:12px;font-size:13px}
.error{color:#c00}
.success{color:#060}
</style>
</head>
<body>
<div class="card">
  <h1>Google 連絡先 → Atena CSV 変換（v3.9.18r3）</h1>
  <form id="f" action="/convert" method="post" enctype="multipart/form-data">
    <input type="file" name="file" accept=".csv" required />
    <button class="btn" type="submit">変換開始</button>
  </form>
  <div class="note">※ 文字コードは自動判定、区切り（, / ; / タブ等）も自動検出します。</div>
  <div id="status" class="status"></div>
</div>
<script>
document.getElementById('f').addEventListener('submit', function(){
  document.getElementById('status').innerHTML = '変換中…';
});
</script>
</body>
</html>
"""

@app.route("/", methods=["GET","POST"])
@app.route("/convert", methods=["POST"])  # ← /convert POST を許可（Render ログの 404 対策）
def index():
    if request.method == "GET":
        return render_template_string(INDEX_HTML)

    # POST: ファイル受領
    if "file" not in request.files:
        return render_template_string(INDEX_HTML + '<p class="error">⚠️ エラーが発生しました。CSVファイルを添付してください。</p>'), 400
    file = request.files["file"]
    b = file.read()
    if not b:
        return render_template_string(INDEX_HTML + '<p class="error">⚠️ エラーが発生しました。CSVが空のようです。</p>'), 400

    enc = detect_encoding(b)
    try:
        text = b.decode(enc, errors="replace")
    except Exception:
        # 最終フォールバック
        text = b.decode("utf-8", errors="replace")

    try:
        rows = convert_text_to_rows(text)
        # CSV (UTF-8 with BOM)
        buf = io.StringIO()
        writer = csv.DictWriter(buf, fieldnames=HEADERS, lineterminator="\n")
        writer.writeheader()
        for r in rows:
            writer.writerow(r)
        data = buf.getvalue().encode("utf-8-sig")
        return send_file(
            io.BytesIO(data),
            mimetype="text/csv; charset=utf-8",
            as_attachment=True,
            download_name="atena.csv"
        )
    except Exception as e:
        # 例外詳細は返さず、UIメッセージ（ログは Render 側で確認）
        return render_template_string(INDEX_HTML + '<p class="error">⚠️ エラーが発生しました。CSVの形式や文字コードをご確認ください。</p>'), 400

if __name__ == "__main__":
    # Render は PORT を環境変数で指定。ローカル用デフォルト 10000。
    import os
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
