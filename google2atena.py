# google2atena.py — v3.9.10a full-format no-pandas（統合・安定・即修正）
from flask import Flask, request, render_template_string, send_file
import csv
import io
import re
import unicodedata

app = Flask(__name__)

HTML = """
<!doctype html>
<title>Google連絡先CSV → 宛名職人CSV 変換（v3.9.10a full-format no-pandas）</title>
<h2>Google連絡先CSV → 宛名職人CSV 変換（v3.9.10a full-format no-pandas）</h2>
<form method=post enctype=multipart/form-data>
  <p><input type=file name=file required>
     <input type=submit value="変換開始">
</form>
"""

# ========= 共通ヘルパー =========

def zenkaku(s: str) -> str:
    """数字/英字/記号/スペースを全角へ。"""
    if not s:
        return ""
    fw = str.maketrans({
        " ": "　",
        "-": "－",
        "#": "＃",
        "/": "／",
        "&": "＆",
        "(": "（", ")":"）",
        "0": "０","1": "１","2": "２","3": "３","4": "４",
        "5": "５","6": "６","7": "７","8": "８","9": "９",
        "A":"Ａ","B":"Ｂ","C":"Ｃ","D":"Ｄ","E":"Ｅ","F":"Ｆ",
        "G":"Ｇ","H":"Ｈ","I":"Ｉ","J":"Ｊ","K":"Ｋ","L":"Ｌ",
        "M":"Ｍ","N":"Ｎ","O":"Ｏ","P":"Ｐ","Q":"Ｑ","R":"Ｒ",
        "S":"Ｓ","T":"Ｔ","U":"Ｕ","V":"Ｖ","W":"Ｗ","X":"Ｘ",
        "Y":"Ｙ","Z":"Ｚ",
        "a":"ａ","b":"ｂ","c":"ｃ","d":"ｄ","e":"ｅ","f":"ｆ",
        "g":"ｇ","h":"ｈ","i":"ｉ","j":"ｊ","k":"ｋ","l":"ｌ",
        "m":"ｍ","n":"ｎ","o":"ｏ","p":"ｐ","q":"ｑ","r":"ｒ",
        "s":"ｓ","t":"ｔ","u":"ｕ","v":"ｖ","w":"ｗ","x":"ｘ",
        "y":"ｙ","z":"ｚ",
    })
    return str(s).translate(fw)

def hankaku_ascii(s: str) -> str:
    """全角英数字を半角ASCIIへ（会社名かなの前処理用）。"""
    if not s:
        return ""
    return unicodedata.normalize("NFKC", s)

def hira_to_kata(s: str) -> str:
    """ひらがな→カタカナ。"""
    if not s:
        return ""
    out = []
    for ch in s:
        code = ord(ch)
        if 0x3041 <= code <= 0x3096:  # ぁ..ゖ
            out.append(chr(code + 0x60))
        else:
            out.append(ch)
    return "".join(out)

# ========= 電話番号 =========

MOBILE_PREFIX = {"070","080","090"}

# ★ここを修正：明示的に set にする（辞書と混在させない）
AREA_PREFIXES = set([
    # 2桁特例（03/06 は別処理）
    # 3桁主要（抜粋・拡張可）
    "011","015","017","018","019","022","023","024","025","026","027","028","029",
    "042","043","044","045","046","047","048","049",
    "052","053","054","055","056","058","059",
    "072","073","075","076","077","078","079",
    "082","083","084","086","087","088","089",
    "092","093","095","096","097","098",
])

def normalize_phone(raw):
    if not raw:
        return ""
    num = re.sub(r"\D", "", str(raw))
    if not num:
        return ""
    # 9桁で先頭0なし → 0付与（112615331 → 0112615331）
    if len(num) == 9 and not num.startswith("0"):
        num = "0" + num

    # 携帯
    if len(num) == 11 and num[:3] in MOBILE_PREFIX:
        return f"{num[:3]}-{num[3:7]}-{num[7:]}"
    # 03/06（10桁）
    if len(num) == 10 and num[:2] in {"03","06"}:
        return f"{num[:2]}-{num[2:6]}-{num[6:]}"
    # 既知の3桁市外局番（10桁）
    if len(num) == 10:
        if num[:3] in AREA_PREFIXES:
            return f"{num[:3]}-{num[3:6]}-{num[6:]}"
    # 11桁の固定電話（稀）→ 3-4-4
    if len(num) == 11 and num.startswith("0"):
        return f"{num[:3]}-{num[3:7]}-{num[7:]}"
    # フォールバック 3-3-4
    if len(num) == 10:
        return f"{num[:3]}-{num[3:6]}-{num[6:]}"
    return num

# ========= 住所分割 =========

def split_address_fields(row: dict):
    """
    Googleの各列から会社住所1/2/3を生成。
    - 会社住所1: 都道府県(Region) + 市区町村(City) + 通り(Streetの通り部分)
    - 会社住所2: 建物・号室（Streetの後半 + Extended）
    """
    region = (row.get("Address 1 - Region", "") or "").strip()
    city = (row.get("Address 1 - City", "") or "").strip()
    street = (row.get("Address 1 - Street", "") or "").strip()
    ext = (row.get("Address 1 - Extended Address", "") or "").strip()
    formatted = (row.get("Address 1 - Formatted", "") or "").strip()

    if not street and formatted:
        lines = [p.strip() for p in re.split(r"[\r\n]+", formatted) if p.strip()]
        if lines:
            street = lines[0]

     # Street から通り/建物を分離（空白で前半=通り、後半=建物等）
    bld = ""
    if street:
        parts = re.split(r"[ 　]+", street)
        if len(parts) >= 2:
            street_core = parts[0]
            bld = "　".join(parts[1:])
        else:
            street_core = street
    else:
        street_core = ""

    if ext:
        bld = (bld + ("　" if bld else "") + ext).strip()

    addr1 = f"{region}{city}{street_core}".strip()
    addr2 = bld
    addr3 = ""

    addr1 = zenkaku(addr1)
    addr2 = zenkaku(addr2)

    return addr1, addr2, addr3

# ========= メモ/Notes =========

def normalize_label(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip()
    table = str.maketrans({
        "①":"1","②":"2","③":"3","④":"4","⑤":"5",
        "１":"1","２":"2","３":"3","４":"4","５":"5",
        "　":"", " ":""
    })
    s = s.translate(table)
    s = s.lower().replace("memo","メモ")
    return s  # "メモ1" 等

def collect_memo_and_notes(row: dict):
    memos = {f"メモ{i}": "" for i in range(1,6)}
    notes = row.get("Notes","") or ""

    # Relation 系（優先）
    for i in range(1,6):
        lbl = normalize_label(row.get(f"Relation {i} - Label",""))
        val = (row.get(f"Relation {i} - Value","") or "").strip()
        if val and lbl == f"メモ{i}":
            memos[f"メモ{i}"] = val

    # 素カラムのメモ（フォールバック）
    for k, v in row.items():
        if not v:
            continue
        lbl = normalize_label(k)
        m = re.fullmatch(r"メモ([1-5])", lbl)
        if m and not memos[f"メモ{m.group(1)}"]:
            memos[f"メモ{m.group(1)}"] = str(v).strip()

    return memos, notes

# ========= 会社名かな =========

ROMA2KATA = {
    "A":"エー","B":"ビー","C":"シー","D":"ディー","E":"イー","F":"エフ",
    "G":"ジー","H":"エイチ","I":"アイ","J":"ジェー","K":"ケー","L":"エル",
    "M":"エム","N":"エヌ","O":"オー","P":"ピー","Q":"キュー","R":"アール",
    "S":"エス","T":"ティー","U":"ユー","V":"ブイ","W":"ダブリュー",
    "X":"エックス","Y":"ワイ","Z":"ゼット"
}

COMPANY_EXCEPT = {
    "札幌厚生病院":"サッポロコウセイビョウイン",
    "湘南東部総合病院":"ショウナントウブソウゴウビョウイン",
    "東京慈恵会医科大学":"トウキョウジケイカイイカダイガク",
}

LEGAL_FORMS = r"(株式会社|有限会社|合同会社|社団法人|財団法人|医療法人|学校法人)"

def company_kana(name: str) -> str:
    if not name:
        return ""
    s = str(name).strip()
    s = re.sub(LEGAL_FORMS, "", s).strip()

    if s in COMPANY_EXCEPT:
        return COMPANY_EXCEPT[s]

    s_ascii = hankaku_ascii(s)      # ＤＹ → DY
    s_ascii = hira_to_kata(s_ascii) # ひら→カナ

    out = []
    for ch in s_ascii:
        if re.match(r"[A-Za-z]", ch):
            out.append(ROMA2KATA[ch.upper()])
        else:
            out.append(ch)
    return "".join(out)

# ========= 出力ヘッダー =========

HEADER = [
    "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称","ニックネーム","旧姓","宛先",
    "自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
    "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social",
    "その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
    "会社名かな","会社名","部署名1","部署名2","役職名","連名","連名ふりがな","連名敬称","連名誕生日",
    "メモ1","メモ2","メモ3","メモ4","メモ5",
    "備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
]

# ========= 行変換 =========

def convert_row(r: dict) -> dict:
    first = r.get("First Name","") or ""
    last = r.get("Last Name","") or ""
    first_kana = r.get("Phonetic First Name","") or ""
    last_kana  = r.get("Phonetic Last Name","") or ""
    nick = r.get("Nickname","") or ""
    org = r.get("Organization Name","") or ""
    dept = r.get("Organization Department","") or ""
    title = r.get("Organization Title","") or ""
    birthday = r.get("Birthday","") or ""

    zip_code = (r.get("Address 1 - Postal Code","") or "").strip()

    addr1, addr2, addr3 = split_address_fields(r)

    phones = []
    for i in range(1,6):
        v = r.get(f"Phone {i} - Value","")
        if v:
            p = normalize_phone(v)
            if p:
                phones.append(p)
    seen = set()
    uniq = []
    for p in phones:
        if p not in seen:
            uniq.append(p); seen.add(p)
    phone_str = ";".join(uniq)

    emails = []
    for i in range(1,6):
        v = r.get(f"E-mail {i} - Value","")
        if v:
            emails.append(v.strip())
    email_str = ";".join(emails)

    memos, note = collect_memo_and_notes(r)

    row = {h:"" for h in HEADER}
    row.update({
        "姓": last,
        "名": first,
        "姓かな": last_kana,
        "名かな": first_kana,
        "姓名": f"{last}　{first}".strip(),
        "姓名かな": f"{last_kana}　{first_kana}".strip(),
        "敬称": "様",
        "ニックネーム": nick,
        "宛先": "会社",

        "会社〒": zip_code,
        "会社住所1": addr1,
        "会社住所2": addr2,
        "会社住所3": addr3,

        "会社電話": phone_str,
        "会社E-mail": email_str,

        "会社名かな": company_kana(org),
        "会社名": org,

        "部署名1": zenkaku(dept),
        "部署名2": "",
        "役職名": zenkaku(title),

        "メモ1": memos["メモ1"],
        "メモ2": memos["メモ2"],
        "メモ3": memos["メモ3"],
        "メモ4": memos["メモ4"],
        "メモ5": memos["メモ5"],

        "備考1": note,
        "誕生日": birthday,
    })
    return row

def convert_google_to_atena(text: str):
    reader = csv.DictReader(io.StringIO(text))
    return [convert_row(r) for r in reader]

# ========= Flask =========

@app.route("/", methods=["GET","POST"])
def root():
    if request.method == "POST":
        f = request.files.get("file")
        if not f:
            return render_template_string(HTML)
        text = f.read().decode("utf-8", errors="ignore")
        rows = convert_google_to_atena(text)

        buf = io.StringIO()
        writer = csv.DictWriter(buf, fieldnames=HEADER, extrasaction="ignore")
        writer.writeheader()
        for r in rows:
            writer.writerow(r)
        data = buf.getvalue().encode("utf-8-sig")
        return send_file(io.BytesIO(data),
                         as_attachment=True,
                         download_name="google_converted.csv",
                         mimetype="text/csv")
    return render_template_string(HTML)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
