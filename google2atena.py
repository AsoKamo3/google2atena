# google2atena.py
# Google連絡先CSV → 宛名職人CSV 変換（v3.9.5 full-format no-pandas 最終安定版）

from flask import Flask, request, render_template_string, send_file
import csv
import io
import re
from collections import OrderedDict

app = Flask(__name__)

TITLE = "Google連絡先CSV → 宛名職人CSV 変換（v3.9.5 full-format no-pandas 最終安定版）"

HTML = f"""
<!doctype html>
<meta charset="utf-8">
<title>{TITLE}</title>
<h2>{TITLE}</h2>
<form method="post" enctype="multipart/form-data">
  <p><input type="file" name="file" accept=".csv">
     <input type="submit" value="変換開始">
</form>
<p style="font-size:12px;color:#666;">
  ・Google連絡先CSVに対応（BOMあり/なし）<br>
  ・「メモゆれ対応」「Notes→備考1」「住所全角化」「電話NTT整形」「会社名かな辞書対応」
</p>
"""

# =====================
# Utility Functions
# =====================

def coalesce(*vals):
    for v in vals:
        if v and str(v).strip():
            return str(v).strip()
    return ""

# 半角→全角（住所用）
def to_fullwidth_address(s):
    if not s:
        return ""
    s = str(s)
    fw = str.maketrans({chr(i): chr(i + 0xFEE0) for i in range(0x21, 0x7F)})
    s = s.translate(fw)
    s = s.replace(" ", "　").replace("-", "－").replace("#", "＃")
    return s

# 郵便番号を半角化（XXX-XXXX）
def to_halfwidth_postal(s):
    if not s:
        return ""
    digits = re.sub(r"\D", "", str(s))
    if len(digits) == 7:
        return f"{digits[:3]}-{digits[3:]}"
    return digits

# 電話番号をNTT形式に
def format_jp_phone(num):
    if not num:
        return ""
    digits = re.sub(r"\D", "", str(num))
    if digits and digits[0] != "0" and len(digits) in (9, 10):
        digits = "0" + digits
    if len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    if len(digits) == 10:
        if digits.startswith(("03", "06")):
            return f"{digits[:2]}-{digits[2:6]}-{digits[6:]}"
        else:
            return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return digits

# =====================
# Address Parsing
# =====================

def parse_formatted_address(row):
    postal_raw = coalesce(row.get("Address 1 - Postal Code"))
    region = coalesce(row.get("Address 1 - Region"))
    city = coalesce(row.get("Address 1 - City"))
    street = coalesce(row.get("Address 1 - Street"))

    # fallback
    if not (region and city and street):
        formatted = coalesce(row.get("Address 1 - Formatted"))
        if formatted:
            lines = [l.strip() for l in re.split(r"[\r\n]+", formatted) if l.strip()]
            if len(lines) >= 1 and not street:
                street = lines[0]
            if len(lines) >= 2 and not city:
                city = lines[1]
            if len(lines) >= 3 and not region:
                if re.search(r"[都道府県]$", lines[2]):
                    region = lines[2]
            if len(lines) >= 4 and not postal_raw:
                postal_raw = lines[3]

    postal = to_halfwidth_postal(postal_raw)

    addr_num, bldg = street, ""
    if street and re.search(r"[ 　]", street):
        idx = re.search(r"[ 　]", street).start()
        addr_num = street[:idx].strip()
        bldg = street[idx:].strip()

    addr1 = to_fullwidth_address(region + city + addr_num)
    addr2 = to_fullwidth_address(bldg)
    addr3 = ""
    return postal, addr1, addr2, addr3

# =====================
# Email & Phone
# =====================

def split_emails(value):
    if not value:
        return []
    s = str(value).replace(":::", ";").replace("；", ";").replace("，", ",")
    parts = re.split(r"[;,\s]+", s)
    return [p.strip().lower() for p in parts if "@" in p]

def collect_emails(row):
    emails = []
    for i in range(1, 8):
        v = row.get(f"E-mail {i} - Value", "")
        if v:
            emails += split_emails(v)
    return ";".join(dict.fromkeys(emails))

def label_rank(label):
    l = str(label).lower()
    if "work" in l: return 0
    if "mobile" in l: return 1
    if "home" in l: return 2
    return 3

def pick_company_phones(row):
    buckets = {0: [], 1: [], 2: [], 3: []}
    for i in range(1, 8):
        lbl = row.get(f"Phone {i} - Label", "")
        val = row.get(f"Phone {i} - Value", "")
        if not val: continue
        raw_list = re.split(r"::+|;|,|\s+", str(val))
        for raw in raw_list:
            fmt = format_jp_phone(raw.strip())
            if fmt:
                buckets[label_rank(lbl)].append(fmt)
    seen, out = set(), []
    for k in (0,1,2,3):
        for p in buckets[k]:
            if p not in seen:
                seen.add(p)
                out.append(p)
    return ";".join(out)

# =====================
# Memo & Notes Handling
# =====================

def collect_memos_and_notes(row):
    memos = {f"メモ{i}": "" for i in range(1, 6)}
    notes_out = coalesce(row.get("Notes", ""))

    for key, val in row.items():
        if not val: continue
        k = key.strip()
        m = re.search(r"(?:メモ|memo|ＭＥＭＯ)\s*([①②③④⑤１２３４５1-5])", k, re.IGNORECASE)
        if m:
            n = m.group(1)
            n = n.translate(str.maketrans("①②③④⑤１２３４５", "12345"))
            memos[f"メモ{n}"] = str(val)
    return memos, notes_out

# =====================
# Company Name Kana
# =====================

COMPANY_KANA_DICT = [
    ("ＮＨＫ", "エヌエイチケー"),
    ("日経", "ニッケイ"),
    ("博報堂", "ハクホウドウ"),
    ("講談社", "コウダンシャ"),
    ("東京", "トウキョウ"),
    ("湘南", "ショウナン"),
    ("夢眠社", "ユメミシャ"),
    ("ライズ＆プレイ", "ライズアンドプレイ"),
    ("河野文庫", "コウノブンコ"),
]

ENGLISH_KANA_MAP = {
    "BP": "ビーピー", "DY": "ディーワイ", "LAB": "ラボ", "WORKS": "ワークス",
    "INC": "インク", "CORP": "コープ", "CO": "シーオー", "LTD": "エルティーディー",
}

def to_company_kana(org_name):
    if not org_name:
        return ""
    s = to_fullwidth_address(org_name)
    for k, v in COMPANY_KANA_DICT:
        if k in s:
            s = s.replace(k, v)
    for k, v in ENGLISH_KANA_MAP.items():
        s = re.sub(rf"\b{k}\b", v, s, flags=re.IGNORECASE)
    return s

# =====================
# Row Conversion
# =====================

OUT_COLUMNS = [
    "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称","ニックネーム","旧姓","宛先",
    "自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
    "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social",
    "その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
    "会社名かな","会社名","部署名1","部署名2","役職名","連名","連名ふりがな","連名敬称","連名誕生日",
    "メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
]

def convert_row(row):
    first, last = row.get("First Name",""), row.get("Last Name","")
    first_kana, last_kana = row.get("Phonetic First Name",""), row.get("Phonetic Last Name","")
    middle, middle_kana = row.get("Middle Name",""), row.get("Phonetic Middle Name","")
    nickname, birthday = row.get("Nickname",""), row.get("Birthday","")
    org, dept, title = row.get("Organization Name",""), row.get("Organization Department",""), row.get("Organization Title","")

    postal, addr1, addr2, addr3 = parse_formatted_address(row)
    phone_str = pick_company_phones(row)
    email_str = collect_emails(row)
    memos, notes1 = collect_memos_and_notes(row)
    org_kana = to_company_kana(org)

    out = OrderedDict()
    out["姓"] = last
    out["名"] = first
    out["姓かな"] = last_kana
    out["名かな"] = first_kana
    out["姓名"] = f"{last}　{first}".strip()
    out["姓名かな"] = f"{last_kana}　{first_kana}".strip()
    out["ミドルネーム"] = middle
    out["ミドルネームかな"] = middle_kana
    out["敬称"] = "様"
    out["ニックネーム"] = nickname
    out["旧姓"] = ""
    out["宛先"] = "会社"

    # 自宅系は空欄
    for col in OUT_COLUMNS[12:21]:
        out[col] = ""

    # 会社
    out["会社〒"] = postal
    out["会社住所1"] = addr1
    out["会社住所2"] = addr2
    out["会社住所3"] = addr3
    out["会社電話"] = phone_str
    out["会社IM ID"] = ""
    out["会社E-mail"] = email_str
    out["会社URL"] = ""
    out["会社Social"] = ""

    for col in OUT_COLUMNS[30:39]:
        out[col] = ""

    out["会社名かな"] = org_kana
    out["会社名"] = org
    out["部署名1"] = dept
    out["部署名2"] = ""
    out["役職名"] = title
    out["連名"] = out["連名ふりがな"] = out["連名敬称"] = out["連名誕生日"] = ""

    for i in range(1,6):
        out[f"メモ{i}"] = memos[f"メモ{i}"]

    out["備考1"] = notes1
    out["備考2"] = out["備考3"] = ""
    out["誕生日"] = birthday
    out["性別"] = "選択なし"
    out["血液型"] = "選択なし"
    out["趣味"] = out["性格"] = ""
    return OrderedDict((k, out.get(k, "")) for k in OUT_COLUMNS)

def convert_google_to_atena(csv_text):
    reader = csv.DictReader(io.StringIO(csv_text))
    return [convert_row(r) for r in reader]

# =====================
# Flask Endpoint
# =====================

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        f = request.files.get("file")
        if not f: return render_template_string(HTML)
        raw = f.read()
        text = raw.decode("utf-8-sig", errors="replace")
        rows = convert_google_to_atena(text)
        buf = io.StringIO()
        writer = csv.DictWriter(buf, fieldnames=OUT_COLUMNS, lineterminator="\n")
        writer.writeheader()
        for r in rows:
            writer.writerow(r)
        data = buf.getvalue().encode("utf-8-sig")
        return send_file(io.BytesIO(data), as_attachment=True, download_name="google_converted.csv", mimetype="text/csv")
    return render_template_string(HTML)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
