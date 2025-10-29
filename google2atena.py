# google2atena.py  v3.9.18r7c  (Render安定版 / no-pandas)
# - 電話番号整形完全対応
# - ::: / , / / / ・ / ： 分割
# - 03補完 + 固定電話ハイフン精密処理
# - 他機能（住所振分・フェイルセーフ・メモ）は r7b と同一

import csv
import io
import re
import unicodedata
from flask import Flask, render_template_string, request, send_file

# ======== 外部辞書フェイルセーフ ========
try:
    from company_dicts import COMPANY_EXCEPT
except Exception:
    COMPANY_EXCEPT = {}

try:
    from kanji_word_map import KANJI_WORD_MAP
except Exception:
    KANJI_WORD_MAP = {}

try:
    from corp_terms import CORP_TERMS
except Exception:
    CORP_TERMS = [
        "株式会社", "有限会社", "合同会社", "合資会社", "相互会社",
        "一般社団法人", "一般財団法人", "公益社団法人", "公益財団法人",
        "特定非営利活動法人", "ＮＰＯ法人", "学校法人", "医療法人",
        "宗教法人", "社会福祉法人", "公立大学法人", "独立行政法人", "地方独立行政法人"
    ]

app = Flask(__name__)

# ======== 住所関連ユーティリティ ========

def to_zenkaku_for_address(s: str) -> str:
    if not s:
        return ""
    z = []
    for ch in s:
        code = ord(ch)
        if 0x21 <= code <= 0x7E:
            if ch == '-':
                z.append('－')
            else:
                z.append(unicodedata.normalize('NFKC', ch))
        else:
            z.append(ch)
    return "".join(z)

def format_postal(postal: str) -> str:
    if not postal:
        return ""
    digits = re.sub(r'\D', '', postal)
    if len(digits) == 7:
        return f"{digits[:3]}-{digits[3:]}"
    return postal

_SPLIT_RE = re.compile(r'[ \u3000]')

def split_first_space(addr_full: str):
    if not addr_full:
        return ("", "")
    m = _SPLIT_RE.search(addr_full)
    if not m:
        return (addr_full, "")
    i = m.start()
    return (addr_full[:i], addr_full[i+1:].strip())

def build_addr12(region, city, street):
    parts = [p for p in [region, city, street] if p]
    full = "".join(parts)
    full_z = to_zenkaku_for_address(full)
    a1, a2 = split_first_space(full_z)
    return (a1, a2)

def route_address_by_label(row, out):
    label = (row.get('Address 1 - Label') or "").strip().lower()
    region = row.get('Address 1 - Region') or ""
    city   = row.get('Address 1 - City') or ""
    street = row.get('Address 1 - Street') or ""
    postal = row.get('Address 1 - Postal Code') or ""

    jp_postal = format_postal(postal)
    addr1, addr2 = build_addr12(region, city, street)

    if label == 'home':
        out['自宅〒']    = jp_postal
        out['自宅住所1'] = addr1
        out['自宅住所2'] = addr2
        out['自宅住所3'] = ""
    elif label == 'other':
        out['その他〒']    = jp_postal
        out['その他住所1'] = addr1
        out['その他住所2'] = addr2
        out['その他住所3'] = ""
    else:
        out['会社〒']    = jp_postal
        out['会社住所1'] = addr1
        out['会社住所2'] = addr2
        out['会社住所3'] = ""

# ======== 電話番号整形（完全版） ========

CITY_CODES = [
    '099','098','097','096','095','094','093','092','089','088','087','086','085','084','083','082','079','078','077','076',
    '075','074','073','072','06','059','058','057','056','055','054','053','052','049','048','047','046','045','044','043','042',
    '04','03','029','028','027','026','025','024','023','022','019','018','017','015','011'
]

def split_multi_numbers(raw_value):
    # ::: , : ： ・ ／ / 空白などで分割
    return re.split(r'[:：;,／／・/]+', raw_value)

def normalize_phones(phone_values):
    all_numbers = []
    for val in phone_values:
        if not val:
            continue
        parts = split_multi_numbers(val)
        for p in parts:
            p = p.strip()
            if not p:
                continue
            nums = re.sub(r'\D', '', p)
            if not nums:
                continue
            if not nums.startswith('0'):
                nums = '0' + nums

            # 市外局番が抜けた9桁番号 (例: 364419772 → 03-6441-9772)
            if len(nums) == 9 and re.match(r'^[3-9]\d{8}$', nums):
                nums = '03' + nums

            formatted = nums

            # 各種形式
            if re.match(r'^0(70|80|90)\d{8}$', nums):
                formatted = f"{nums[:3]}-{nums[3:7]}-{nums[7:]}"
            elif re.match(r'^050\d{8}$', nums):
                formatted = f"{nums[:3]}-{nums[3:7]}-{nums[7:]}"
            elif re.match(r'^(0120|0800|0570)\d{6}$', nums):
                formatted = f"{nums[:4]}-{nums[4:7]}-{nums[7:]}"
            elif len(nums) == 10:
                for code in sorted(CITY_CODES, key=len, reverse=True):
                    if nums.startswith(code):
                        remain = nums[len(code):]
                        if len(remain) == 7:
                            formatted = f"{code}-{remain[:3]}-{remain[3:]}"
                        elif len(remain) == 6:
                            formatted = f"{code}-{remain[:2]}-{remain[2:]}"
                        break
            elif len(nums) == 9:
                formatted = f"{nums[:2]}-{nums[2:5]}-{nums[5:]}"
            else:
                formatted = nums

            if formatted not in all_numbers:
                all_numbers.append(formatted)
    return ";".join(all_numbers)

# ======== メール・メモ ========

def normalize_emails(email_values):
    emails = []
    for val in email_values:
        if not val:
            continue
        for part in re.split(r'[:：;,／／・/ ]+', val):
            part = part.strip()
            if part and part not in emails:
                emails.append(part)
    return ";".join(emails)

def extract_memos(row):
    memos = []
    for i in range(1, 6):
        label = row.get(f"Relation {i} - Label", "")
        value = row.get(f"Relation {i} - Value", "")
        if label and "メモ" in label and value:
            memos.append(value)
    notes = row.get("Notes", "")
    if notes:
        memos.append(notes)
    return memos

def kana_company_name(name):
    if not name:
        return ""
    if name in COMPANY_EXCEPT:
        return COMPANY_EXCEPT[name]
    result = name
    for k, v in KANJI_WORD_MAP.items():
        if k in result:
            result = result.replace(k, v)
    return result

# ======== Flaskビュー ========

html_form = """
<!DOCTYPE html>
<html lang="ja">
<head><meta charset="UTF-8"><title>Google→宛名職人 CSV 変換</title></head>
<body>
<h2>Google連絡先 → 宛名職人 CSV 変換ツール</h2>
<form action="/convert" method="post" enctype="multipart/form-data">
  <input type="file" name="file" accept=".csv" required>
  <input type="submit" value="変換開始">
</form>
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(html_form)

@app.route("/convert", methods=["POST"])
def convert():
    file = request.files["file"]
    if not file:
        return "⚠️ ファイルが選択されていません。"

    text = file.read().decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    output = io.StringIO()
    writer = csv.writer(output)

    header = [
        "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称",
        "ニックネーム","旧姓","宛先","自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話",
        "自宅IM ID","自宅E-mail","自宅URL","自宅Social",
        "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail",
        "会社URL","会社Social",
        "その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID",
        "その他E-mail","その他URL","その他Social",
        "会社名かな","会社名","部署名1","部署名2","役職名",
        "連名","連名ふりがな","連名敬称","連名誕生日",
        "メモ1","メモ2","メモ3","メモ4","メモ5",
        "備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
    ]
    writer.writerow(header)

    for row in reader:
        out = {}
        route_address_by_label(row, out)

        phones = [row.get("Phone 1 - Value", ""), row.get("Phone 2 - Value", "")]
        out["会社電話"] = normalize_phones(phones)

        emails = [
            row.get("E-mail 1 - Value", ""),
            row.get("E-mail 2 - Value", ""),
            row.get("E-mail 3 - Value", "")
        ]
        out["会社E-mail"] = normalize_emails(emails)

        memos = extract_memos(row)
        for i in range(5):
            out[f"メモ{i+1}"] = memos[i] if i < len(memos) else ""

        company_name = row.get("Organization Name", "")
        out["会社名"] = company_name
        out["会社名かな"] = kana_company_name(company_name)

        writer.writerow([
            row.get("Last Name",""), row.get("First Name",""),
            row.get("Phonetic Last Name",""), row.get("Phonetic First Name",""),
            f"{row.get('Last Name','')}　{row.get('First Name','')}",
            f"{row.get('Phonetic Last Name','')}{row.get('Phonetic First Name','')}",
            "", "", "様", row.get("Nickname",""), "",
            "会社",
            out.get("自宅〒",""), out.get("自宅住所1",""), out.get("自宅住所2",""), out.get("自宅住所3",""),
            out.get("自宅電話",""), "", "", "", "",
            out.get("会社〒",""), out.get("会社住所1",""), out.get("会社住所2",""), out.get("会社住所3",""),
            out.get("会社電話",""), "", out.get("会社E-mail",""), "", "",
            out.get("その他〒",""), out.get("その他住所1",""), out.get("その他住所2",""), out.get("その他住所3",""),
            out.get("その他電話",""), "", "", "", "",
            out.get("会社名かな",""), out.get("会社名",""),
            row.get("Organization Department",""), "", row.get("Organization Title",""),
            "","","","",
            out.get("メモ1",""), out.get("メモ2",""), out.get("メモ3",""), out.get("メモ4",""), out.get("メモ5",""),
            "", "", "", "", "", "", "", ""
        ])

    output.seek(0)
    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name="converted.csv"
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
