# google2atena.py  v3.9.18r7a  (Render安定版 / no-pandas)
# - 住所振り分け: Address 1 - Label に応じて Work/Home/Other 各住所カラムへ
# - 住所分割: Region + City + Street → 全角化 → 最初の空白で分割
# - 郵便番号: 半角数字＋半角ハイフン
# - 電話・メール・メモなど他機能は v3.9.18r6 と同じ
# - 出力: CSV (UTF-8-SIG)

import csv
import io
import re
import unicodedata
from flask import Flask, render_template_string, request, send_file

# ======== 外部辞書のフェイルセーフ読込 ========
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

# ======== アプリケーション設定 ========
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
        return f"{digits[0:3]}-{digits[3:7]}"
    return postal

_SPLIT_RE = re.compile(r'[ \u3000]')

def split_first_space(addr_full: str) -> tuple[str, str]:
    if not addr_full:
        return ("", "")
    m = _SPLIT_RE.search(addr_full)
    if not m:
        return (addr_full, "")
    i = m.start()
    return (addr_full[:i], addr_full[i+1:].strip())

def build_addr12(region: str, city: str, street: str) -> tuple[str, str]:
    parts = [p for p in [region, city, street] if p]
    full = "".join(parts)
    full_z = to_zenkaku_for_address(full)
    a1, a2 = split_first_space(full_z)
    return (a1, a2)

def route_address_by_label(row: dict, out: dict) -> None:
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

# ======== 電話番号・メール整形 ========

def normalize_phones(phone_values):
    phones = []
    for val in phone_values:
        if not val:
            continue
        nums = re.sub(r'\D', '', val)
        if not nums:
            continue
        # 桁数から0補完
        if not nums.startswith('0'):
            nums = '0' + nums
        # 11桁携帯 or 固定電話パターン
        if re.match(r'^0\d{9,10}$', nums):
            if nums.startswith(('070', '080', '090', '050', '0120', '0800', '0570')):
                # 携帯/VoIP/フリーダイヤル/ナビダイヤル
                if len(nums) == 11:
                    val = f"{nums[:3]}-{nums[3:7]}-{nums[7:]}"
                elif len(nums) == 10:
                    val = f"{nums[:4]}-{nums[4:7]}-{nums[7:]}"
                else:
                    val = nums
            else:
                # 市外局番推定
                if len(nums) == 10:
                    val = f"{nums[:2]}-{nums[2:6]}-{nums[6:]}"
                else:
                    val = nums
        else:
            val = nums
        phones.append(val)
    uniq = []
    for p in phones:
        if p not in uniq:
            uniq.append(p)
    return ";".join(uniq)

def normalize_emails(email_values):
    emails = []
    for val in email_values:
        if not val:
            continue
        val = val.strip()
        if val and val not in emails:
            emails.append(val)
    return ";".join(emails)

# ======== メモ抽出 ========

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

# ======== 会社かな変換 ========

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

# ======== HTMLフォーム ========

html_form = """
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>Google→宛名職人 CSV 変換</title>
</head>
<body>
<h2>Google連絡先 → 宛名職人 CSV 変換ツール</h2>
<form action="/convert" method="post" enctype="multipart/form-data">
  <input type="file" name="file" accept=".csv" required>
  <input type="submit" value="変換開始">
</form>
</body>
</html>
"""

# ======== Flask Routes ========

@app.route("/")
def index():
    return render_template_string(html_form)

@app.route("/convert", methods=["POST"])
def convert():
    file = request.files["file"]
    if not file:
        return "⚠️ ファイルが選択されていません。"

    data = file.read()
    text = data.decode("utf-8-sig", errors="replace")
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

        # --- 住所 ---
        route_address_by_label(row, out)

        # --- 電話 ---
        phones = [row.get("Phone 1 - Value", ""), row.get("Phone 2 - Value", "")]
        out["会社電話"] = normalize_phones(phones)

        # --- メール ---
        emails = [
            row.get("E-mail 1 - Value", ""),
            row.get("E-mail 2 - Value", ""),
            row.get("E-mail 3 - Value", "")
        ]
        out["会社E-mail"] = normalize_emails(emails)

        # --- メモ ---
        memos = extract_memos(row)
        for i in range(5):
            out[f"メモ{i+1}"] = memos[i] if i < len(memos) else ""

        # --- 会社名かな ---
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
