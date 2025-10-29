# google2atena.py
# v3.9.18r4（住所・電話整形強化版／Render安定構成）

import csv
import io
import re
import chardet
from flask import Flask, request, render_template, send_file

# --- 外部辞書読込 ---
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
    CORP_TERMS = []

try:
    from jp_area_codes import AREA_CODES
except Exception:
    AREA_CODES = {}

# --- Flask 初期化 ---
app = Flask(__name__)

# --- 文字変換ユーティリティ ---
def to_zenkaku_address(text):
    """住所用：英数字・記号を全角に"""
    if not text:
        return ""
    return text.translate(str.maketrans({
        **{chr(i): chr(i + 0xFEE0) for i in range(0x21, 0x7F)},
        '-': '－',
    }))

def to_hankaku_numhyphen(text):
    """電話・郵便番号用：数字とハイフンを半角に"""
    if not text:
        return ""
    return text.translate(str.maketrans({
        **{chr(i + 0xFEE0): chr(i) for i in range(0xFF10, 0xFF1A)},
        'ー': '-',
        '−': '-',
        '―': '-',
        '‐': '-',
        '－': '-'
    }))

# --- 電話番号整形 ---
def normalize_phones(*phones):
    numbers = []
    for p in phones:
        if not p:
            continue
        parts = re.split(r"[;:：\s／/]+", p)
        for n in parts:
            n = re.sub(r"[^\d]", "", n)
            if not n:
                continue
            # 先頭0補正
            if re.match(r"^9\d{8,9}$", n):
                n = "0" + n  # 携帯0抜け補正
            elif re.match(r"^[1-9]\d{8,9}$", n):
                n = "0" + n  # 固定電話0抜け補正

            # 市外局番マッチング（ハイフン挿入）
            n = hyphenate_phone_by_area(n)
            if n not in numbers:
                numbers.append(n)
    return ";".join(numbers)

def hyphenate_phone_by_area(num):
    """AREA_CODES辞書に基づいて電話番号にハイフンを入れる"""
    if not num:
        return ""
    for prefix in sorted(AREA_CODES.keys(), key=lambda x: -len(x)):
        if num.startswith(prefix):
            body_len = AREA_CODES[prefix]
            return f"{prefix}-{num[len(prefix):len(prefix)+body_len]}-{num[len(prefix)+body_len:]}"
    # fallback
    if len(num) == 11:
        return f"{num[:3]}-{num[3:7]}-{num[7:]}"
    if len(num) == 10:
        return f"{num[:2]}-{num[2:6]}-{num[6:]}"
    return num

# --- 住所整形 ---
def split_building(street):
    """最初に出てくるスペースまたは建物語で分割"""
    if not street:
        return "", ""
    street = street.strip()
    # ビル・マンション等を基準に分離
    m = re.search(r"(.+?)(?:(?:　|\s|,|、|,)(.*[ビル|マンション|タワー|荘|ハイツ|ヒルズ|#].*))", street)
    if m:
        return m.group(1), m.group(2)
    # fallback: 最初のスペースで分割
    parts = re.split(r"[ 　]", street, 1)
    if len(parts) == 2:
        return parts[0], parts[1]
    return street, ""

def split_google_address(addr):
    """Google連絡先形式住所から都道府県、市区町村、番地等を抽出"""
    if not addr:
        return "", "", "", "", "", "", "", "", "", "", ""
    lines = [line.strip() for line in addr.split("\n") if line.strip()]
    street = city = region = postal = country = ""
    for line in lines:
        if re.match(r"^\d{3}-\d{4}$", line):
            postal = to_hankaku_numhyphen(line)
        elif line in ["日本", "Japan"]:
            country = "日本"
        elif re.search(r"(都|道|府|県)$", line):
            region = line
        elif re.search(r"(区|市|町|村)", line):
            city = line
        else:
            street = line
    # 建物分割を street のみ対象に
    addr2, bldg = split_building(street)
    return postal, region, city, addr2, bldg, country

# --- 会社名かな変換 ---
def kana_company_name(org):
    if not org:
        return "", ""
    clean = org
    for term in CORP_TERMS:
        clean = clean.replace(term, "").strip()
    kana = COMPANY_EXCEPT.get(clean, "")
    if not kana:
        kana = "".join(KANJI_WORD_MAP.get(ch, ch) for ch in clean)
    return kana, org.strip()

# --- 変換メイン ---
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return "⚠️ ファイルが見つかりません。"
    file = request.files["file"]
    if file.filename == "":
        return "⚠️ ファイルを選択してください。"

    data = file.read()
    encoding = chardet.detect(data)["encoding"] or "utf-8"
    text = data.decode(encoding, errors="ignore")
    sample = text[:1024]
    dialect = csv.Sniffer().sniff(sample)
    reader = csv.DictReader(io.StringIO(text), dialect=dialect)

    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")

    # ヘッダ（既存仕様通り）
    headers = [
        "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称","ニックネーム",
        "旧姓","宛先","自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
        "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social",
        "その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
        "会社名かな","会社名","部署名1","部署名2","役職名","連名","連名ふりがな","連名敬称","連名誕生日",
        "メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
    ]
    writer.writerow(headers)

    for row in reader:
        sei, mei = row.get("Last Name", ""), row.get("First Name", "")
        sei_kana, mei_kana = row.get("Phonetic Last Name", ""), row.get("Phonetic First Name", "")
        full = f"{sei}　{mei}".strip()
        full_kana = f"{sei_kana}　{mei_kana}".strip()
        org = row.get("Organization Name", "")
        org_kana, org_full = kana_company_name(org)

        # 住所
        a1_label = row.get("Address 1 - Label", "")
        postal, region, city, addr2, bldg, country = split_google_address(row.get("Address 1 - Formatted", ""))
        home_postal = other_postal = ""
        home_addr1 = home_addr2 = home_addr3 = ""
        other_addr1 = other_addr2 = other_addr3 = ""
        if a1_label == "Home":
            home_postal, home_addr1, home_addr2, home_addr3 = postal, region+city+addr2, bldg, country
        else:
            other_postal, other_addr1, other_addr2, other_addr3 = postal, region+city+addr2, bldg, country

        # 電話
        phone = normalize_phones(row.get("Phone 1 - Value", ""), row.get("Phone 2 - Value", ""))

        # メール
        emails = [v for k, v in row.items() if "E-mail" in k and v]
        email_combined = ";".join(emails)

        # メモ
        notes = [row.get(f"メモ{i}", "") for i in range(1, 6)]

        writer.writerow([
            sei, mei, sei_kana, mei_kana, full, full_kana, "", "", "様", row.get("Nickname", ""), "",
            "会社",
            home_postal, home_addr1, home_addr2, home_addr3,
            "", "", "", "", "",
            postal, region+city+addr2, bldg, country,
            phone, "", email_combined, "", "",
            other_postal, other_addr1, other_addr2, other_addr3,
            "", "", "", "", "",
            org_kana, org_full, row.get("Organization Department", ""), "", row.get("Organization Title", ""),
            "", "", "", "",
            *notes,
            "", "", "", row.get("Birthday", ""), "選択なし", "選択なし", "", ""
        ])

    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode("utf-8-sig")),
                     mimetype="text/csv",
                     as_attachment=True,
                     download_name="converted_v3.9.18r4.csv")

# --- Render 実行 ---
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
