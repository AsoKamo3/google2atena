# google2atena.py  v3.9.18r5  (Render改善安定版 / no-pandas)
# - フェイルセーフ対応（外部辞書が無い場合も動作）
# - 住所整形：Region + City + Street、全角化、最初の空白で建物名分離
# - 電話整形：桁数判定＋ゼロ補完＋ハイフン挿入、;連結、重複排除
# - メール整形：;連結、重複排除
# - 会社かな：外部辞書（company_dicts / kanji_word_map / corp_terms）＋英字→カタカナ変換
# - 出力：CSV (UTF-8-SIG) ダウンロード対応

import io
import csv
import re
import unicodedata
from flask import Flask, render_template, request, send_file

# --- 外部辞書フェイルセーフ読込 ---
try:
    from company_dicts import COMPANY_EXCEPT
except ImportError:
    COMPANY_EXCEPT = {}
try:
    from kanji_word_map import KANJI_WORD_MAP, EN_TO_KATAKANA
except ImportError:
    KANJI_WORD_MAP, EN_TO_KATAKANA = {}, {}
try:
    from corp_terms import CORP_TERMS
except ImportError:
    CORP_TERMS = ["株式会社", "有限会社", "合同会社", "Inc.", "Co.", "Ltd."]

# --- フェイルセーフ辞書補完 ---
if not COMPANY_EXCEPT:
    COMPANY_EXCEPT = {"ＮＨＫエデュケーショナル": "エヌエイチケーエデュケーショナル"}
if not KANJI_WORD_MAP:
    KANJI_WORD_MAP = {"社": "シャ", "新聞": "シンブン", "放送": "ホウソウ"}
if not EN_TO_KATAKANA:
    EN_TO_KATAKANA = {
        'A': 'エー','B': 'ビー','C': 'シー','D': 'ディー','E': 'イー','F': 'エフ','G': 'ジー',
        'H': 'エイチ','I': 'アイ','J': 'ジェー','K': 'ケー','L': 'エル','M': 'エム','N': 'エヌ',
        'O': 'オー','P': 'ピー','Q': 'キュー','R': 'アール','S': 'エス','T': 'ティー','U': 'ユー',
        'V': 'ブイ','W': 'ダブリュー','X': 'エックス','Y': 'ワイ','Z': 'ズィー','&': 'アンド',
        '+': 'プラス','-': 'ハイフン'
    }

# --- 英字→カタカナ変換 ---
def alpha_to_katakana(text):
    return "".join(EN_TO_KATAKANA.get(ch.upper(), ch) for ch in text)

# --- 半角英数→全角変換 ---
def to_zenkaku(text):
    return "".join(chr(ord(c) + 0xFEE0) if "!" <= c <= "~" else c for c in text)

# --- 住所整形 ---
def normalize_address(row):
    region = (row.get("Address 1 - Region") or "").strip()
    city = (row.get("Address 1 - City") or "").strip()
    street = (row.get("Address 1 - Street") or "").strip()
    zipcode = re.sub(r"[^\d]", "", row.get("Address 1 - Postal Code", ""))
    zipcode = f"{zipcode[:3]}-{zipcode[3:]}" if len(zipcode) == 7 else ""

    full = f"{region}{city}{street}"
    if not full:
        return zipcode, "", "", ""
    full = to_zenkaku(unicodedata.normalize("NFKC", full))
    full = re.sub(r"\s+", " ", full.strip())

    addr1, addr2 = (full.split(" ", 1) + [""])[:2]
    return zipcode, addr1.strip(), addr2.strip(), ""

# --- 電話番号整形 ---
def normalize_phones(values):
    if not values:
        return ""
    phones = []
    for v in values:
        for part in re.split(r"[,;:：／/ ::: ]+", v):
            digits = re.sub(r"\D", "", part)
            if not digits:
                continue
            digits = re.sub(r"^\+?81", "0", digits)  # 国際→国内
            if len(digits) == 11:
                digits = f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
            elif len(digits) == 10:
                digits = f"{digits[:2]}-{digits[2:6]}-{digits[6:]}"
            elif len(digits) == 9:
                digits = f"0{digits[:2]}-{digits[2:5]}-{digits[5:]}"
            phones.append(digits)
    return ";".join(sorted(set(phones)))

# --- メール整形 ---
def normalize_emails(values):
    if not values:
        return ""
    allmails = []
    for v in values:
        if not v:
            continue
        for part in re.split(r"[,;:：:::／/ ]+", v):
            if "@" in part:
                allmails.append(part.strip())
    return ";".join(sorted(set(allmails)))

# --- 会社名かな生成 ---
def kana_company_name(org):
    if not org:
        return "", ""
    clean = org
    for term in CORP_TERMS:
        clean = clean.replace(term, "").strip()
    kana = COMPANY_EXCEPT.get(clean, "")
    if not kana:
        kana = "".join(KANJI_WORD_MAP.get(ch, ch) for ch in clean)
        kana = alpha_to_katakana(kana)
    return kana, org.strip()

# --- Flaskアプリ ---
app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    file = request.files.get("file")
    if not file:
        return "⚠️ ファイルが選択されていません。", 400

    import chardet
    raw = file.read()
    enc = chardet.detect(raw)["encoding"] or "utf-8"
    text = raw.decode(enc, errors="ignore")

    sniffer = csv.Sniffer()
    dialect = sniffer.sniff(text.splitlines()[0])
    reader = csv.DictReader(io.StringIO(text), dialect=dialect)

    output = io.StringIO()
    writer = csv.writer(output, quoting=csv.QUOTE_ALL)

    headers = ["姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな",
               "敬称","ニックネーム","旧姓","宛先","自宅〒","自宅住所1","自宅住所2","自宅住所3",
               "自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
               "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail",
               "会社URL","会社Social","その他〒","その他住所1","その他住所2","その他住所3",
               "その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
               "会社名かな","会社名","部署名1","部署名2","役職名","連名","連名ふりがな","連名敬称",
               "連名誕生日","メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3",
               "誕生日","性別","血液型","趣味","性格"]
    writer.writerow(headers)

    for row in reader:
        first = row.get("First Name", "").strip()
        last = row.get("Last Name", "").strip()
        first_k = row.get("Phonetic First Name", "").strip()
        last_k = row.get("Phonetic Last Name", "").strip()
        org = row.get("Organization Name", "").strip()
        dept = row.get("Organization Department", "").strip()
        title = row.get("Organization Title", "").strip()

        addr_zip, addr1, addr2, addr3 = normalize_address(row)
        phones = normalize_phones([row.get("Phone 1 - Value",""), row.get("Phone 2 - Value","")])
        emails = normalize_emails([row.get("E-mail 1 - Value",""), row.get("E-mail 2 - Value",""), row.get("E-mail 3 - Value","")])
        kana_org, org_name = kana_company_name(org)

        writer.writerow([
            last, first, last_k, first_k, f"{last}　{first}", f"{last_k}　{first_k}", "", "", "様", "", "",
            "会社", addr_zip, addr1, addr2, addr3, phones, "", emails, "", "",
            addr_zip, addr1, addr2, addr3, phones, "", emails, "", "", "", "", "", "", "", "", "", "", "",
            kana_org, org_name, dept, "", title, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        ])

    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode("utf-8-sig")),
                     mimetype="text/csv",
                     as_attachment=True,
                     download_name="converted.csv")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
