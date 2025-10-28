import csv
import io
import re
from flask import Flask, request, render_template, send_file, jsonify
from datetime import datetime

app = Flask(__name__)

# =========================================================
# ユーティリティ
# =========================================================
def to_zenkaku(s: str) -> str:
    """半角→全角（英数字＋スペース）"""
    if not s:
        return ""
    table = str.maketrans({
        **{chr(i): chr(i + 0xFEE0) for i in range(0x21, 0x7F)},
        " ": "　"
    })
    return s.translate(table)


def normalize_address(addr: str):
    """住所を最初のスペース（全角/半角）で二分割"""
    if not addr:
        return "", ""
    s = to_zenkaku(addr.strip())
    parts = re.split(r"[ 　]+", s, 1)
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[1]


def company_name_kana(name: str) -> str:
    """法人格を除外し本文を全角カタカナ化"""
    if not name:
        return ""
    name = re.sub(r"(株式会社|有限会社|合同会社|一般社団法人|一般財団法人|学校法人|医療法人|社会福祉法人)", "", name)
    name = to_zenkaku(name)
    name = re.sub(r"[ぁ-ん]", lambda m: chr(ord(m.group(0)) + 0x60), name)
    return name


def get_val(row, *keys):
    """複数候補キーから最初に存在する値を返す"""
    for k in keys:
        if k in row and row[k].strip():
            return row[k].strip()
    return ""


# =========================================================
# Flask ルート
# =========================================================
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    if not file:
        return jsonify({"error": "ファイルが選択されていません"}), 400

    raw = file.stream.read().decode('utf-8-sig')
    delimiter = '\t' if '\t' in raw.split('\n')[0] else ','

    input_stream = io.StringIO(raw)
    reader = csv.DictReader(input_stream, delimiter=delimiter)

    output = io.StringIO()
    writer = csv.writer(output, delimiter='\t', lineterminator='\n')

    # 理想の宛名職人フォーマットに合わせたヘッダー
    headers = [
        "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな",
        "敬称","ニックネーム","旧姓","宛先","自宅〒","自宅住所1","自宅住所2","自宅住所3",
        "自宅電話1〜10","自宅IM ID1〜10","自宅E-mail1〜5","自宅URL1〜5","自宅Social1〜10",
        "会社〒","会社住所1","会社住所2","会社住所3","会社電話1〜10","会社IM ID1〜10",
        "会社E-mail1〜5","会社URL1〜5","会社Social1〜10","その他〒","その他住所1","その他住所2",
        "その他住所3","その他電話1〜10","その他IM ID1〜10","その他E-mail1〜5","その他URL1〜5",
        "その他Social1〜10","会社名かな","会社名","部署名1","部署名2","役職名",
        "連名1〜20","連名ふりがな1〜20","連名敬称1〜20","連名誕生日1〜20",
        "メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3",
        "誕生日","性別","血液型","趣味","性格"
    ]
    writer.writerow(headers)

    for row in reader:
        sei = get_val(row, "Last Name", "姓")
        mei = get_val(row, "First Name", "名")
        sei_kana = get_val(row, "Phonetic Last Name", "せい")
        mei_kana = get_val(row, "Phonetic First Name", "めい")
        mid = get_val(row, "Middle Name", "ミドルネーム")
        mid_kana = get_val(row, "Phonetic Middle Name", "みどるねーむ")

        full = f"{sei}　{mei}".strip()
        full_kana = f"{sei_kana}　{mei_kana}".strip()

        company = get_val(row, "Organization Name", "会社名", "株式会社社名")
        company_kana = company_name_kana(company)
        dept = get_val(row, "Organization Department", "部署")
        title = get_val(row, "Organization Title", "肩書き")
        nickname = get_val(row, "Nickname", "ニックネーム")
        birthday = get_val(row, "Birthday", "誕生日")
        note = get_val(row, "Notes", "ノート")

        # メモ項目（連番対応）
        memo1 = get_val(row, "メモ1")
        memo2 = get_val(row, "メモ2")
        memo3 = get_val(row, "メモ3")
        memo4 = get_val(row, "メモ4")
        memo5 = get_val(row, "メモ5")

        # 住所（Street から分割）
        addr = get_val(row, "Address 1 - Street")
        addr1, addr2 = normalize_address(addr)
        zip_code = get_val(row, "Address 1 - Postal Code")

        # 連絡先
        company_phones = ";".join([
            get_val(row, "Phone 1 - Value"),
            get_val(row, "Phone 2 - Value"),
            get_val(row, "Phone 3 - Value")
        ]).strip(";")
        company_emails = ";".join([
            get_val(row, "E-mail 1 - Value"),
            get_val(row, "E-mail 2 - Value")
        ]).strip(";")
        home_emails = get_val(row, "E-mail 3 - Value")

        writer.writerow([
            sei, mei, sei_kana, mei_kana, full, full_kana, mid, mid_kana,
            "様", nickname, "", "会社", "", "", "", "",
            "", "", home_emails, "", "",
            zip_code, addr1, addr2, "", company_phones, "", company_emails, "", "",
            "", "", "", "", "", "", "", "", "",
            company_kana, company, dept, "", title,
            "", "", "", "", memo1, memo2, memo3, memo4, memo5,
            note, "", "", birthday, "選択なし", "選択なし", "", ""
        ])

    output.seek(0)
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        io.BytesIO(output.getvalue().encode('utf-8-sig')),
        mimetype="text/tab-separated-values",
        as_attachment=True,
        download_name=f"converted_{now}.tsv"
    )


if __name__ == '__main__':
    app.run(debug=True)
