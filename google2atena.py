import csv
import io
import re
from flask import Flask, request, render_template, send_file, jsonify
from datetime import datetime

app = Flask(__name__)

# =========================================================
# 住所の最初のスペース以降を建物名・階層・号室として扱う簡潔ロジック
# =========================================================
def to_zenkaku(s: str) -> str:
    """半角英数字・カナを全角へ"""
    if not s:
        return ""
    table = str.maketrans({
        **{chr(i): chr(i + 0xFEE0) for i in range(0x21, 0x7F)},  # 英数
        " ": "　"
    })
    return s.translate(table)

def normalize_address(addr: str):
    """住所を最初のスペース（半角 or 全角）で二分割"""
    if not addr:
        return "", ""
    s = to_zenkaku(addr.strip())
    parts = re.split(r"[ 　]+", s, 1)  # 最初のスペースで1回だけ分割
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[1]


# =========================================================
# 会社名かな（法人格除外・全角カタカナ変換）
# =========================================================
def company_name_kana(name: str) -> str:
    """会社名から法人格を除外し、カタカナ化"""
    if not name:
        return ""
    name = re.sub(r"(株式会社|有限会社|合同会社|一般社団法人|一般財団法人|学校法人|医療法人|社会福祉法人)", "", name)
    name = to_zenkaku(name)
    # ひらがな→カタカナ
    name = re.sub(r"[ぁ-ん]", lambda m: chr(ord(m.group(0)) + 0x60), name)
    return name


# =========================================================
# Flaskルート
# =========================================================
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    if not file:
        return jsonify({"error": "ファイルが選択されていません"}), 400

    input_stream = io.StringIO(file.stream.read().decode('utf-8-sig'))
    reader = csv.DictReader(input_stream, delimiter='\t')

    output = io.StringIO()
    writer = csv.writer(output, delimiter='\t', lineterminator='\n')

    # 理想の出力ヘッダー順
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
        sei = row.get("姓", "")
        mei = row.get("名", "")
        sei_kana = row.get("せい", "")
        mei_kana = row.get("めい", "")
        full = f"{sei}　{mei}".strip()
        full_kana = f"{sei_kana}　{mei_kana}".strip()
        nickname = row.get("ニックネーム", "")
        company = row.get("株式会社社名", "") or row.get("会社名", "")
        dept = row.get("部署", "")
        title = row.get("肩書き", "")
        birthday = row.get("誕生日", "")
        note = row.get("ノート", "")
        memo1 = row.get("メモ1", "")
        memo2 = row.get("メモ2", "")
        memo3 = row.get("メモ3", "")
        memo4 = row.get("メモ4", "")
        memo5 = row.get("メモ5", "")

        addr = row.get("Address 1 - Street", "")
        addr1, addr2 = normalize_address(addr)

        company_kana = company_name_kana(company)

        # 各種連絡先（会社・自宅メール/電話）
        company_phones = ";".join([
            row.get("Phone 1 - Value", ""),
            row.get("Phone 2 - Value", "")
        ]).strip(";")
        company_emails = ";".join([
            row.get("E-mail 1 - Value", ""),
            row.get("E-mail 2 - Value", "")
        ]).strip(";")
        home_emails = row.get("E-mail 3 - Value", "")

        writer.writerow([
            sei, mei, sei_kana, mei_kana, full, full_kana, "", "", "様", nickname, "", "会社",
            "", "", "", "", "", "", home_emails, "", "",
            row.get("Address 1 - Postal Code", ""), addr1, addr2, "", company_phones, "", company_emails, "", "",
            "", "", "", "", "", "", "", "", "", company_kana, company, dept, "", title,
            "", "", "", "", memo1, memo2, memo3, memo4, memo5, note, "", "", birthday,
            "選択なし", "選択なし", "", ""
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
