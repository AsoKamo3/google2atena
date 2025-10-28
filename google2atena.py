import csv, io, re, unicodedata
from flask import Flask, request, send_file, render_template, jsonify
from datetime import datetime

VERSION = "v3.8 fix 完全版"

app = Flask(__name__)

# ========== Utility ==========
def to_zenkaku(s):
    if not s: return ""
    table = str.maketrans({**{chr(i): chr(i+0xFEE0) for i in range(0x21,0x7F)}, " ": "　"})
    return s.translate(table)

def hira_to_kata(s):
    return re.sub(r"[ぁ-ん]", lambda m: chr(ord(m.group(0)) + 0x60), s or "")

def norm_key(s):
    if not s: return ""
    s = unicodedata.normalize("NFKC", s).lower()
    s = re.sub(r"[ \u3000\-_/・.]+", "", s)
    return s

def company_name_kana(name):
    if not name: return ""
    name = re.sub(r"(株式会社|有限会社|合同会社|一般社団法人|一般財団法人|学校法人|医療法人|社会福祉法人)", "", name)
    name = to_zenkaku(name.strip())
    return hira_to_kata(name)

def normalize_address(addr):
    if not addr: return "", ""
    s = addr.strip()
    s = to_zenkaku(s)
    parts = re.split(r"[ 　]+", s, 1)
    return (parts[0], parts[1]) if len(parts) > 1 else (s, "")

def pick(row, *keys):
    for k in keys:
        if k in row and row[k].strip():
            return row[k].strip()
    return ""

def detect_reader(raw):
    head = "\n".join(raw.splitlines()[:2])
    try:
        dialect = csv.Sniffer().sniff(head, delimiters=[",", "\t"])
        delim = dialect.delimiter
    except:
        delim = "\t" if "\t" in head else ","
    return csv.DictReader(io.StringIO(raw), delimiter=delim)

# ========== Flask ==========
@app.route("/")
def index():
    # 起動時のタイトル画面
    html = f"""
    <html>
        <head><meta charset="utf-8"><title>Google連絡先CSV → 宛名職人CSV 変換 ({VERSION})</title></head>
        <body style="font-family: sans-serif; margin: 40px;">
            <h2>Google連絡先CSV → 宛名職人CSV 変換（{VERSION}）</h2>
            <p>Google 連絡先のエクスポートファイル（CSV または TSV）を選択してください。</p>
            <form action="/convert" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".csv,.tsv" required>
                <br><br>
                <input type="submit" value="変換開始">
            </form>
            <p style="margin-top:20px; color:#555;">変換後のファイルは UTF-8（BOM付）の TSV形式でダウンロードされます。</p>
        </body>
    </html>
    """
    return html


@app.route("/convert", methods=["POST"])
def convert():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "ファイル未選択"}), 400

    raw = f.stream.read().decode("utf-8-sig")
    reader = detect_reader(raw)
    out = io.StringIO()
    w = csv.writer(out, delimiter="\t", lineterminator="\n")

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
    w.writerow(headers)

    for r in reader:
        row = {norm_key(k): (v or "").strip() for k, v in r.items()}

        last = pick(row, "lastname", "姓")
        first = pick(row, "firstname", "名")
        last_k = pick(row, "phoneticlastname", "せい")
        first_k = pick(row, "phoneticfirstname", "めい")
        mid = pick(row, "middlename", "ミドルネーム")
        mid_k = pick(row, "phoneticmiddlename", "みどるねーむ")

        full = "　".join(x for x in [last, first] if x)
        full_k = "　".join(x for x in [last_k, first_k] if x)

        org = pick(row, "organizationname", "会社名")
        org_kana = company_name_kana(org)
        dept = pick(row, "organizationdepartment", "部署")
        title = pick(row, "organizationtitle", "肩書き")
        nickname = pick(row, "nickname", "ニックネーム")
        birthday = pick(row, "birthday", "誕生日")
        note = pick(row, "notes", "ノート")

        street = pick(row, "address1street", "address1formatted")
        zip_code = pick(row, "address1postalcode", "郵便番号")
        addr1, addr2 = normalize_address(street)

        phone_all = ";".join(filter(None, [
            pick(row, "phone1value"), pick(row, "phone2value"), pick(row, "phone3value")
        ]))
        email_company = ";".join(filter(None, [
            pick(row, "email1value"), pick(row, "email2value")
        ]))
        email_home = pick(row, "email3value")

        memos = [r.get(f"メモ{i}", "") for i in range(1,6)]
        relations = [r.get(f"Relation {i} - Value", "") for i in range(1,6)]
        for i, rel in enumerate(relations):
            if rel and not memos[i]:
                memos[i] = rel

        w.writerow([
            last, first, last_k, first_k, full, full_k, mid, mid_k,
            "様", nickname, "", "会社", "", "", "", "",
            "", "", email_home, "", "",
            to_zenkaku(zip_code), to_zenkaku(addr1), to_zenkaku(addr2), "",
            to_zenkaku(phone_all), "", email_company, "", "",
            "", "", "", "", "", "", "", "", "",
            org_kana, org, dept, "", title,
            "", "", "", "",
            *memos, note, "", "",
            birthday, "選択なし", "選択なし", "", ""
        ])

    out.seek(0)
    return send_file(io.BytesIO(out.getvalue().encode("utf-8-sig")),
                     mimetype="text/tab-separated-values",
                     as_attachment=True,
                     download_name=f"converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.tsv")

if __name__ == "__main__":
    print(f"Google連絡先CSV → 宛名職人CSV 変換（{VERSION}）を起動しました。")
    app.run(debug=True)
