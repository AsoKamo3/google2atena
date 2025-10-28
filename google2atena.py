import csv
import io
import re
import unicodedata
from flask import Flask, request, render_template, send_file, jsonify
from datetime import datetime

app = Flask(__name__)

# =========================
# 正規化ユーティリティ
# =========================
def to_zenkaku(s: str) -> str:
    if not s:
        return ""
    # 半角英数・スペースのみ全角化
    table = str.maketrans({**{chr(i): chr(i + 0xFEE0) for i in range(0x21, 0x7F)}, " ": "　"})
    return s.translate(table)

def hira_to_kata(s: str) -> str:
    if not s:
        return ""
    return re.sub(r"[ぁ-ん]", lambda m: chr(ord(m.group(0)) + 0x60), s)

def norm_key(s: str) -> str:
    """ヘッダー名のゆらぎを吸収：NFKC→小文字→全角/半角スペース削除→ハイフン/記号除去"""
    if s is None:
        return ""
    x = unicodedata.normalize("NFKC", s)
    x = x.lower()
    x = x.replace("－", "-").replace("–", "-").replace("—", "-")
    x = re.sub(r"[ \u3000]+", "", x)          # 半角/全角スペース除去
    x = re.sub(r"[-_/・.]+", "", x)           # 記号除去
    return x

# ターゲット→候補ヘッダー（正規化した同義語）
KEY_ALIASES = {
    "firstname": ["firstname", "名", "givenname"],
    "middlename": ["middlename", "ミドルネーム"],
    "lastname": ["lastname", "姓", "familyname", "surname"],
    "phoneticfirstname": ["phoneticfirstname", "phoneticfirst", "めい"],
    "phoneticmiddlename": ["phoneticmiddlename", "みどるねーむ"],
    "phoneticlastname": ["phoneticlastname", "phoneticlast", "せい"],
    "nickname": ["nickname", "ニックネーム"],
    "orgname": ["organizationname", "organization", "会社名", "株式会社社名"],
    "orgdept": ["organizationdepartment", "部署", "department"],
    "orgtitle": ["organizationtitle", "肩書き", "title"],
    "birthday": ["birthday", "誕生日"],
    "notes": ["notes", "ノート"],
    "email1": ["email1value", "e-mail1value", "email1", "e-mail1"],
    "email2": ["email2value", "e-mail2value", "email2", "e-mail2"],
    "email3": ["email3value", "e-mail3value", "email3", "e-mail3"],
    "phone1": ["phone1value", "phone1", "電話1"],
    "phone2": ["phone2value", "phone2", "電話2"],
    "phone3": ["phone3value", "phone3", "電話3"],
    "addr_street": ["address1street", "address1formatted", "住所1street"],
    "addr_zip": ["address1postalcode", "郵便番号", "zip", "postalcode"],
}

def pick(row_norm: dict, target_key: str) -> str:
    """正規化済みrowから、エイリアス群のどれか一致を拾う"""
    for alias in KEY_ALIASES.get(target_key, []):
        if alias in row_norm and row_norm[alias].strip():
            return row_norm[alias].strip()
    return ""

def normalize_address(addr: str):
    """住所の最初のスペース（全角/半角）で二分割"""
    if not addr:
        return "", ""
    s = addr.strip()
    # まずは NFKC で幅をそろえ、スペースはどちらも許容
    s = unicodedata.normalize("NFKC", s)
    parts = re.split(r"[ \u3000]+", s, 1)
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[1]

CORP_PATTERN = re.compile(r"(株式会社|有限会社|合同会社|一般社団法人|一般財団法人|学校法人|医療法人|社会福祉法人)")

def company_name_kana(name: str) -> str:
    if not name:
        return ""
    body = CORP_PATTERN.sub("", name)
    body = to_zenkaku(body)
    body = hira_to_kata(body)
    return body

# =========================
# CSV 読み込み（堅牢化）
# =========================
def detect_reader(raw_text: str):
    """Snifferで判定→失敗時は \t → , の順でフォールバック"""
    # 先頭1〜2行で判定
    head = "\n".join(raw_text.splitlines()[:2])
    try:
        dialect = csv.Sniffer().sniff(head, delimiters=[",", "\t", ";"])
        delim = dialect.delimiter
    except Exception:
        delim = "\t" if "\t" in head else ("," if "," in head else "\t")

    return csv.DictReader(io.StringIO(raw_text), delimiter=delim)

# =========================
# Flask routes
# =========================
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "ファイルが選択されていません"}), 400

    raw = f.stream.read().decode("utf-8-sig")
    reader = detect_reader(raw)

    out = io.StringIO()
    w = csv.writer(out, delimiter="\t", lineterminator="\n")

    headers = [
        "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな",
        "敬称","ニックネーム","旧姓","宛先","自宅〒","自宅住所1","自宅住所2","自宅住所3",
        "自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
        "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID",
        "会社E-mail","会社URL","会社Social","その他〒","その他住所1","その他住所2",
        "その他住所3","その他電話","その他IM ID","その他E-mail","その他URL",
        "その他Social","会社名かな","会社名","部署名1","部署名2","役職名",
        "連名","連名ふりがな","連名敬称","連名誕生日",
        "メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3",
        "誕生日","性別","血液型","趣味","性格"
    ]
    w.writerow(headers)

    for raw_row in reader:
        # ヘッダー正規化：{正規化キー: 値}
        row_norm = {}
        for k, v in raw_row.items():
            nk = norm_key(k)
            row_norm[nk] = v or ""

        # 人名
        last = pick(row_norm, "lastname")
        first = pick(row_norm, "firstname")
        last_k = pick(row_norm, "phoneticlastname")
        first_k = pick(row_norm, "phoneticfirstname")
        mid = pick(row_norm, "middlename")
        mid_k = pick(row_norm, "phoneticmiddlename")

        full = "　".join([x for x in [last, first] if x])
        full_k = "　".join([x for x in [last_k, first_k] if x])

        # 会社
        org = pick(row_norm, "orgname")
        org_kana = company_name_kana(org)
        dept = pick(row_norm, "orgdept")
        title = pick(row_norm, "orgtitle")
        nickname = pick(row_norm, "nickname")
        birthday = pick(row_norm, "birthday")
        notes = pick(row_norm, "notes")  # 備考1 に入れる

        # メモ（日本語列がある場合のみ拾う）
        memo1 = raw_row.get("メモ1", "")
        memo2 = raw_row.get("メモ2", "")
        memo3 = raw_row.get("メモ3", "")
        memo4 = raw_row.get("メモ4", "")
        memo5 = raw_row.get("メモ5", "")

        # アドレス
        street = pick(row_norm, "addr_street")
        zip_code = pick(row_norm, "addr_zip")
        addr1, addr2 = normalize_address(street)

        # 連絡先
        phones = ";".join([pick(row_norm, "phone1"), pick(row_norm, "phone2"), pick(row_norm, "phone3")]).strip(";")
        emails_company = ";".join([pick(row_norm, "email1"), pick(row_norm, "email2")]).strip(";")
        email_home = pick(row_norm, "email3")

        w.writerow([
            last, first, last_k, first_k, full, full_k, mid, mid_k,
            "様", nickname, "", "会社", "", "", "", "",
            "", "", email_home, "", "",
            zip_code, addr1, addr2, "", phones, "", emails_company, "", "",
            "", "", "", "", "", "", "", "", "",
            org_kana, org, dept, "", title,
            "", "", "", "",
            memo1, memo2, memo3, memo4, memo5, notes, "", "",
            birthday, "選択なし", "選択なし", "", ""
        ])

    out.seek(0)
    fn = f"converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.tsv"
    return send_file(io.BytesIO(out.getvalue().encode("utf-8-sig")),
                     mimetype="text/tab-separated-values",
                     as_attachment=True,
                     download_name=fn)

if __name__ == "__main__":
    app.run(debug=True)
