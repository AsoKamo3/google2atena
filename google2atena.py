# google2atena.py  — v3.9.9 full-format no-pandas（統合・安定）
from flask import Flask, request, render_template_string, send_file
import csv
import io
import re

app = Flask(__name__)

HTML = """
<!doctype html>
<title>Google連絡先CSV → 宛名職人CSV 変換（v3.9.9 full-format no-pandas）</title>
<h2>Google連絡先CSV → 宛名職人CSV 変換（v3.9.9 full-format no-pandas）</h2>
<form method=post enctype=multipart/form-data>
  <p><input type=file name=file required>
     <input type=submit value="変換開始">
</form>
"""

# --------------------------
# 正規化ヘルパー
# --------------------------

def normalize_phone(phone):
    if not phone:
        return ""
    phone = re.sub(r"[^\d]", "", str(phone))
    # 先頭0が無い9桁ケースは0を付与（例: 112615331 → 0112615331）
    if len(phone) == 9 and not phone.startswith("0"):
        phone = "0" + phone
    # 10桁（03系は 2-4-4、それ以外は 3-3-4 を優先…など完全NTT準拠ではない簡易版）
    if len(phone) == 10 and phone.startswith("0"):
        if phone.startswith(("03", "06")):
            return f"{phone[0:2]}-{phone[2:6]}-{phone[6:]}"
        else:
            return f"{phone[0:3]}-{phone[3:6]}-{phone[6:]}"
    # 11桁は 3-4-4
    if len(phone) == 11 and phone.startswith("0"):
        return f"{phone[0:3]}-{phone[3:7]}-{phone[7:]}"
    return phone

def normalize_label(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip()
    # 丸数字/全角数字 → 半角
    table = str.maketrans({
        "①":"1","②":"2","③":"3","④":"4","⑤":"5",
        "１":"1","２":"2","３":"3","４":"4","５":"5",
        "　":"", " ":""
    })
    s = s.translate(table)
    # 大文字小文字/全角英字も吸収
    s = s.lower()
    s = s.replace("memo", "メモ")
    return s  # 例: "メモ1", "メモ2", ...

def to_fullwidth(text):
    if not text:
        return ""
    s = str(text)
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
    return s.translate(fw)

def split_address(addr):
    """Googleの Address 1 - Formatted など改行区切りを
       [会社住所1=都道府県市区町村 等, 会社住所2=番地・建物等, 会社住所3=（予備）] へ分割"""
    if not addr:
        return "", "", ""
    addr = addr.strip()
    parts = [p.strip() for p in re.split(r"[\r\n]+", addr) if p.strip()]
    # よくある 4～5行の塊: [通り, 市区町村, 都道府県, 郵便番号, 日本]
    # ここでは「都道府県行 + 通り」を優先して住所1/2を作る
    if len(parts) >= 3:
        street = parts[0]
        region = parts[2]  # 都道府県
        addr1 = f"{region}{street}"
        addr2 = ""
        if " " in street:
            # 稀なパターンはここで微調整（必要な場合のみ）
            pass
        return addr1, "", ""
    elif len(parts) == 2:
        return parts[1], parts[0], ""
    elif len(parts) == 1:
        return parts[0], "", ""
    return "", "", ""

def kana_company_name(name):
    """法人格を除去し、英字は読み上げカタカナ化（簡易）。漢字→カタカナは外部辞書なし簡易のため保持。"""
    if not name:
        return ""
    s = str(name)
    # 法人格カット
    s = re.sub(r"(株式会社|有限会社|合同会社|社団法人|財団法人|医療法人|学校法人)", "", s)
    s = s.strip()

    # 英字 → 読み上げカタカナ
    roman = {
        "A":"エー","B":"ビー","C":"シー","D":"ディー","E":"イー","F":"エフ",
        "G":"ジー","H":"エイチ","I":"アイ","J":"ジェー","K":"ケー","L":"エル",
        "M":"エム","N":"エヌ","O":"オー","P":"ピー","Q":"キュー","R":"アール",
        "S":"エス","T":"ティー","U":"ユー","V":"ブイ","W":"ダブリュー",
        "X":"エックス","Y":"ワイ","Z":"ゼット"
    }
    out = []
    for ch in s:
        if re.match(r"[A-Za-z]", ch):
            out.append(roman[ch.upper()])
        else:
            out.append(ch)
    return "".join(out)

# --------------------------
# メモ/Notes 収集
# --------------------------
def collect_memo_and_notes(row: dict):
    """Relation n - Label/Value から メモ1..5 を抽出。併存の素カラム「メモ1..5」「memo 1..5」も拾う。"""
    memos = {f"メモ{i}": "" for i in range(1, 6)}
    notes = row.get("Notes", "") or ""

    # 1) Relation系を優先
    for i in range(1, 6):
        lbl = normalize_label(row.get(f"Relation {i} - Label", ""))
        val = (row.get(f"Relation {i} - Value", "") or "").strip()
        if not val:
            continue
        if lbl == f"メモ{i}":
            memos[f"メモ{i}"] = val

    # 2) 行内に素の「メモ」「memo」カラムがある場合のフォールバック
    for k, v in row.items():
        if not v:
            continue
        lbl = normalize_label(k)
        m = re.fullmatch(r"メモ([1-5])", lbl)
        if m and not memos[f"メモ{m.group(1)}"]:
            memos[f"メモ{m.group(1)}"] = str(v).strip()

    return memos, notes

# --------------------------
# 変換本体
# --------------------------
HEADER = [
    "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称","ニックネーム","旧姓","宛先",
    "自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
    "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social",
    "その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
    "会社名かな","会社名","部署名1","部署名2","役職名","連名","連名ふりがな","連名敬称","連名誕生日",
    "メモ1","メモ2","メモ3","メモ4","メモ5",
    "備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
]

def convert_row(r: dict) -> dict:
    first = r.get("First Name", "") or ""
    last = r.get("Last Name", "") or ""
    first_kana = r.get("Phonetic First Name", "") or ""
    last_kana = r.get("Phonetic Last Name", "") or ""
    nick = r.get("Nickname", "") or ""
    org = r.get("Organization Name", "") or ""
    dept = r.get("Organization Department", "") or ""
    title = r.get("Organization Title", "") or ""
    birthday = r.get("Birthday", "") or ""

    # 住所
    zip_code = (r.get("Address 1 - Postal Code", "") or "").strip()  # 郵便番号は半角維持
    region = r.get("Address 1 - Region", "") or ""
    formatted = r.get("Address 1 - Formatted", "") or r.get("Address 1 - Street", "") or ""
    a1, a2, a3 = split_address(formatted)
    # 会社住所1は「都道府県＋通り」の合成を優先
    if region and a1 and not a1.startswith(region):
        a1 = f"{region}{a1}"
    # 会社住所1/2 は全角化
    a1 = to_fullwidth(a1)
    a2 = to_fullwidth(a2)

    # 電話（会社側へ集約）
    phones = []
    for i in range(1, 6):
        v = r.get(f"Phone {i} - Value", "")
        if v:
            p = normalize_phone(v)
            if p:
                phones.append(p)
    phone_str = ";".join(phones)

    # メール（会社側へ集約）
    emails = []
    for i in range(1, 6):
        v = r.get(f"E-mail {i} - Value", "")
        if v:
            emails.append(v.strip())
    email_str = ";".join(emails)

    # メモとNotes
    memos, note = collect_memo_and_notes(r)

    row = {h: "" for h in HEADER}
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
        "会社住所1": a1,
        "会社住所2": a2,
        "会社住所3": a3,
        "会社電話": phone_str,
        "会社E-mail": email_str,
        "会社名かな": kana_company_name(org),
        "会社名": org,
        "部署名1": to_fullwidth(dept),
        "役職名": to_fullwidth(title),
        "メモ1": memos["メモ1"],
        "メモ2": memos["メモ2"],
        "メモ3": memos["メモ3"],
        "メモ4": memos["メモ4"],
        "メモ5": memos["メモ5"],
        "備考1": note,
        "誕生日": birthday
    })
    return row

def convert_google_to_atena(text: str):
    reader = csv.DictReader(io.StringIO(text))
    return [convert_row(r) for r in reader]

# --------------------------
# Flask ルート
# --------------------------
@app.route("/", methods=["GET", "POST"])
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
