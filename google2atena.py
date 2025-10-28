# google2atena.py
from flask import Flask, request, render_template_string, send_file
import csv
import io
import re

app = Flask(__name__)

HTML = """
<!doctype html>
<title>Google連絡先CSV → 宛名職人CSV 変換（v3.9.8 full-format no-pandas final）</title>
<h2>Google連絡先CSV → 宛名職人CSV 変換（v3.9.8 full-format no-pandas final）</h2>
<form method=post enctype=multipart/form-data>
  <p><input type=file name=file>
     <input type=submit value="変換開始">
</form>
{% if download_link %}
  <p><a href="{{ download_link }}">変換後CSVをダウンロード</a></p>
{% endif %}
"""

# =====================
# 正規化ヘルパー群
# =====================

def normalize_phone(phone):
    if not phone:
        return ""
    phone = re.sub(r"[^\d]", "", str(phone))
    if len(phone) == 10 and phone.startswith("0"):
        return f"{phone[0:2]}-{phone[2:6]}-{phone[6:]}"
    elif len(phone) == 11 and phone.startswith("0"):
        return f"{phone[0:3]}-{phone[3:7]}-{phone[7:]}"
    return phone

def normalize_label(s):
    if not s:
        return ""
    s = str(s).strip()
    table = str.maketrans({
        "①": "1", "②": "2", "③": "3", "④": "4", "⑤": "5",
        "１": "1", "２": "2", "３": "3", "４": "4", "５": "5"
    })
    s = s.translate(table)
    s = re.sub(r"\s+", "", s)
    return s

def kana_company_name(name):
    if not name:
        return ""
    s = str(name)
    s = re.sub(r"(株式会社|有限会社|合同会社|社団法人|財団法人|医療法人|学校法人)", "", s)
    s = s.strip("　 ").replace("　", " ")

    eng_map = {
        "A": "エー", "B": "ビー", "C": "シー", "D": "ディー", "E": "イー", "F": "エフ",
        "G": "ジー", "H": "エイチ", "I": "アイ", "J": "ジェー", "K": "ケー", "L": "エル",
        "M": "エム", "N": "エヌ", "O": "オー", "P": "ピー", "Q": "キュー", "R": "アール",
        "S": "エス", "T": "ティー", "U": "ユー", "V": "ブイ", "W": "ダブリュー",
        "X": "エックス", "Y": "ワイ", "Z": "ゼット",
    }
    s = "".join(eng_map.get(ch.upper(), ch) for ch in s)
    s = "".join(chr(ord(ch) + 96) if "ぁ" <= ch <= "ん" else ch for ch in s)
    return s

def to_fullwidth(text):
    if not text:
        return ""
    s = str(text)
    fw_table = str.maketrans({
        " ": "　",
        "-": "－",
        "0": "０", "1": "１", "2": "２", "3": "３", "4": "４",
        "5": "５", "6": "６", "7": "７", "8": "８", "9": "９",
        "A": "Ａ", "B": "Ｂ", "C": "Ｃ", "D": "Ｄ", "E": "Ｅ", "F": "Ｆ",
        "G": "Ｇ", "H": "Ｈ", "I": "Ｉ", "J": "Ｊ", "K": "Ｋ", "L": "Ｌ",
        "M": "Ｍ", "N": "Ｎ", "O": "Ｏ", "P": "Ｐ", "Q": "Ｑ", "R": "Ｒ",
        "S": "Ｓ", "T": "Ｔ", "U": "Ｕ", "V": "Ｖ", "W": "Ｗ",
        "X": "Ｘ", "Y": "Ｙ", "Z": "Ｚ",
        "a": "ａ", "b": "ｂ", "c": "ｃ", "d": "ｄ", "e": "ｅ", "f": "ｆ",
        "g": "ｇ", "h": "ｈ", "i": "ｉ", "j": "ｊ", "k": "ｋ", "l": "ｌ",
        "m": "ｍ", "n": "ｎ", "o": "ｏ", "p": "ｐ", "q": "ｑ", "r": "ｒ",
        "s": "ｓ", "t": "ｔ", "u": "ｕ", "v": "ｖ", "w": "ｗ",
        "x": "ｘ", "y": "ｙ", "z": "ｚ"
    })
    return s.translate(fw_table)

def split_address(addr):
    if not addr:
        return "", "", ""
    addr = addr.strip().replace("　", " ")
    parts = re.split(r"[\n\r]+", addr)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) == 3:
        return parts[1], parts[0], parts[2]
    elif len(parts) == 2:
        return parts[1], parts[0], ""
    elif len(parts) == 1:
        return parts[0], "", ""
    else:
        return "", "", ""

def collect_memo_and_notes(row):
    memos = {"メモ1": "", "メモ2": "", "メモ3": "", "メモ4": "", "メモ5": ""}
    notes = row.get("Notes", "")

    for k, v in row.items():
        if not v:
            continue
        lbl = normalize_label(k)
        val = str(v).strip()
        for i in range(1, 6):
            if lbl in [f"メモ{i}", f"memo{i}"]:
                memos[f"メモ{i}"] = val
    return memos, notes

# =====================
# メイン変換ロジック
# =====================
def convert_google_to_atena(text):
    reader = csv.DictReader(io.StringIO(text))
    rows = []

    for r in reader:
        first = r.get("First Name", "")
        last = r.get("Last Name", "")
        first_kana = r.get("Phonetic First Name", "")
        last_kana = r.get("Phonetic Last Name", "")
        org = r.get("Organization Name", "")
        dept = r.get("Organization Department", "")
        title = r.get("Organization Title", "")
        zip_code = (r.get("Address 1 - Postal Code", "") or "").strip()
        region = r.get("Address 1 - Region", "")
        addr_full = r.get("Address 1 - Formatted", "") or r.get("Address 1 - Street", "")
        addr1, addr2, addr3 = split_address(addr_full)
        addr1 = (region or "") + (addr1 or "")

        emails = [r.get(f"E-mail {i} - Value", "") for i in range(1, 6) if r.get(f"E-mail {i} - Value", "")]
        email_str = ";".join(emails)

        phones = [normalize_phone(r.get(f"Phone {i} - Value", "")) for i in range(1, 6) if r.get(f"Phone {i} - Value", "")]
        phone_str = ";".join([p for p in phones if p])

        memos, note = collect_memo_and_notes(r)

        rows.append({
            "姓": last,
            "名": first,
            "姓かな": last_kana,
            "名かな": first_kana,
            "姓名": f"{last}　{first}".strip(),
            "姓名かな": f"{last_kana}　{first_kana}".strip(),
            "敬称": "様",
            "宛先": "会社",
            "会社〒": zip_code,
            "会社住所1": addr1,
            "会社住所2": addr2,
            "会社住所3": addr3,
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
        })
    return rows

# =====================
# Flaskルート
# =====================
@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        file = request.files["file"]
        if not file:
            return render_template_string(HTML, download_link=None)

        text = file.read().decode("utf-8", errors="ignore")
        rows = convert_google_to_atena(text)

        output = io.StringIO()
        fieldnames = list(rows[0].keys()) if rows else []
        writer = csv.DictWriter(output, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
        output.seek(0)

        return send_file(io.BytesIO(output.getvalue().encode("utf-8-sig")),
                         as_attachment=True,
                         download_name="google_converted.csv",
                         mimetype="text/csv")
    return render_template_string(HTML, download_link=None)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
