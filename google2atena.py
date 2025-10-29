# google2atena.py
# v3.9.18r (robust CSV版 / import済 / no-pandas)

from flask import Flask, request, render_template, send_file
import csv
import io
import re
import unicodedata
from datetime import datetime
import chardet

# 外部辞書
from company_dicts import COMPANY_EXCEPT
from kanji_word_map import KANJI_WORD_MAP
from corp_terms import CORP_TERMS
from jp_area_codes import AREA_CODES

app = Flask(__name__)

# ===============================
# ファイル入力部（堅牢化）
# ===============================
def load_csv(file):
    raw = file.stream.read()
    detect = chardet.detect(raw)
    encoding = detect["encoding"] or "utf-8-sig"

    sample_text = raw.decode(encoding, errors="ignore")
    first_line = sample_text.splitlines()[0]
    delimiter = "\t" if "\t" in first_line else ","

    # csv.DictReader に渡すストリーム生成
    stream = io.StringIO(sample_text)
    reader = csv.DictReader(stream, delimiter=delimiter)
    return reader

def normalize_keys(row):
    """列名を正規化（全角→半角、スペース除去）"""
    return {
        unicodedata.normalize("NFKC", k).replace("－", "-").replace("–", "-").strip(): v
        for k, v in row.items()
    }

# ===============================
# 住所分割
# ===============================
def split_building(address):
    if not address:
        return "", ""
    addr = address.strip().replace("\r", "")
    parts = addr.split(" ", 1)
    if len(parts) == 2:
        return parts[0], parts[1]
    keywords = ["ビル","マンション","ハイツ","コーポ","レジデンス","タワー","ヒルズ","メゾン","センター","プラザ","アネックス","ガーデン"]
    for kw in keywords:
        i = addr.find(kw)
        if i != -1:
            return addr[:i+len(kw)], addr[i+len(kw):].strip()
    return addr, ""

# ===============================
# 全角化（郵便番号・電話除く）
# ===============================
def _zenkaku_text(text):
    if not text:
        return ""
    out = ""
    for c in text:
        if re.match(r"[A-Za-z0-9\-]", c):
            out += chr(ord(c) + 0xFEE0)
        else:
            out += c
    return out

# ===============================
# 電話番号整形
# ===============================
def _normalize_number(num):
    num = unicodedata.normalize("NFKC", str(num))
    return re.sub(r"[^\d]", "", num)

def _format_phone(num):
    num = _normalize_number(num)
    if not num.startswith("0") and len(num) >= 9:
        num = "0" + num
    for code in sorted(AREA_CODES, key=len, reverse=True):
        if num.startswith(code):
            rest = num[len(code):]
            if len(rest) > 4:
                return f"{code}-{rest[:-4]}-{rest[-4:]}"
            else:
                return f"{code}-{rest}"
    return num

def _format_phone_numbers(raw):
    if not raw:
        return ""
    raw = unicodedata.normalize("NFKC", raw)
    parts = re.split(r"[;,:／／／\s]+|:::+", raw)
    cleaned = []
    for p in parts:
        n = _normalize_number(p)
        if n:
            formatted = _format_phone(n)
            if formatted not in cleaned:
                cleaned.append(formatted)
    return ";".join(cleaned)

# ===============================
# メール整形
# ===============================
def _normalize_emails(*emails):
    all_mails = []
    for e in emails:
        e = unicodedata.normalize("NFKC", str(e)).strip()
        if e and "@" in e:
            e = e.lower()
            if e not in all_mails:
                all_mails.append(e)
    return ";".join(all_mails)

# ===============================
# 誕生日整形
# ===============================
def _format_birthday(bday):
    if not bday:
        return ""
    try:
        dt = datetime.strptime(bday.strip(), "%Y-%m-%d")
        return dt.strftime("%Y/%m/%d")
    except Exception:
        return bday

# ===============================
# 会社名かな変換
# ===============================
def _company_to_kana(name):
    if not name:
        return ""
    name = re.sub("|".join(map(re.escape, CORP_TERMS)), "", name)
    name = name.replace("　", "").replace(" ", "")
    if name in COMPANY_EXCEPT:
        kana = COMPANY_EXCEPT[name]
    else:
        kana = name
        for k, v in KANJI_WORD_MAP.items():
            kana = kana.replace(k, v)
        kana = unicodedata.normalize("NFKC", kana)
    kana = re.sub(r"[・.,，]", "", kana)
    kana = re.sub(r"[A-Za-z]", lambda m: chr(ord(m.group(0)) + 0xFEE0), kana)
    return kana

# ===============================
# Flask メイン
# ===============================
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if not file:
            return "⚠️ ファイルをアップロードしてください。"

        try:
            reader = load_csv(file)
        except Exception as e:
            return f"⚠️ エラーが発生しました。CSVの形式や文字コードをご確認ください。\n{e}"

        output = io.StringIO()
        fieldnames = [
            "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称",
            "ニックネーム","旧姓","宛先","自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話",
            "自宅IM ID","自宅E-mail","自宅URL","自宅Social","会社〒","会社住所1","会社住所2","会社住所3",
            "会社電話","会社IM ID","会社E-mail","会社URL","会社Social","その他〒","その他住所1","その他住所2",
            "その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social","会社名かな",
            "会社名","部署名1","部署名2","役職名","連名","連名ふりがな","連名敬称","連名誕生日",
            "メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
        ]
        writer = csv.DictWriter(output, fieldnames=fieldnames)
        writer.writeheader()

        for row_raw in reader:
            row = normalize_keys(row_raw)

            last, first = row.get("Last Name",""), row.get("First Name","")
            last_kana, first_kana = row.get("Phonetic Last Name",""), row.get("Phonetic First Name","")
            full, full_kana = f"{last}　{first}", f"{last_kana}　{first_kana}"

            org = row.get("Organization Name","")
            org_kana = _company_to_kana(org)
            birthday = _format_birthday(row.get("Birthday",""))

            label = row.get("Address 1 - Label","")
            addr = row.get("Address 1 - Formatted","")
            home_zip = other_zip = work_zip = ""
            home_a1 = other_a1 = work_a1 = ""
            home_a2 = other_a2 = work_a2 = ""

            if addr:
                addr = addr.replace("\r", "").strip()
                zip_match = re.search(r"\d{3}-?\d{4}", addr)
                zip_code = zip_match.group(0) if zip_match else ""
                addr = addr.replace(zip_code, "")
                lines = addr.split("\n")
                addr_core = lines[0] if lines else addr
                a1, a2 = split_building(addr_core)
                a1, a2 = _zenkaku_text(a1), _zenkaku_text(a2)
                if "Home" in label:
                    home_zip, home_a1, home_a2 = zip_code, a1, a2
                elif "Other" in label:
                    other_zip, other_a1, other_a2 = zip_code, a1, a2
                else:
                    work_zip, work_a1, work_a2 = zip_code, a1, a2

            phones = _format_phone_numbers(
                (row.get("Phone 1 - Value","") or "") + ";" + (row.get("Phone 2 - Value","") or "")
            )
            emails = _normalize_emails(
                row.get("E-mail 1 - Value",""), row.get("E-mail 2 - Value",""), row.get("E-mail 3 - Value","")
            )

            notes = [""] * 5
            for k,v in row.items():
                if re.match(r"メモ ?[1-5１-５]|memo ?[1-5１-５]", k, re.IGNORECASE):
                    num = int(re.findall(r"[1-5１-５]", k)[0])
                    notes[num-1] = v
                elif k.lower() == "notes":
                    notes[0] = v

            writer.writerow({
                "姓": last, "名": first, "姓かな": last_kana, "名かな": first_kana,
                "姓名": full, "姓名かな": full_kana, "敬称": "様",
                "宛先": "会社",
                "自宅〒": home_zip, "自宅住所1": home_a1, "自宅住所2": home_a2,
                "会社〒": work_zip, "会社住所1": work_a1, "会社住所2": work_a2,
                "会社電話": phones, "会社E-mail": emails,
                "会社名": org, "会社名かな": org_kana,
                "部署名1": row.get("Organization Department",""),
                "役職名": row.get("Organization Title",""),
                "メモ1": notes[0], "メモ2": notes[1], "メモ3": notes[2],
                "メモ4": notes[3], "メモ5": notes[4],
                "誕生日": birthday
            })

        mem = io.BytesIO()
        mem.write(output.getvalue().encode("utf-8-sig"))
        mem.seek(0)
        return send_file(mem, as_attachment=True, download_name="converted.csv", mimetype="text/csv")

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
