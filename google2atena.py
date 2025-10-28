from flask import Flask, request, render_template_string, send_file
import pandas as pd
import io, re

app = Flask(__name__)

HTML = """
<!doctype html>
<title>Google連絡先CSV → 宛名職人CSV 変換（v3.9 full fix（phone-format + address/postal fix 統合版））</title>
<h2>Google連絡先CSV → 宛名職人CSV 変換（v3.9 full fix：電話＋住所＋会社名かな）</h2>
<form method=post enctype=multipart/form-data>
  <p><input type=file name=file>
     <input type=submit value="変換開始">
</form>
{% if download_link %}
  <p><a href="{{ download_link }}">変換後CSVをダウンロード</a></p>
{% endif %}
"""

# ==============================================================
# Utility functions
# ==============================================================

def normalize_phone(phone):
    """電話番号：数字抽出→半角→ハイフン挿入"""
    if not phone or pd.isna(phone): return ""
    phone = re.sub(r"[^\d]", "", str(phone))
    if len(phone) == 10:
        return f"{phone[:2]}-{phone[2:6]}-{phone[6:]}"
    elif len(phone) == 11:
        return f"{phone[:3]}-{phone[3:7]}-{phone[7:]}"
    return phone

def to_zenkaku(s):
    """数字・英字・記号を全角に"""
    if not s: return ""
    return s.translate(str.maketrans(
        ''.join(chr(i) for i in range(33, 127)),
        ''.join(chr(i+0xFEE0) for i in range(33, 127))
    ))

def to_katakana(s):
    """会社名かな生成（ひらがな→カタカナ、英字→全角）"""
    if not s: return ""
    s = re.sub(r"^(株式会社|有限会社|合同会社|一般社団法人|公益社団法人|社団法人|財団法人)\s*", "", s)
    hira = str.maketrans("ぁ-ん", "ァ-ン")
    s = s.translate(hira)
    s = to_zenkaku(s)
    replacements = {
        "LAB": "ラボ", "WORKS": "ワークス", "OFFICE": "オフィス", "NHK": "エヌエイチケー",
        "KADOKAWA": "カドカワ", "STAND": "スタンド", "NEO": "ネオ", "REAL": "リアル",
        "MARUZEN": "マルゼン", "YADOKARI": "ヤドカリ", "TOI": "トイ", "PLAN": "プラン",
        "ALL": "オール", "REVIEWS": "レビューズ", "COUNTER": "カウンター", "ODD": "オッド"
    }
    for k, v in replacements.items():
        s = re.sub(k, v, s, flags=re.IGNORECASE)
    return s

def format_company_name(name):
    """株式会社などの後に全角スペース追加"""
    if not name: return ""
    name = re.sub(r"^(株式会社|有限会社|合同会社|一般社団法人|社団法人|財団法人)(?=[^\s　])", r"\1　", name)
    return to_zenkaku(name)

def split_address(addr):
    """最初のスペース以降を建物名として扱う簡易分割"""
    if not addr or pd.isna(addr): return "", ""
    addr = str(addr).strip().replace("　", " ")
    if " " in addr:
        parts = addr.split(" ", 1)
        return to_zenkaku(parts[0]), to_zenkaku(parts[1])
    return to_zenkaku(addr), ""

def normalize_zip(zipcode):
    """郵便番号を全角化"""
    if not zipcode: return ""
    zipcode = str(zipcode).replace("-", "－")
    return to_zenkaku(zipcode)

# ==============================================================
# Main conversion logic
# ==============================================================

def convert_google_to_atena(df):
    rows = []
    for _, r in df.iterrows():
        first, last = r.get("First Name", ""), r.get("Last Name", "")
        first_kana, last_kana = r.get("Phonetic First Name", ""), r.get("Phonetic Last Name", "")
        org, dept, title = r.get("Organization Name", ""), r.get("Organization Department", ""), r.get("Organization Title", "")
        note = r.get("Notes", "")

        # --- メール ---
        emails = []
        for i in range(1, 6):
            v = r.get(f"E-mail {i} - Value", "")
            if v and v not in emails:
                emails.append(v)
        email_str = ";".join(emails)

        # --- 電話 ---
        phone_dict = {}
        for i in range(1, 6):
            label = str(r.get(f"Phone {i} - Label", "")).lower()
            value = normalize_phone(r.get(f"Phone {i} - Value", ""))
            if not value:
                continue
            if "work" in label:
                phone_dict.setdefault("work", []).append(value)
            elif "mobile" in label:
                phone_dict.setdefault("mobile", []).append(value)
            elif "home" in label:
                phone_dict.setdefault("home", []).append(value)
        phones = []
        for k in ["work", "mobile", "home"]:
            if k in phone_dict:
                phones.extend(phone_dict[k])
        phone_str = ";".join(sorted(set(phones)))

        # --- 住所 ---
        addr_full = r.get("Address 1 - Street", "") or r.get("Address 1 - Formatted", "")
        addr1, addr2 = split_address(addr_full)
        zip_code = normalize_zip(r.get("Address 1 - Postal Code", ""))
        region = to_zenkaku(r.get("Address 1 - Region", ""))

        # --- メモ ---
        memos = {}
        for i in range(1, 6):
            memos[f"メモ{i}"] = ""
        for k in r.keys():
            if "メモ" in str(k):
                for i in range(1, 6):
                    if f"メモ{i}" in k:
                        memos[f"メモ{i}"] = str(r[k])

        # --- 会社名 ---
        formatted_org = format_company_name(org)
        org_kana = to_katakana(org)

        # --- 行構築 ---
        rows.append({
            "姓": last, "名": first,
            "姓かな": last_kana, "名かな": first_kana,
            "姓名": f"{last}　{first}", "姓名かな": f"{last_kana}　{first_kana}",
            "ミドルネーム": "", "ミドルネームかな": "",
            "敬称": "様", "ニックネーム": "", "旧姓": "",
            "宛先": "会社",
            "自宅〒": "", "自宅住所1": "", "自宅住所2": "", "自宅住所3": "",
            "自宅電話1〜10": "", "自宅IM ID1〜10": "", "自宅E-mail1〜5": "", "自宅URL1〜5": "", "自宅Social1〜10": "",
            "会社〒": zip_code,
            "会社住所1": region + addr1, "会社住所2": addr2, "会社住所3": "",
            "会社電話1〜10": phone_str, "会社IM ID1〜10": "", "会社E-mail1〜5": email_str,
            "会社URL1〜5": "", "会社Social1〜10": "",
            "その他〒": "", "その他住所1": "", "その他住所2": "", "その他住所3": "",
            "その他電話1〜10": "", "その他IM ID1〜10": "", "その他E-mail1〜5": "", "その他URL1〜5": "", "その他Social1〜10": "",
            "会社名かな": org_kana, "会社名": formatted_org, "部署名1": dept, "部署名2": "", "役職名": title,
            "連名1〜20": "", "連名ふりがな1〜20": "", "連名敬称1〜20": "", "連名誕生日1〜20": "",
            "メモ1": memos["メモ1"], "メモ2": memos["メモ2"], "メモ3": memos["メモ3"], "メモ4": memos["メモ4"], "メモ5": memos["メモ5"],
            "備考1": note, "備考2": "", "備考3": "",
            "誕生日": r.get("Birthday", ""), "性別": "選択なし", "血液型": "選択なし", "趣味": "", "性格": ""
        })
    return pd.DataFrame(rows)

# ==============================================================
# Flask routes
# ==============================================================

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        file = request.files["file"]
        if not file:
            return render_template_string(HTML, download_link=None)

        df = pd.read_csv(file)
        result = convert_google_to_atena(df)

        buf = io.BytesIO()
        result.to_csv(buf, index=False, encoding="utf-8-sig")
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name="google_converted.csv", mimetype="text/csv")

    return render_template_string(HTML, download_link=None)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
