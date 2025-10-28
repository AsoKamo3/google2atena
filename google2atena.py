from flask import Flask, request, render_template_string, send_file
import pandas as pd
import io
import re

app = Flask(__name__)

HTML = """
<!doctype html>
<title>Google連絡先CSV → 宛名職人CSV 変換（v3.8.2 safe）</title>
<h2>Google連絡先CSV → 宛名職人CSV 変換（v3.8.2 safe）</h2>
<form method=post enctype=multipart/form-data>
  <p><input type=file name=file>
     <input type=submit value="変換開始">
</form>
{% if download_link %}
  <p><a href="{{ download_link }}">変換後CSVをダウンロード</a></p>
{% endif %}
"""

# ======== 電話番号整形（半角保持） ========
def normalize_phone(phone):
    if not phone or pd.isna(phone):
        return ""
    phone = str(phone)
    phone = re.sub(r"[^\d]", "", phone)  # 数字以外を除去
    # 国内番号: 10桁 or 11桁
    if len(phone) == 10:
        phone = f"{phone[0:2]}-{phone[2:6]}-{phone[6:]}" if phone.startswith("0") else phone
    elif len(phone) == 11:
        phone = f"{phone[0:3]}-{phone[3:7]}-{phone[7:]}" if phone.startswith("0") else phone
    return phone

# ======== 住所分割（単純な最初のスペース区切り＋改行対応） ========
def split_address(addr):
    if not addr or pd.isna(addr):
        return "", "", ""
    addr = str(addr).replace("　", " ").strip()
    if " " in addr:
        parts = addr.split(" ", 1)
        return parts[0], parts[1], ""
    else:
        parts = re.split(r"[\n\r]", addr)
        parts = [p.strip() for p in parts if p.strip()]
        if len(parts) == 3:
            return parts[1], parts[0], parts[2]
        elif len(parts) == 2:
            return parts[1], parts[0], ""
        else:
            return addr, "", ""

# ======== メイン変換 ========
def convert_google_to_atena(df):
    rows = []
    for _, r in df.iterrows():
        first = r.get("First Name", "")
        last = r.get("Last Name", "")
        first_kana = r.get("Phonetic First Name", "")
        last_kana = r.get("Phonetic Last Name", "")
        org = r.get("Organization Name", "")
        dept = r.get("Organization Department", "")
        title = r.get("Organization Title", "")
        note = r.get("Notes", "")

        # メール抽出（NaN対策）
        emails = []
        for i in range(1, 6):
            v = r.get(f"E-mail {i} - Value", "")
            if pd.notna(v) and str(v).strip() != "":
                emails.append(str(v).strip())
        email_str = ";".join(emails)

        # 電話番号抽出（Work > Mobile > Home）
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

        # 住所分割
        addr_full = r.get("Address 1 - Street", "") or r.get("Address 1 - Formatted", "")
        zip_code = str(r.get("Address 1 - Postal Code", "")).replace("-", "－")
        region = str(r.get("Address 1 - Region", "")).replace("-", "－")
        addr1, addr2, addr3 = split_address(addr_full)

        # メモ関連（NaN回避）
        memos = {f"メモ{i}": "" for i in range(1, 6)}
        for k in r.keys():
            if "メモ" in str(k):
                for i in range(1, 6):
                    if f"メモ{i}" in k:
                        v = r[k]
                        if pd.notna(v):
                            memos[f"メモ{i}"] = str(v)

        rows.append({
            "姓": last, "名": first, "姓かな": last_kana, "名かな": first_kana,
            "姓名": f"{last}　{first}".strip(), "姓名かな": f"{last_kana}　{first_kana}".strip(),
            "敬称": "様", "宛先": "会社",
            "会社〒": zip_code, "会社住所1": region + addr1, "会社住所2": addr2, "会社住所3": addr3,
            "会社電話1〜10": phone_str,
            "会社E-mail1〜5": email_str,
            "会社名": org, "部署名1": dept, "役職名": title,
            "メモ1": memos["メモ1"], "メモ2": memos["メモ2"], "メモ3": memos["メモ3"],
            "メモ4": memos["メモ4"], "メモ5": memos["メモ5"],
            "備考1": note, "誕生日": r.get("Birthday", ""),
        })
    return pd.DataFrame(rows)

# ======== Flaskルート ========
@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        file = request.files["file"]
        if not file:
            return render_template_string(HTML, download_link=None)

        # UTF-8優先、他エンコーディングは無視
        df = pd.read_csv(file, encoding="utf-8", errors="ignore")
        result = convert_google_to_atena(df)

        buf = io.BytesIO()
        result.to_csv(buf, index=False, encoding="utf-8-sig")
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name="google_converted.csv", mimetype="text/csv")

    return render_template_string(HTML, download_link=None)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
