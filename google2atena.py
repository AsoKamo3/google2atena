from flask import Flask, request, render_template_string, send_file
import csv, io, re

app = Flask(__name__)

HTML = """
<!doctype html>
<title>Google連絡先CSV → 宛名職人CSV 変換（v3.9.1 full-fix no-pandas）</title>
<h2>Google連絡先CSV → 宛名職人CSV 変換（v3.9.1 full-fix no-pandas）</h2>
<form method=post enctype=multipart/form-data>
  <p><input type=file name=file>
     <input type=submit value="変換開始">
</form>
"""

# ========= 文字列全角化 =========
def to_zenkaku(s):
    if not s:
        return ""
    table = str.maketrans({
        "0":"０","1":"１","2":"２","3":"３","4":"４","5":"５","6":"６","7":"７","8":"８","9":"９",
        "-":"－","#":"＃"," ":"　",
        "a":"ａ","b":"ｂ","c":"ｃ","d":"ｄ","e":"ｅ","f":"ｆ","g":"ｇ","h":"ｈ","i":"ｉ","j":"ｊ","k":"ｋ","l":"ｌ","m":"ｍ","n":"ｎ","o":"ｏ","p":"ｐ","q":"ｑ","r":"ｒ","s":"ｓ","t":"ｔ","u":"ｕ","v":"ｖ","w":"ｗ","x":"ｘ","y":"ｙ","z":"ｚ",
        "A":"Ａ","B":"Ｂ","C":"Ｃ","D":"Ｄ","E":"Ｅ","F":"Ｆ","G":"Ｇ","H":"Ｈ","I":"Ｉ","J":"Ｊ","K":"Ｋ","L":"Ｌ","M":"Ｍ","N":"Ｎ","O":"Ｏ","P":"Ｐ","Q":"Ｑ","R":"Ｒ","S":"Ｓ","T":"Ｔ","U":"Ｕ","V":"Ｖ","W":"Ｗ","X":"Ｘ","Y":"Ｙ","Z":"Ｚ",
    })
    return s.translate(table)

# ========= 電話番号整形（NTT形式補完） =========
def normalize_phone(phone):
    if not phone:
        return ""
    # 区切り文字を分割
    parts = re.split(r"[;:,：／／・・　\s]+|:::+", str(phone))
    results = []
    for p in parts:
        p = re.sub(r"[^\d]", "", p)
        if not p:
            continue
        # 先頭0が欠けていたら推定補完（携帯など）
        if len(p) == 10 and not p.startswith("0"):
            if p.startswith(("70","80","90")):
                p = "0" + p
        elif len(p) == 9 and p.startswith(("3","6","4")):
            p = "0" + p
        # ハイフン付け
        if len(p) == 10 and p.startswith("0"):
            p = f"{p[:2]}-{p[2:6]}-{p[6:]}"
        elif len(p) == 11 and p.startswith("0"):
            p = f"{p[:3]}-{p[3:7]}-{p[7:]}"
        results.append(p)
    # 重複排除
    uniq = []
    for r in results:
        if r not in uniq:
            uniq.append(r)
    return ";".join(uniq)

# ========= メール整形 =========
def normalize_emails(value):
    if not value:
        return ""
    parts = re.split(r"[:;,：　\s]+|:::+", str(value))
    cleaned = []
    for p in parts:
        p = p.strip()
        if p and p not in cleaned:
            cleaned.append(p)
    return ";".join(cleaned)

# ========= 住所分割 + 全角変換 =========
def split_address(addr, region, city):
    if not addr:
        return "", "", ""
    addr = addr.replace("　", " ").strip()
    parts = [p.strip() for p in re.split(r"[\n\r]", addr) if p.strip()]
    if len(parts) >= 2:
        a1, a2 = parts[0], parts[1]
        a3 = parts[2] if len(parts) > 2 else ""
    else:
        a1, a2, a3 = addr, "", ""
    full1 = f"{region}{city}{a1}"
    return to_zenkaku(full1), to_zenkaku(a2), to_zenkaku(a3)

# ========= 会社名かな生成 =========
def kana_company_name(name):
    if not name:
        return ""
    name = str(name)
    name = re.sub(r"(株式会社|有限会社|合同会社|一般社団法人|公益財団法人|財団法人)", "", name)
    hira_to_kata = str.maketrans({
        **{chr(i): chr(i + 0x60) for i in range(ord("ぁ"), ord("ゖ") + 1)},
        **{chr(i): chr(i + 0xFEE0) for i in range(ord("a"), ord("z") + 1)},
        **{chr(i): chr(i + 0xFEE0) for i in range(ord("A"), ord("Z") + 1)},
    })
    name = name.translate(hira_to_kata)
    return name.upper()

# ========= メイン変換 =========
def convert_google_to_atena(reader):
    rows=[]
    for r in reader:
        first, last = r.get("First Name",""), r.get("Last Name","")
        first_kana, last_kana = r.get("Phonetic First Name",""), r.get("Phonetic Last Name","")
        org, dept, title = r.get("Organization Name",""), r.get("Organization Department",""), r.get("Organization Title","")
        nickname, note, birthday = r.get("Nickname",""), r.get("Notes",""), r.get("Birthday","")

        # --- メール ---
        emails=[]
        for i in range(1,6):
            v=r.get(f"E-mail {i} - Value","")
            if v: emails.extend(normalize_emails(v).split(";"))
        email_str=";".join(sorted(set(e for e in emails if e)))

        # --- 電話 ---
        phone_dict={}
        for i in range(1,6):
            label=(r.get(f"Phone {i} - Label","") or "").lower()
            val=normalize_phone(r.get(f"Phone {i} - Value",""))
            if not val: continue
            if "work" in label: phone_dict.setdefault("work",[]).append(val)
            elif "mobile" in label: phone_dict.setdefault("mobile",[]).append(val)
            elif "home" in label: phone_dict.setdefault("home",[]).append(val)
        phones=[]
        for k in ["work","mobile","home"]:
            if k in phone_dict: phones.extend(phone_dict[k])
        phone_str=";".join(sorted(set(";".join(phones).split(";"))))

        # --- 住所 ---
        addr_full=r.get("Address 1 - Street","") or r.get("Address 1 - Formatted","")
        region, city = r.get("Address 1 - Region","") or "", r.get("Address 1 - City","") or ""
        zip_code=(r.get("Address 1 - Postal Code","") or "").replace("－","-").replace("ー","-")
        a1,a2,a3=split_address(addr_full, region, city)
        zip_code = to_zenkaku(zip_code)

        # --- メモ ---
        memos={f"メモ{i}":"" for i in range(1,6)}
        for k,v in r.items():
            if "メモ" in k and v:
                for i in range(1,6):
                    if f"メモ{i}" in k:
                        memos[f"メモ{i}"]=v.strip()

        # --- 出力 ---
        rows.append({
            "姓":last,"名":first,"姓かな":last_kana,"名かな":first_kana,
            "姓名":f"{last}　{first}","姓名かな":f"{last_kana}　{first_kana}",
            "ミドルネーム":r.get("Middle Name",""),"ミドルネームかな":r.get("Phonetic Middle Name",""),
            "敬称":"様","ニックネーム":nickname,"旧姓":"","宛先":"会社",
            "自宅〒":"","自宅住所1":"","自宅住所2":"","自宅住所3":"","自宅電話":"","自宅IM ID":"","自宅E-mail":"","自宅URL":"","自宅Social":"",
            "会社〒":zip_code,"会社住所1":a1,"会社住所2":a2,"会社住所3":a3,
            "会社電話":phone_str,"会社IM ID":"","会社E-mail":email_str,"会社URL":"","会社Social":"",
            "その他〒":"","その他住所1":"","その他住所2":"","その他住所3":"","その他電話":"","その他IM ID":"","その他E-mail":"","その他URL":"","その他Social":"",
            "会社名かな":kana_company_name(org),"会社名":org,"部署名1":dept,"部署名2":"","役職名":title,
            "連名":"","連名ふりがな":"","連名敬称":"","連名誕生日":"",
            "メモ1":memos["メモ1"],"メモ2":memos["メモ2"],"メモ3":memos["メモ3"],"メモ4":memos["メモ4"],"メモ5":memos["メモ5"],
            "備考1":note,"備考2":"","備考3":"","誕生日":birthday,"性別":"選択なし","血液型":"選択なし","趣味":"","性格":""
        })
    return rows

# ========= Flaskルート =========
@app.route("/", methods=["GET","POST"])
def upload_file():
    if request.method=="POST":
        file=request.files["file"]
        if not file:
            return render_template_string(HTML)
        try:
            text=file.read().decode("utf-8")
        except UnicodeDecodeError:
            file.seek(0)
            text=file.read().decode("utf-16")
        reader=csv.DictReader(io.StringIO(text))
        result_rows=convert_google_to_atena(reader)

        out=io.StringIO()
        writer=csv.DictWriter(out, fieldnames=list(result_rows[0].keys()))
        writer.writeheader(); writer.writerows(result_rows)
        mem=io.BytesIO(); mem.write(out.getvalue().encode("utf-8-sig")); mem.seek(0)
        return send_file(mem, as_attachment=True, download_name="google_converted.csv", mimetype="text/csv")
    return render_template_string(HTML)

if __name__=="__main__":
    app.run(host="0.0.0.0", port=10000)
