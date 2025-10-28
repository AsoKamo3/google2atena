# google2atena_v396.py
# Google連絡先CSV → 宛名職人CSV 変換
# v3.9.6 full-format no-pandas refined stable版

from flask import Flask, request, render_template_string, send_file
import csv, io, re
from collections import OrderedDict

app = Flask(__name__)

TITLE = "Google連絡先CSV → 宛名職人CSV 変換（v3.9.6 refined stable）"

HTML = f"""
<!doctype html>
<meta charset="utf-8">
<title>{TITLE}</title>
<h2>{TITLE}</h2>
<form method="post" enctype="multipart/form-data">
  <p><input type="file" name="file" accept=".csv">
     <input type="submit" value="変換開始">
</form>
<p style="font-size:12px;color:#555;">
・Google連絡先CSV（BOMあり/なし）対応<br>
・「メモ抽出(Label/Value対応)」「Notes→備考1」「会社名かな法人格除去＋英字読み」<br>
・pandas不使用／Render対応済
</p>
"""

# ============ Utility ============

def coalesce(*vals):
    for v in vals:
        if v and str(v).strip():
            return str(v).strip()
    return ""

def to_fullwidth(s):
    if not s: return ""
    fw = str.maketrans({chr(i): chr(i + 0xFEE0) for i in range(0x21, 0x7F)})
    return str(s).translate(fw).replace(" ", "　")

def to_half_postal(s):
    if not s: return ""
    digits = re.sub(r"\D", "", str(s))
    return f"{digits[:3]}-{digits[3:]}" if len(digits)==7 else digits

def format_phone(num):
    if not num: return ""
    digits = re.sub(r"\D", "", str(num))
    if digits and digits[0]!="0" and len(digits) in (9,10): digits="0"+digits
    if len(digits)==11: return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    if len(digits)==10:
        return f"{digits[:2]}-{digits[2:6]}-{digits[6:]}" if digits.startswith(("03","06")) \
               else f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return digits

# ============ Address ============

def parse_address(row):
    postal_raw = coalesce(row.get("Address 1 - Postal Code"))
    region = coalesce(row.get("Address 1 - Region"))
    city = coalesce(row.get("Address 1 - City"))
    street = coalesce(row.get("Address 1 - Street"))
    if not (region and city and street):
        formatted = coalesce(row.get("Address 1 - Formatted"))
        if formatted:
            lines=[l.strip() for l in re.split(r"[\r\n]+", formatted) if l.strip()]
            if len(lines)>=1 and not street: street=lines[0]
            if len(lines)>=2 and not city: city=lines[1]
            if len(lines)>=3 and not region: region=lines[2]
            if len(lines)>=4 and not postal_raw: postal_raw=lines[3]
    postal = to_half_postal(postal_raw)
    addr_num, bldg = street, ""
    if street and re.search(r"[ 　]", street):
        idx = re.search(r"[ 　]", street).start()
        addr_num, bldg = street[:idx].strip(), street[idx:].strip()
    addr1 = to_fullwidth(region+city+addr_num)
    addr2 = to_fullwidth(bldg)
    return postal, addr1, addr2, ""

# ============ Email / Phone ============

def collect_emails(row):
    emails=[]
    for i in range(1,8):
        v=row.get(f"E-mail {i} - Value","")
        if v: emails += re.split(r"[;,\s:]+",v)
    seen=[]
    for e in emails:
        e=e.strip()
        if e and "@" in e and e not in seen: seen.append(e)
    return ";".join(seen)

def collect_phones(row):
    phones=[]
    for i in range(1,8):
        val=row.get(f"Phone {i} - Value","")
        if not val: continue
        for part in re.split(r"[;,\s:]+",val):
            fmt=format_phone(part)
            if fmt and fmt not in phones: phones.append(fmt)
    return ";".join(phones)

# ============ Memo / Notes ============

def normalize_label(s):
    if not s: return ""
    s=str(s).strip()
    s=s.translate(str.maketrans("①②③④⑤１２３４５","12345"))
    return re.sub(r"\s+","",s)

def collect_memo_and_notes(row):
    memos={f"メモ{i}":"" for i in range(1,6)}
    # Relation n - Label / Value
    for i in range(1,10):
        lbl=normalize_label(row.get(f"Relation {i} - Label",""))
        val=row.get(f"Relation {i} - Value","")
        if not val: continue
        m=re.search(r"(?:メモ|memo)([1-5]?)",lbl,re.IGNORECASE)
        if m:
            idx=m.group(1) or None
            if idx and f"メモ{idx}" in memos:
                memos[f"メモ{idx}"]=val.strip()
            else:
                # 未指定の場合は空きスロットに
                for k in memos:
                    if not memos[k]:
                        memos[k]=val.strip(); break
    notes=coalesce(row.get("Notes",""))
    return memos, notes

# ============ Company Kana ============

COMPANY_REPLACE = [
    ("ＮＨＫ","エヌエイチケー"),("日経","ニッケイ"),("博報堂","ハクホウドウ"),
    ("講談社","コウダンシャ"),("テレビ東京","テレビトウキョウ"),
    ("夢眠社","ユメミシャ"),("ライズ＆プレイ","ライズアンドプレイ"),
    ("河野文庫","コウノブンコ"),("東京","トウキョウ"),("湘南","ショウナン")
]
ENGLISH_MAP={
    "NHK":"エヌエイチケー","DY":"ディーワイ","BP":"ビーピー","TV":"ティーブイ",
    "INC":"インク","CORP":"コープ","CO":"シーオー","LTD":"エルティーディー"
}

def to_company_kana(org):
    if not org: return ""
    s=str(org)
    s=re.sub(r"㈱|株式会社|有限会社|合同会社|社団法人|学校法人","",s)
    s=s.strip()
    s=to_fullwidth(s)
    for k,v in COMPANY_REPLACE:
        if k in s: s=s.replace(k,v)
    for k,v in ENGLISH_MAP.items():
        s=re.sub(rf"\b{k}\b",v,s,flags=re.IGNORECASE)
    return s

# ============ Conversion ============

OUT_COLS=[
"姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称","ニックネーム","旧姓","宛先",
"自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
"会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social",
"その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
"会社名かな","会社名","部署名1","部署名2","役職名","連名","連名ふりがな","連名敬称","連名誕生日",
"メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
]

def convert_row(r):
    first,last=r.get("First Name",""),r.get("Last Name","")
    first_kana,last_kana=r.get("Phonetic First Name",""),r.get("Phonetic Last Name","")
    middle,middle_kana=r.get("Middle Name",""),r.get("Phonetic Middle Name","")
    nickname,birthday=r.get("Nickname",""),r.get("Birthday","")
    org,dept,title=r.get("Organization Name",""),r.get("Organization Department",""),r.get("Organization Title","")
    postal,addr1,addr2,addr3=parse_address(r)
    phones,emails=collect_phones(r),collect_emails(r)
    memos,notes=collect_memo_and_notes(r)
    org_kana=to_company_kana(org)

    out=OrderedDict()
    out["姓"]=last; out["名"]=first
    out["姓かな"]=last_kana; out["名かな"]=first_kana
    out["姓名"]=f"{last}　{first}".strip()
    out["姓名かな"]=f"{last_kana}　{first_kana}".strip()
    out["ミドルネーム"]=middle; out["ミドルネームかな"]=middle_kana
    out["敬称"]="様"; out["ニックネーム"]=nickname; out["旧姓"]=""
    out["宛先"]="会社"
    for c in OUT_COLS[12:21]: out[c]=""
    out["会社〒"]=postal; out["会社住所1"]=addr1; out["会社住所2"]=addr2; out["会社住所3"]=addr3
    out["会社電話"]=phones; out["会社IM ID"]=""
    out["会社E-mail"]=emails; out["会社URL"]=out["会社Social"]=""
    for c in OUT_COLS[30:39]: out[c]=""
    out["会社名かな"]=org_kana; out["会社名"]=org; out["部署名1"]=dept
    out["部署名2"]="" ; out["役職名"]=title
    for f in ("連名","連名ふりがな","連名敬称","連名誕生日"): out[f]=""
    for i in range(1,6): out[f"メモ{i}"]=memos[f"メモ{i}"]
    out["備考1"]=notes; out["備考2"]=out["備考3"]=""
    out["誕生日"]=birthday; out["性別"]="選択なし"; out["血液型"]="選択なし"
    out["趣味"]=out["性格"]=""
    return OrderedDict((k,out.get(k,"")) for k in OUT_COLS)

def convert_google_to_atena(text):
    reader=csv.DictReader(io.StringIO(text))
    return [convert_row(r) for r in reader]

# ============ Flask ============

@app.route("/",methods=["GET","POST"])
def upload():
    if request.method=="POST":
        f=request.files.get("file")
        if not f: return render_template_string(HTML)
        raw=f.read(); text=raw.decode("utf-8-sig",errors="replace")
        rows=convert_google_to_atena(text)
        buf=io.StringIO(); w=csv.DictWriter(buf,fieldnames=OUT_COLS,lineterminator="\n")
        w.writeheader(); [w.writerow(r) for r in rows]
        data=buf.getvalue().encode("utf-8-sig")
        return send_file(io.BytesIO(data),as_attachment=True,download_name="google_converted.csv",mimetype="text/csv")
    return render_template_string(HTML)

if __name__=="__main__":
    app.run(host="0.0.0.0",port=10000)
