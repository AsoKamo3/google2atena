# google2atena.py  v3.9.18r7b+universal_csvfix_10x  (Render安定版 / no-pandas)
# - 電話番号整形：Phone 1-10 に対応、重複削除、;連結
# - メール整形：E-mail 1-10 に対応、:を排除し;統一
# - バージョン表記をHTMLに表示
# - その他の機能は v3.9.18r7b と完全同一

import csv
import io
import re
import unicodedata
from flask import Flask, render_template_string, request, send_file

# ======== 外部辞書フェイルセーフ ========
try:
    from company_dicts import COMPANY_EXCEPT
except Exception:
    COMPANY_EXCEPT = {}

try:
    from kanji_word_map import KANJI_WORD_MAP
except Exception:
    KANJI_WORD_MAP = {}

try:
    from corp_terms import CORP_TERMS
except Exception:
    CORP_TERMS = [
        "株式会社","有限会社","合同会社","合資会社","相互会社",
        "一般社団法人","一般財団法人","公益社団法人","公益財団法人",
        "特定非営利活動法人","ＮＰＯ法人","学校法人","医療法人",
        "宗教法人","社会福祉法人","公立大学法人","独立行政法人","地方独立行政法人"
    ]

app = Flask(__name__)

# ======== 住所ユーティリティ ========

def to_zenkaku_for_address(s: str) -> str:
    if not s: return ""
    z=[]
    for ch in s:
        code=ord(ch)
        if 0x21<=code<=0x7E:
            z.append('－' if ch=='-' else unicodedata.normalize('NFKC',ch))
        else:
            z.append(ch)
    return "".join(z)

def format_postal(postal: str) -> str:
    if not postal: return ""
    d=re.sub(r'\D','',postal)
    return f"{d[:3]}-{d[3:]}" if len(d)==7 else postal

_SPLIT_RE=re.compile(r'[ \u3000]')
def split_first_space(a:str):
    if not a: return ("","")
    m=_SPLIT_RE.search(a)
    if not m: return (a,"")
    i=m.start()
    return (a[:i],a[i+1:].strip())

def build_addr12(r,c,s):
    parts=[p for p in [r,c,s] if p]
    f="".join(parts)
    fz=to_zenkaku_for_address(f)
    a1,a2=split_first_space(fz)
    return (a1,a2)

def route_address_by_label(row,out):
    for i in range(1,11):
        label=(row.get(f"Address {i} - Label") or "").strip().lower()
        region=row.get(f"Address {i} - Region") or ""
        city=row.get(f"Address {i} - City") or ""
        street=row.get(f"Address {i} - Street") or ""
        postal=row.get(f"Address {i} - Postal Code") or ""
        if not (region or city or street or postal): continue
        jp_postal=format_postal(postal)
        a1,a2=build_addr12(region,city,street)
        target="会社"
        if "home" in label: target="自宅"
        elif "other" in label: target="その他"
        out[f"{target}〒"]=jp_postal
        out[f"{target}住所1"]=a1
        out[f"{target}住所2"]=a2
        out[f"{target}住所3"]=""
        # 最初の有効データを採用
        if target=="会社": break

# ======== 電話番号整形 ========

CITY_CODES=['011','015','017','018','019','022','023','024','025','026','027','028','029',
'03','04','042','043','044','045','046','047','048','049','052','053','054',
'055','056','057','058','059','06','072','073','074','075','076','077','078',
'079','082','083','084','085','086','087','088','089','092','093','094','095',
'096','097','098','099']

def normalize_phones(vals):
    phones=[]
    for v in vals:
        if not v: continue
        for raw in re.split(r'[:;、／]', v):
            n=re.sub(r'\D','',raw)
            if not n: continue
            if not n.startswith('0'): n='0'+n
            f=n
            if re.match(r'^0(70|80|90)\d{8}$',n): f=f"{n[:3]}-{n[3:7]}-{n[7:]}"
            elif re.match(r'^050\d{8}$',n): f=f"{n[:3]}-{n[3:7]}-{n[7:]}"
            elif re.match(r'^(0120|0800|0570)\d{6}$',n): f=f"{n[:4]}-{n[4:7]}-{n[7:]}"
            elif len(n)==10:
                for code in sorted(CITY_CODES,key=len,reverse=True):
                    if n.startswith(code):
                        rem=n[len(code):]
                        f=f"{code}-{rem[:len(rem)//2]}-{rem[len(rem)//2:]}"
                        break
            elif len(n)==9: f=f"{n[:2]}-{n[2:5]}-{n[5:]}"
            if f not in phones: phones.append(f)
    return ";".join(phones)

# ======== メール整形 ========

def normalize_emails(vals):
    emails=[]
    for v in vals:
        if not v: continue
        v=v.strip().replace(':',';')
        for e in re.split(r'[;、／]',v):
            e=e.strip()
            if e and e not in emails: emails.append(e)
    return ";".join(emails)

# ======== メモ抽出 ========

def extract_memos(row):
    m=[]
    for i in range(1,11):
        l=row.get(f"Relation {i} - Label","")
        v=row.get(f"Relation {i} - Value","")
        if l and "メモ" in l and v: m.append(v)
    n=row.get("Notes","")
    if n: m.append(n)
    return m

# ======== 会社名かな変換 ========

def kana_company_name(n):
    if not n: return ""
    if n in COMPANY_EXCEPT: return COMPANY_EXCEPT[n]
    r=n
    for k,v in KANJI_WORD_MAP.items():
        if k in r: r=r.replace(k,v)
    return r

# ======== HTML ========

html_form=f"""
<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8">
<title>Google→宛名職人 CSV 変換ツール</title></head><body>
<h2>Google連絡先 → 宛名職人 CSV 変換ツール</h2>
<p><b>Version:</b> v3.9.18r7b+universal_csvfix_10x (Render安定版 / no-pandas)</p>
<form action="/convert" method="post" enctype="multipart/form-data">
<input type="file" name="file" accept=".csv" required>
<input type="submit" value="変換開始">
</form></body></html>
"""

# ======== Flask Routes ========

@app.route("/")
def index():
    return render_template_string(html_form)

@app.route("/convert", methods=["POST"])
def convert():
    f=request.files["file"]
    if not f: return "⚠️ ファイルが選択されていません。"
    text=f.read().decode("utf-8-sig",errors="replace")
    r=csv.DictReader(io.StringIO(text))
    o=io.StringIO()
    w=csv.writer(o)
    header=[
        "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称",
        "ニックネーム","旧姓","宛先","自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話",
        "自宅IM ID","自宅E-mail","自宅URL","自宅Social","会社〒","会社住所1","会社住所2",
        "会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social","その他〒",
        "その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail",
        "その他URL","その他Social","会社名かな","会社名","部署名1","部署名2","役職名",
        "連名","連名ふりがな","連名敬称","連名誕生日","メモ1","メモ2","メモ3","メモ4","メモ5",
        "備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
    ]
    w.writerow(header)

    for row in r:
        out={}
        route_address_by_label(row,out)

        phones=[row.get(f"Phone {i} - Value","") for i in range(1,11)]
        out["会社電話"]=normalize_phones(phones)

        emails=[row.get(f"E-mail {i} - Value","") for i in range(1,11)]
        out["会社E-mail"]=normalize_emails(emails)

        memos=extract_memos(row)
        for i in range(5): out[f"メモ{i+1}"]=memos[i] if i<len(memos) else ""

        cname=row.get("Organization Name","")
        out["会社名"]=cname
        out["会社名かな"]=kana_company_name(cname)

        w.writerow([
            row.get("Last Name",""),row.get("First Name",""),
            row.get("Phonetic Last Name",""),row.get("Phonetic First Name",""),
            f"{row.get('Last Name','')}　{row.get('First Name','')}",
            f"{row.get('Phonetic Last Name','')}{row.get('Phonetic First Name','')}",
            "","", "様", row.get("Nickname",""),"",
            "会社",
            out.get("自宅〒",""),out.get("自宅住所1",""),out.get("自宅住所2",""),out.get("自宅住所3",""),
            out.get("自宅電話",""),"","","","",
            out.get("会社〒",""),out.get("会社住所1",""),out.get("会社住所2",""),out.get("会社住所3",""),
            out.get("会社電話",""),"",out.get("会社E-mail",""),"","",
            out.get("その他〒",""),out.get("その他住所1",""),out.get("その他住所2",""),out.get("その他住所3",""),
            out.get("その他電話",""),"","","","",
            out.get("会社名かな",""),out.get("会社名",""),
            row.get("Organization Department",""),"",row.get("Organization Title",""),
            "","","","",
            out.get("メモ1",""),out.get("メモ2",""),out.get("メモ3",""),out.get("メモ4",""),out.get("メモ5",""),
            "","","","","","","",""
        ])

    o.seek(0)
    return send_file(
        io.BytesIO(o.getvalue().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name="converted.csv"
    )

if __name__=="__main__":
    app.run(host="0.0.0.0",port=10000)
