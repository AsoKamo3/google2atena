# -*- coding: utf-8 -*-
"""
google2atena v3.9.16 aligned full-format no-pandas（import済）
"""

from flask import Flask, request, Response, render_template_string
import csv, io, re, unicodedata

# ---- 外部辞書 ----
from company_dicts import COMPANY_EXCEPT
from kanji_word_map import KANJI_WORD_MAP
from corp_terms import CORP_TERMS

app = Flask(__name__)

# ======== 文字処理ユーティリティ ========

def to_half(s: str) -> str:
    """英数記号を半角化"""
    return unicodedata.normalize("NFKC", s) if s else ""

def to_full_for_address(s: str) -> str:
    """住所中の数字・記号・英字を全角化（地名漢字はそのまま）"""
    if not s:
        return ""
    out = []
    for ch in s:
        if re.match(r"[A-Za-z0-9\-#]", ch):
            out.append(unicodedata.normalize("NFKC", ch).translate(str.maketrans("-#","－＃")))
            out[-1] = unicodedata.normalize("NFKC", out[-1])
            out[-1] = unicodedata.normalize("NFKC", out[-1])
            out[-1] = unicodedata.normalize("NFKC", out[-1])
            out[-1] = out[-1].translate(str.maketrans("-#","－＃"))
        else:
            out.append(ch)
    return "".join(out)

# ======== 郵便番号・電話番号 ========

def normalize_postal(s: str) -> str:
    if not s: return ""
    s = re.sub(r"[^\d-]", "", s)
    return s if s else ""

def normalize_phone(raw: str) -> str:
    if not raw: return ""
    raw = re.sub(r"[,、／／;]+", ";", raw)
    parts = [p.strip() for p in raw.split(";") if p.strip()]
    out = []
    for p in parts:
        num = re.sub(r"\D", "", p)
        if len(num) == 11 and num.startswith(("070","080","090")):
            out.append(f"{num[:3]}-{num[3:7]}-{num[7:]}")
        elif len(num) == 10:
            if num.startswith(("03","06")):
                out.append(f"{num[:2]}-{num[2:6]}-{num[6:]}")
            else:
                out.append(f"{num[:3]}-{num[3:6]}-{num[6:]}")
        else:
            out.append(num)
    return ";".join(out)

# ======== 住所分割 ========

_BUILDING_HINT = re.compile(r"(ビル|号室|号|階|Ｆ|F|＃|#|棟|室)")

def split_street_and_building(street: str):
    if not street: return ("","")
    s = to_full_for_address(street.strip())
    lines = re.split(r"[　\s]+", s)
    for i, part in enumerate(lines):
        if _BUILDING_HINT.search(part):
            return ("　".join(lines[:i]) or s, "　".join(lines[i:]))
    return (s, "")

def compose_address(region, city, street, ext):
    region = region or ""
    city = city or ""
    street = street or ""
    ext = ext or ""
    base, bld = split_street_and_building(street)
    line1 = to_full_for_address(region + city + base)
    line2 = to_full_for_address(bld or ext)
    return line1.strip(), line2.strip()

# ======== メモ・Notes ========

_MEMO_KEYS = [f"メモ{i}" for i in range(1,6)]
def collect_memos_and_notes(row):
    memos = {k:"" for k in _MEMO_KEYS}
    notes = []
    for i in range(1,10):
        lbl = row.get(f"Relation {i} - Label","")
        val = row.get(f"Relation {i} - Value","")
        if not val: continue
        if "メモ" in lbl: 
            m = re.search(r"メモ(\d)", lbl)
            if m: memos[f"メモ{m.group(1)}"] = val
        elif "備考" in lbl or "Note" in lbl or "notes" in lbl.lower():
            notes.append(val)
    nfield = row.get("Notes","")
    if nfield: notes.append(nfield)
    n1, n2, n3 = (notes+[ "", "", "" ])[:3]
    return memos, n1, n2, n3

# ======== 会社名かな変換 ========

_ABC = {
 "A":"エー","B":"ビー","C":"シー","D":"ディー","E":"イー","F":"エフ","G":"ジー","H":"エイチ","I":"アイ",
 "J":"ジェー","K":"ケー","L":"エル","M":"エム","N":"エヌ","O":"オー","P":"ピー","Q":"キュー","R":"アール",
 "S":"エス","T":"ティー","U":"ユー","V":"ブイ","W":"ダブリュー","X":"エックス","Y":"ワイ","Z":"ズィー"
}

def ascii_to_kana(text):
    return re.sub(r"[A-Za-z]+", lambda m: "".join(_ABC.get(c.upper(), c) for c in m.group()), text)

def strip_corp(name: str):
    s = name.strip()
    for term in sorted(CORP_TERMS, key=len, reverse=True):
        s = re.sub(rf"^{term}", "", s)
        s = re.sub(rf"{term}$", "", s)
    return s.strip()

def company_to_kana(name: str) -> str:
    if not name: return ""
    if name in COMPANY_EXCEPT: return COMPANY_EXCEPT[name]
    s = strip_corp(name)
    for k,v in sorted(KANJI_WORD_MAP.items(), key=lambda x: len(x[0]), reverse=True):
        s = s.replace(k, v)
    s = ascii_to_kana(s)
    s = re.sub(r"[・.,，]", "", s)
    return unicodedata.normalize("NFKC", s)

# ======== CSV 変換 ========

OUTPUT_HEADERS = [
 "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称","ニックネーム","旧姓","宛先",
 "自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
 "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social",
 "その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
 "会社名かな","会社名","部署名1","部署名2","役職名","連名","連名ふりがな","連名敬称","連名誕生日",
 "メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
]

def convert_address_block(row, idx):
    prefix = f"Address {idx} - "
    lbl = (row.get(prefix+"Label","") or "").lower()
    reg, city, st, ext = row.get(prefix+"Region",""), row.get(prefix+"City",""), row.get(prefix+"Street",""), row.get(prefix+"Extended Address","")
    postal = normalize_postal(row.get(prefix+"Postal Code",""))
    a1,a2 = compose_address(reg,city,st,ext)
    if "home" in lbl: t="自宅"
    elif "other" in lbl: t="その他"
    else: t="会社"
    return t, postal, a1, a2

def put_address(out,t,post,a1,a2):
    if t=="自宅":
        out["自宅〒"]=post; out["自宅住所1"]=a1; out["自宅住所2"]=a2
    elif t=="その他":
        out["その他〒"]=post; out["その他住所1"]=a1; out["その他住所2"]=a2
    else:
        out["会社〒"]=post; out["会社住所1"]=a1; out["会社住所2"]=a2

def clean_emails(value: str) -> str:
    if not value: return ""
    emails = re.split(r"[:;、\s]+", value)
    uniq = []
    for e in emails:
        e=e.strip()
        if e and e not in uniq: uniq.append(e)
    return ";".join(uniq)

def convert_row(row):
    out={k:"" for k in OUTPUT_HEADERS}
    out["姓"],out["名"]=row.get("Last Name",""),row.get("First Name","")
    out["姓かな"],out["名かな"]=row.get("Phonetic Last Name",""),row.get("Phonetic First Name","")
    out["姓名"]=out["姓"]+"　"+out["名"]; out["姓名かな"]=out["姓かな"]+"　"+out["名かな"]
    out["敬称"]="様"; out["宛先"]="会社"
    # 住所
    for i in range(1,4):
        t,p,a1,a2=convert_address_block(row,i)
        put_address(out,t,p,a1,a2)
    # 電話
    for i in range(1,5):
        lbl=(row.get(f"Phone {i} - Label","") or "").lower()
        val=row.get(f"Phone {i} - Value","")
        if not val: continue
        num=normalize_phone(val)
        if "home" in lbl: out["自宅電話"]=num
        elif "other" in lbl: out["その他電話"]=num
        else: out["会社電話"]=num
    # メール
    for i in range(1,6):
        val=row.get(f"E-mail {i} - Value","")
        lbl=(row.get(f"E-mail {i} - Label","") or "").lower()
        if not val: continue
        v=clean_emails(val)
        if "home" in lbl: out["自宅E-mail"]=v
        elif "other" in lbl: out["その他E-mail"]=v
        else: out["会社E-mail"]=v
    # 会社
    org=row.get("Organization Name","")
    out["会社名"]=org
    out["会社名かな"]=company_to_kana(org)
    out["部署名1"]=to_full_for_address(row.get("Organization Department",""))
    out["役職名"]=to_full_for_address(row.get("Organization Title",""))
    # メモ
    memos,n1,n2,n3=collect_memos_and_notes(row)
    for k,v in memos.items(): out[k]=v
    out["備考1"],out["備考2"],out["備考3"]=n1,n2,n3
    out["誕生日"]=row.get("Birthday","")
    return out

def convert_google_to_atena(text):
    reader=csv.DictReader(io.StringIO(text))
    return [convert_row(r) for r in reader]

# ======== Flask UI ========

HTML = """
<!doctype html><html lang="ja"><meta charset="utf-8"><title>google2atena v3.9.16</title>
<body style="font-family:sans-serif;padding:24px"><h1>google2atena v3.9.16</h1>
<form method="post" enctype="multipart/form-data">
<input type="file" name="file"><button type="submit">変換</button></form></body></html>
"""

@app.route("/",methods=["GET","POST"])
def home():
    if request.method=="GET": return HTML
    f=request.files.get("file")
    if not f: return "ファイル未指定",400
    text=f.read().decode("utf-8","ignore")
    rows=convert_google_to_atena(text)
    sio=io.StringIO()
    w=csv.DictWriter(sio,fieldnames=OUTPUT_HEADERS,lineterminator="\n"); w.writeheader()
    for r in rows: w.writerow(r)
    return Response(sio.getvalue(),headers={
        "Content-Type":"text/csv; charset=utf-8",
        "Content-Disposition":'attachment; filename="atena_v3_9_16.csv"'
    })

if __name__=="__main__":
    app.run(host="0.0.0.0",port=10000,debug=True)
