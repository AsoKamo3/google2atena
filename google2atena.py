# google2atena_v3_4.py
# Google連絡先CSV → 宛名職人CSV 変換
# 会社名かな(法人格除外) + 宛先固定 + 住所整形 + メモ復活

import io, csv, os, re
from flask import Flask, request, send_file, abort

app = Flask(__name__)

# ====== 全角・半角変換 ======
FULLWIDTH_OFFSET = ord("！") - ord("!")

def to_zenkaku(s):
    if not isinstance(s, str): return ""
    return "".join(chr(ord(ch)+FULLWIDTH_OFFSET) if 33 <= ord(ch) <= 126 else "　" if ch==" " else ch for ch in s)

def to_hankaku(s):
    if not isinstance(s, str): return ""
    return "".join(chr(ord(ch)-FULLWIDTH_OFFSET) if 65281 <= ord(ch) <= 65374 else " " if ch=="　" else ch for ch in s)

# ====== 法人格除外リスト ======
CORP_PREFIXES = ["株式会社","有限会社","合同会社","一般社団法人","一般財団法人","公益社団法人","公益財団法人","社団法人","財団法人"]

# ====== 会社名かな生成 ======
def to_katakana(text):
    if not text:
        return ""
    hira_to_kata = str.maketrans(
        "ぁあぃいぅうぇえぉおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん",
        "ァアィイゥウェエォオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"
    )
    text = "".join("　" if ch in [" ", "\t"] else ch for ch in text)
    return text.translate(hira_to_kata)

def generate_company_kana(name):
    """法人格を除いた本体のみカタカナ変換"""
    name = name.strip()
    if not name:
        return ""
    for prefix in CORP_PREFIXES:
        if name.startswith(prefix):
            name = name[len(prefix):].strip("　 ")
            break
    return to_zenkaku(to_katakana(name))

# ====== 会社名整形 ======
def normalize_company_name(name):
    if not name: return ""
    for p in CORP_PREFIXES:
        if name.startswith(p):
            rest = name[len(p):].lstrip("　 ")
            return p + "　" + rest
    return name

# ====== 電話・メール ======
def normalize_phone_one(num):
    n = re.sub(r"\D","",num or "")
    if not n: return ""
    if len(n)==11 and n.startswith(("090","080","070")):
        return f"{n[:3]}-{n[3:7]}-{n[7:]}"
    if len(n)==10:
        if n.startswith(("03","06")):
            return f"{n[:2]}-{n[2:6]}-{n[6:]}"
        return f"{n[:3]}-{n[3:6]}-{n[6:]}"
    return n

def normalize_phones(val):
    tmp = val.replace(":::", ";")
    out = []
    for p in [x.strip() for x in tmp.split(";") if x.strip()]:
        fmt = normalize_phone_one(p)
        if fmt: out.append(fmt)
    return out

def normalize_emails(val):
    tmp = val.replace(":::", ";")
    seen, out = set(), []
    for p in [x.strip() for x in tmp.split(";") if x.strip()]:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out

# ====== 住所分割 ======
BUILDING_KEYWORDS = ["ビル","マンション","ハイツ","アパート","コーポ","タワー","ヒルズ","荘","レジデンス","ハウス","テラス","メゾン","棟","号館","プレイス","キャッスル","プラザ","センター","タウン","コート","パレス","ガーデン","ヒル","ビュー","スクエア"]

def normalize_address(addr):
    if not addr: return "",""
    s = to_zenkaku(addr.strip())
    s = re.sub(r"[‐–—−-]", "－", s)
    s = re.sub(r"([０-９]+)丁目", r"\1－", s)
    s = re.sub(r"([０-９]+)番地?", r"\1－", s)
    s = re.sub(r"([０-９]+)号(室)?", r"\1", s)
    s = re.sub(r"－{2,}", "－", s)

    # 空白で建物名を判別
    spaces = [m.start() for m in re.finditer(r"[ 　]+", s)]
    if spaces:
        for sp in spaces:
            left = s[:sp].rstrip()
            right = s[sp:].strip()
            if any(k in right for k in BUILDING_KEYWORDS):
                return left, right

    for kw in BUILDING_KEYWORDS:
        m = re.search(re.escape(kw), s)
        if m:
            pos = m.start()
            msp = list(re.finditer(r"[ 　]+", s[:pos]))
            if msp:
                sp = msp[-1].start()
                left = s[:sp].rstrip()
                right = s[sp:].strip()
                return left, right
            else:
                return s[:pos].rstrip(), s[pos:].strip()
    return s, ""

# ====== CSV読み込み ======
def read_csv(file_bytes):
    for enc in ("utf-8-sig","utf-8","cp932","shift_jis"):
        try:
            f = io.StringIO(file_bytes.decode(enc))
            return list(csv.DictReader(f))
        except: continue
    raise ValueError("CSV読み込み失敗")

def classify_label(label):
    if not label: return "other"
    l = label.lower()
    if "work" in l or "会社" in l: return "work"
    if "home" in l or "自宅" in l: return "home"
    return "other"

# ====== 1件変換 ======
def build_record(row):
    get = lambda k: (row.get(k,"") or "").strip()
    last, first = get("Last Name"), get("First Name")
    last_k, first_k = get("Phonetic Last Name"), get("Phonetic First Name")
    sei_mei = f"{last}　{first}".strip()
    sei_mei_kana = f"{last_k}　{first_k}".strip()

    org = normalize_company_name(get("Organization Name"))
    org_kana = generate_company_kana(org)
    dept, title = get("Organization Department"), get("Organization Title")
    nickname, notes, birthday = get("Nickname"), get("Notes"), get("Birthday")

    # メール・電話
    emails, phones = {"work":[],"home":[],"other":[]}, {"work":[],"home":[],"other":[]}
    for n in range(1,10):
        label = get(f"E-mail {n} - Label"); vals = normalize_emails(get(f"E-mail {n} - Value"))
        if vals: emails[classify_label(label)].extend(vals)
        label = get(f"Phone {n} - Label"); vals = normalize_phones(get(f"Phone {n} - Value"))
        if vals: phones[classify_label(label)].extend(vals)

    # 住所
    addrs = {"work":{"postal":"","line":""},"home":{"postal":"","line":""},"other":{"postal":"","line":""}}
    for n in range(1,10):
        typ = classify_label(get(f"Address {n} - Label"))
        region, city, street, postal = get(f"Address {n} - Region"), get(f"Address {n} - City"), get(f"Address {n} - Street"), get(f"Address {n} - Postal Code")
        if any([region,city,street,postal]):
            addrs[typ]["postal"] = to_hankaku(postal)
            addrs[typ]["line"]   = f"{region}{city}{street}"

    # メモ
    memos = {}
    for n in range(1,10):
        label, val = get(f"Relation {n} - Label"), get(f"Relation {n} - Value")
        if "メモ" in label and val:
            idx = re.sub(r"[^0-9]","",label)
            if idx.isdigit() and 1 <= int(idx) <= 5:
                memos[f"メモ{idx}"] = to_zenkaku(val)

    # 分割
    def pack_addr(d):
        base, bld = normalize_address(d["line"])
        return to_zenkaku(base), to_zenkaku(bld), d["postal"]

    w_base,w_bld,w_post = pack_addr(addrs["work"])
    h_base,h_bld,h_post = pack_addr(addrs["home"])
    o_base,o_bld,o_post = pack_addr(addrs["other"])

    join = lambda v:";".join([x for x in v if x])

    # 出力
    return {
        "姓":to_zenkaku(last),"名":to_zenkaku(first),
        "姓かな":to_zenkaku(last_k),"名かな":to_zenkaku(first_k),
        "姓名":to_zenkaku(sei_mei),"姓名かな":to_zenkaku(sei_mei_kana),
        "ミドルネーム":"","ミドルネームかな":"",
        "敬称":"様","ニックネーム":to_zenkaku(nickname),"旧姓":"","宛先":"会社",
        "自宅〒":h_post,"自宅住所1":h_base,"自宅住所2":h_bld,"自宅住所3":"",
        "自宅電話":join(phones["home"]),"自宅IM ID":"","自宅E-mail":join(emails["home"]),
        "自宅URL":"","自宅Social":"",
        "会社〒":w_post,"会社住所1":w_base,"会社住所2":w_bld,"会社住所3":"",
        "会社電話":join(phones["work"]),"会社IM ID":"","会社E-mail":join(emails["work"]),
        "会社URL":"","会社Social":"",
        "その他〒":o_post,"その他住所1":o_base,"その他住所2":o_bld,"その他住所3":"",
        "その他電話":join(phones["other"]),"その他IM ID":"","その他E-mail":join(emails["other"]),
        "その他URL":"","その他Social":"",
        "会社名かな":org_kana,"会社名":to_zenkaku(org),
        "部署名1":to_zenkaku(dept),"部署名2":"","役職名":to_zenkaku(title),
        "連名":"","連名ふりがな":"","連名敬称":"","連名誕生日":"",
        "メモ1":memos.get("メモ1",""),"メモ2":memos.get("メモ2",""),
        "メモ3":memos.get("メモ3",""),"メモ4":memos.get("メモ4",""),"メモ5":memos.get("メモ5",""),
        "備考1":to_zenkaku(notes),"備考2":"","備考3":"",
        "誕生日":to_zenkaku(birthday),"性別":"選択なし","血液型":"選択なし","趣味":"","性格":""
    }

# ====== Flask ======
@app.route("/")
def index():
    return """
    <h2>Google → 宛名職人 変換ツール v3.4</h2>
    <form method="post" action="/convert" enctype="multipart/form-data">
      <input type="file" name="file" accept=".csv" required>
      <button type="submit">変換開始</button>
    </form>
    """

@app.route("/convert", methods=["POST"])
def convert():
    file = request.files["file"]
    rows = read_csv(file.read())
    out = [build_record(r) for r in rows]

    field_order = [
        "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな",
        "敬称","ニックネーム","旧姓","宛先",
        "自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
        "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social",
        "その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
        "会社名かな","会社名","部署名1","部署名2","役職名",
        "連名","連名ふりがな","連名敬称","連名誕生日",
        "メモ1","メモ2","メモ3","メモ4","メモ5",
        "備考1","備考2","備考3","誕生日","性別","血液型","趣味","性格"
    ]

    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=field_order)
    writer.writeheader(); writer.writerows(out)

    return send_file(io.BytesIO(output.getvalue().encode("utf-8-sig")),
                     mimetype="text/csv",
                     as_attachment=True,
                     download_name="google_converted.csv")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
