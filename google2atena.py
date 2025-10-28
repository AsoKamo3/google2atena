# google2atena_v3_1.py
# Google連絡先CSV → 宛名職人形式CSV（住所分割ロジック改良版）

import io, csv, os, re
from flask import Flask, request, send_file, abort

app = Flask(__name__)

# ================== 文字変換ユーティリティ ==================

FULLWIDTH_OFFSET = ord("！") - ord("!")

def to_zenkaku(s):
    if not isinstance(s, str): return ""
    res = []
    for ch in s:
        code = ord(ch)
        if ch == " ":
            res.append("　")
        elif 33 <= code <= 126:
            res.append(chr(code + FULLWIDTH_OFFSET))
        else:
            res.append(ch)
    return "".join(res)

def to_hankaku(s):
    if not isinstance(s, str): return ""
    res = []
    for ch in s:
        if ch == "　":
            res.append(" ")
            continue
        code = ord(ch)
        if 65281 <= code <= 65374:
            res.append(chr(code - FULLWIDTH_OFFSET))
        else:
            res.append(ch)
    return "".join(res)

# ================== 会社名正規化 ==================

CORP_PREFIXES = ["株式会社","有限会社","合同会社","一般社団法人","一般財団法人","公益社団法人","公益財団法人","社団法人","財団法人"]

def normalize_company_name(name):
    if not name: return ""
    for p in CORP_PREFIXES:
        if name.startswith(p):
            rest = name[len(p):].lstrip("　 ")
            return p + "　" + rest
    return name

# ================== 電話番号正規化 ==================

def normalize_phone(num):
    num = re.sub(r"\D", "", num or "")
    if len(num) == 11 and num.startswith(("090","080","070")):
        return f"{num[:3]}-{num[3:7]}-{num[7:]}"
    elif len(num) == 10:
        return f"{num[:2]}-{num[2:6]}-{num[6:]}"
    return num

# ================== メール正規化 ==================

def normalize_email(val):
    if not val: return ""
    val = val.replace(":::", ";").replace("：：：", ";")
    parts = [v.strip() for v in val.split(";") if v.strip()]
    seen, clean = set(), []
    for p in parts:
        if p not in seen:
            seen.add(p)
            clean.append(p)
    return ";".join(clean)

# ================== 住所正規化 ==================

BUILDING_KEYWORDS = [
    "ビル","マンション","ハイツ","アパート","コーポ","タワー","ヒルズ","荘","レジデンス",
    "ハウス","テラス","メゾン","棟","号館","レーベン","キャッスル","プレイス"
]

def normalize_address(addr):
    """住所文字列を住所1（番地まで）と住所2（建物名）に分割"""
    if not addr:
        return "", ""
    s = to_zenkaku(addr.strip())
    s = re.sub(r"[‐–—−-]", "－", s)
    s = re.sub(r"([０-９]+)丁目", r"\1－", s)
    s = re.sub(r"([０-９]+)番地?", r"\1－", s)
    s = re.sub(r"([０-９]+)号(室)?", r"\1", s)
    s = re.sub(r"－{2,}", "－", s)

    # --- 改良: スペースを利用して建物名を検出 ---
    # （例）宇田川７－１３ 第二共同ビル　５Ｆ → base: 宇田川７－１３, bld: 第二共同ビル　５Ｆ
    pattern_space = rf"(.+?)\s+((?:{'|'.join(BUILDING_KEYWORDS)}).*?[ＦF階＃0-9０-９]+)"
    m = re.search(pattern_space, s)
    if m:
        base = m.group(1).strip()
        bld = m.group(2).strip()
        return base, bld

    # --- 建物キーワードでの分割（フォールバック）---
    pattern_bld = rf"({'|'.join(BUILDING_KEYWORDS)}.*?[ＦF階＃0-9０-９]+)"
    m2 = re.search(pattern_bld, s)
    if m2:
        base = s[:m2.start()].rstrip()
        bld = m2.group(1).strip()
        return base, bld

    return s, ""

# ================== CSV読み込み ==================

def read_csv(file_bytes):
    for enc in ("utf-8-sig","utf-8","cp932","shift_jis"):
        try:
            f = io.StringIO(file_bytes.decode(enc))
            reader = csv.DictReader(f)
            return [{(k or "").strip(): (v or "").strip() for k,v in row.items()} for row in reader]
        except Exception:
            continue
    raise ValueError("CSV読み込みに失敗しました")

# ================== ラベル分類 ==================

def classify_label(label):
    if not label: return "other"
    l = label.lower()
    if any(w in l for w in ["work","mobile","勤務先","会社","職場"]):
        return "work"
    if any(w in l for w in ["home","自宅","家"]):
        return "home"
    return "other"

# ================== レコード生成 ==================

def build_record(row):
    get = lambda k: row.get(k,"").strip()

    last, first = get("Last Name"), get("First Name")
    last_k, first_k = get("Phonetic Last Name"), get("Phonetic First Name")
    sei_mei = f"{last}　{first}".strip()
    sei_mei_kana = f"{last_k}　{first_k}".strip()

    nickname = get("Nickname")
    org = normalize_company_name(get("Organization Name"))
    dept = get("Organization Department")
    title = get("Organization Title")
    notes = get("Notes")
    birthday = get("Birthday")

    # メール
    emails = {"work":[],"home":[],"other":[]}
    for n in range(1,10):
        label = get(f"E-mail {n} - Label")
        val = normalize_email(get(f"E-mail {n} - Value"))
        if val:
            grp = classify_label(label)
            emails[grp].append(val)

    # 電話
    phones = {"work":[],"home":[],"other":[]}
    for n in range(1,10):
        label = get(f"Phone {n} - Label")
        val = normalize_phone(get(f"Phone {n} - Value"))
        if val:
            grp = classify_label(label)
            phones[grp].append(val)

    # 住所
    addrs = {"work":{"postal":"","line":""},
             "home":{"postal":"","line":""},
             "other":{"postal":"","line":""}}
    for n in range(1,10):
        typ = classify_label(get(f"Address {n} - Label"))
        city = get(f"Address {n} - City")
        region = get(f"Address {n} - Region")
        street = get(f"Address {n} - Street")
        postal = get(f"Address {n} - Postal Code")
        if any([city,region,street,postal]):
            addrs[typ]["postal"] = to_hankaku(postal)
            addrs[typ]["line"] = f"{region}{city}{street}"

    def pack_addr(d):
        base,bld = normalize_address(d["line"])
        return to_zenkaku(base), to_zenkaku(bld), d["postal"]

    w_base,w_bld,w_post = pack_addr(addrs["work"])
    h_base,h_bld,h_post = pack_addr(addrs["home"])
    o_base,o_bld,o_post = pack_addr(addrs["other"])

    def join_vals(lst): return ";".join([v for v in lst if v])

    row_out = {
        "姓": to_zenkaku(last),
        "名": to_zenkaku(first),
        "姓かな": to_zenkaku(last_k),
        "名かな": to_zenkaku(first_k),
        "姓名": to_zenkaku(sei_mei),
        "姓名かな": to_zenkaku(sei_mei_kana),
        "ミドルネーム": "",
        "ミドルネームかな": "",
        "敬称": "様",
        "ニックネーム": to_zenkaku(nickname),
        "旧姓": "",
        "宛先": "会社",
        "自宅〒": h_post, "自宅住所1": h_base, "自宅住所2": h_bld, "自宅住所3": "",
        "自宅電話": join_vals(phones["home"]), "自宅IM ID": "", "自宅E-mail": join_vals(emails["home"]),
        "自宅URL": "", "自宅Social": "",
        "会社〒": w_post, "会社住所1": w_base, "会社住所2": w_bld, "会社住所3": "",
        "会社電話": join_vals(phones["work"]), "会社IM ID": "", "会社E-mail": join_vals(emails["work"]),
        "会社URL": "", "会社Social": "",
        "その他〒": o_post, "その他住所1": o_base, "その他住所2": o_bld, "その他住所3": "",
        "その他電話": join_vals(phones["other"]), "その他IM ID": "", "その他E-mail": join_vals(emails["other"]),
        "その他URL": "", "その他Social": "",
        "会社名かな": "", "会社名": to_zenkaku(org),
        "部署名1": to_zenkaku(dept), "部署名2": "", "役職名": to_zenkaku(title),
        "連名": "", "連名ふりがな": "", "連名敬称": "", "連名誕生日": "",
        "メモ1": "", "メモ2": "", "メモ3": "", "メモ4": "", "メモ5": "",
        "備考1": to_zenkaku(notes), "備考2": "", "備考3": "",
        "誕生日": to_zenkaku(birthday), "性別": "選択なし", "血液型": "選択なし",
        "趣味": "", "性格": ""
    }
    return row_out

# ================== Flaskルート ==================

@app.route("/")
def index():
    return """
    <h2>Google → 宛名職人 変換ツール v3.1</h2>
    <form method="post" action="/convert" enctype="multipart/form-data">
      <input type="file" name="file" accept=".csv" required>
      <button type="submit">変換開始</button>
    </form>
    """

@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        abort(400, "CSVファイルを選択してください")
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
    writer.writeheader()
    writer.writerows(out)

    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name="google_converted.csv"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
