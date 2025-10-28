# google2atena.py
# Google連絡先のエクスポートCSVを宛名職人形式に変換するWebアプリ（Flask版）

import io, csv, os, re
from flask import Flask, render_template, request, send_file, abort

app = Flask(__name__)

# ======================
# 文字種変換ユーティリティ
# ======================

FULLWIDTH_OFFSET = ord("！") - ord("!")
ASCII_MIN, ASCII_MAX = 33, 126

def to_zenkaku(s):
    if not isinstance(s, str): return ""
    result = []
    for ch in s:
        code = ord(ch)
        if ch == " ":
            result.append("　")
        elif 33 <= code <= 126:
            result.append(chr(code + FULLWIDTH_OFFSET))
        else:
            result.append(ch)
    return "".join(result)

def to_hankaku(s):
    if not isinstance(s, str): return ""
    result = []
    for ch in s:
        if ch == "　":
            result.append(" ")
            continue
        code = ord(ch)
        if 65281 <= code <= 65374:
            result.append(chr(code - FULLWIDTH_OFFSET))
        else:
            result.append(ch)
    return "".join(result)

def normalize_half(s):
    """郵便番号・電話・メールを半角化"""
    s = (s or "")
    return to_hankaku(s.strip())

# ======================
# 法人格スペース処理
# ======================

CORP_PREFIXES = ["株式会社", "有限会社", "合同会社",
                 "一般社団法人", "一般財団法人",
                 "公益社団法人", "公益財団法人",
                 "社団法人", "財団法人"]

def normalize_company_name(name):
    if not name: return ""
    for p in CORP_PREFIXES:
        if name.startswith(p):
            rest = name[len(p):].lstrip("　 ")
            return p + "　" + rest
    return name

# ======================
# 住所正規化
# ======================

BUILDING_KEYWORDS = [
    "ビル","マンション","ハイツ","アパート","コーポ",
    "タワー","ヒルズ","荘","レジデンス","ハウス","テラス","メゾン"
]

def normalize_address(addr):
    """丁目・番・号をハイフン化／ビル名抽出"""
    if not addr: return "",""
    s = addr.strip()
    s = to_zenkaku(s)
    s = re.sub(r"[‐-–—−-]", "－", s)
    s = re.sub(r"([０-９]+)丁目", r"\1－", s)
    s = re.sub(r"([０-９]+)番地?", r"\1－", s)
    s = re.sub(r"([０-９]+)号(室)?", r"\1", s)
    s = re.sub(r"－{2,}", "－", s)

    # ビル・部屋番号抽出
    m = re.search(rf"({'|'.join(BUILDING_KEYWORDS)})([　 ]*[０-９]+)", s)
    if m:
        base = s[:m.start()].rstrip()
        bname = m.group(1)
        num = m.group(2).strip()
        return base, f"　{bname}　＃{num}"
    # 末尾数字3桁以上は地番とみなして残す
    if re.search(r"[０-９]{3,}$", s): return s,""
    # 末尾2〜3桁なら部屋番号として抽出
    m2 = re.search(r"(.+?)－([０-９]{1,3})$", s)
    if m2:
        return m2.group(1), f"　＃{m2.group(2)}"
    return s,""

# ======================
# CSV 読み込み
# ======================

def read_csv(file_bytes):
    for enc in ("utf-8-sig", "utf-8", "cp932", "shift_jis"):
        try:
            f = io.StringIO(file_bytes.decode(enc))
            reader = csv.DictReader(f)
            rows = [{(k or "").strip(): (v or "").strip() for k,v in row.items()} for row in reader]
            return rows
        except Exception:
            continue
    raise ValueError("CSV読み込みに失敗しました")

# ======================
# データ変換ロジック
# ======================

def classify_label(label):
    """ラベル文字列をwork/home/other分類に"""
    if not label: return "other"
    l = label.lower()
    if any(w in l for w in ["work","mobile","勤務先","会社","職場"]):
        return "work"
    if any(w in l for w in ["home","自宅","家"]):
        return "home"
    return "other"

def build_record(row):
    """Google1件を宛名職人形式に変換"""
    get = lambda k: row.get(k,"").strip()

    # 基本情報
    last, first = get("Last Name"), get("First Name")
    last_k, first_k = get("Phonetic Last Name"), get("Phonetic First Name")

    sei_mei = (last + "　" + first).strip()
    sei_mei_kana = (last_k + "　" + first_k).strip()

    nickname = get("Nickname")
    org = normalize_company_name(get("Organization Name"))
    dept = get("Organization Department")
    title = get("Organization Title")
    notes = get("Notes")
    birthday = get("Birthday")

    # メール
    emails = {"work":[],"home":[],"other":[]}
    for n in range(1,10):
        for variant in [f"E-mail {n} - Label", f"E-mail {n} - Type"]:
            if variant in row:
                label = get(variant)
                valkey = f"E-mail {n} - Value"
                val = get(valkey)
                if val:
                    grp = classify_label(label)
                    emails[grp].append(normalize_half(val))

    # 電話
    phones = {"work":[],"home":[],"other":[]}
    for n in range(1,10):
        for variant in [f"Phone {n} - Label", f"Phone {n} - Type"]:
            if variant in row:
                label = get(variant)
                valkey = f"Phone {n} - Value"
                val = get(valkey)
                if val:
                    grp = classify_label(label)
                    phones[grp].append(normalize_half(val))

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
            addrs[typ]["postal"] = normalize_half(postal)
            addrs[typ]["line"] = f"{region}{city}{street}"

    # Relation（メモ）
    memos = {}
    for n in range(1,10):
        label = get(f"Relation {n} - Label")
        val = get(f"Relation {n} - Value")
        if "メモ" in label:
            idx = re.sub(r"[^0-9]", "", label)
            if not idx: continue
            memos[f"メモ{idx}"] = to_zenkaku(val)

    # 住所正規化
    def pack_addr(d):
        base, bld = normalize_address(d["line"])
        return to_zenkaku(base), to_zenkaku(bld), d["postal"]

    w_base, w_bld, w_post = pack_addr(addrs["work"])
    h_base, h_bld, h_post = pack_addr(addrs["home"])
    o_base, o_bld, o_post = pack_addr(addrs["other"])

    def join_vals(lst): return ";".join([v for v in lst if v])

    row_out = {
        "姓": to_zenkaku(last),
        "名": to_zenkaku(first),
        "姓かな": to_zenkaku(last_k),
        "名かな": to_zenkaku(first_k),
        "姓名": to_zenkaku(sei_mei),
        "姓名かな": to_zenkaku(sei_mei_kana),
        "ニックネーム": to_zenkaku(nickname),
        "会社名": to_zenkaku(org),
        "部署名1": to_zenkaku(dept),
        "役職名": to_zenkaku(title),
        "会社〒": w_post,
        "会社住所1": w_base,
        "会社住所2": w_bld,
        "会社電話1〜10": join_vals(phones["work"]),
        "会社E-mail1〜5": join_vals(emails["work"]),
        "自宅〒": h_post,
        "自宅住所1": h_base,
        "自宅住所2": h_bld,
        "自宅電話1〜10": join_vals(phones["home"]),
        "自宅E-mail1〜5": join_vals(emails["home"]),
        "その他〒": o_post,
        "その他住所1": o_base,
        "その他住所2": o_bld,
        "その他電話1〜10": join_vals(phones["other"]),
        "その他E-mail1〜5": join_vals(emails["other"]),
        "備考1": to_zenkaku(notes),
        "誕生日": to_zenkaku(birthday)
    }

    # メモを追加
    for i in range(1,6):
        row_out[f"メモ{i}"] = memos.get(f"メモ{i}","")

    return row_out

# ======================
# Flaskルート
# ======================

@app.route("/")
def index():
    return """
    <h2>Google → 宛名職人 変換ツール</h2>
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

    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=out[0].keys())
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
