# -*- coding: utf-8 -*-
"""
v3.9.15 aligned full-format no-pandas（import済）
- Flask + csv 標準ライブラリのみ（no pandas）
- 住所分岐：Work→会社*, Home→自宅*, Other→その他*
- 郵便番号：各ブロックに格納、半角→全角、「-」→「－」
- 電話番号：半角→整形（携帯：3-4-4/3-4-4、固定：市外局番優先で 2/3/4-4/3/2 分割、なければ 3-4-4）
- 会社名かな：法人格除去、記号削除（「・ . ，」）、全角化、辞書 + 漢字マップ + 英字読み上げ（A→エー... DY→ディーワイ）
- メモ/Notes：漏れなく収集（Relation系のメモラベル、Notes列、ばらつく表記を正規化）
- 出力カラム順：理想型を完全維持
"""

from flask import Flask, request, Response, render_template_string
import csv
import io
import re
import unicodedata

# ---- 外部辞書（同ディレクトリに配置）----
from company_dicts import COMPANY_EXCEPT   # 例外：完全一致の会社名→かな
from kanji_word_map import KANJI_WORD_MAP  # 置換：部分一致の漢字語→カタカナ
from corp_terms import CORP_TERMS          # 法人格の網羅リスト（約40種以上）

app = Flask(__name__)

# ========= ユーティリティ =========

# 全角化（英数・ハイフン・スペース）
def to_zenkaku(s: str) -> str:
    if not s:
        return ""
    # 基本：英数・記号を全角に
    z = unicodedata.normalize("NFKC", s)
    # ここで半角→全角の「-」を全角ダッシュへ
    z = z.replace("-", "－")
    # 半角スペース→全角スペース
    z = z.replace(" ", "　")
    return z

# 郵便番号 正規化（半角→全角、ハイフン統一）
def normalize_postal(s: str) -> str:
    if not s:
        return ""
    s = s.strip()
    s = re.sub(r"[^\d\-－]", "", s)
    s = s.replace("-", "－")
    # 全角化（数字とハイフン）
    s = to_zenkaku(s)
    return s

# 電話番号：半角数字抽出→分割→「-」は全角「－」
_MOBILE_HEADS = ("070", "080", "090")
# 市外局番（主要な先頭パターン。より詳細にしたければ拡張可）
_LONG_AREA_PREFIX = ("011", "015", "017", "018", "019", "022", "023", "024", "025", "026", "027",
                     "028", "029", "042", "043", "044", "045", "046", "047", "048", "049", "052",
                     "053", "054", "055", "058", "059", "072", "073", "074", "075", "076", "077",
                     "078", "079", "082", "083", "084", "085", "086", "087", "088", "089", "092",
                     "093", "094", "095", "096", "097", "098")
def normalize_phone(raw: str) -> str:
    if not raw:
        return ""
    # 複数値区切り（:::や;や,や全角読点など）→統一
    raw = re.sub(r"\s*(:::+|[;,、／／/])\s*", ";", raw.strip())
    parts = [p for p in raw.split(";") if p.strip()]
    out = []
    for p in parts:
        digits = re.sub(r"\D", "", p)
        if not digits:
            continue
        # 携帯
        if len(digits) in (10, 11) and digits.startswith(_MOBILE_HEADS):
            if len(digits) == 11:
                pretty = f"{digits[0:3]}-{digits[3:7]}-{digits[7:]}"
            else:  # 10桁（稀）
                pretty = f"{digits[0:3]}-{digits[3:6]}-{digits[6:]}"
        else:
            # 固定電話：市外局番推定（3桁優先→2桁→4桁）
            pretty = None
            if len(digits) >= 10:
                # 3桁候補
                if digits[:3] in _LONG_AREA_PREFIX:
                    pretty = f"{digits[:3]}-{digits[3:-4]}-{digits[-4:]}"
                # 2桁候補（03,06など）
                elif digits[:2] in ("03", "06"):
                    pretty = f"{digits[:2]}-{digits[2:-4]}-{digits[-4:]}"
                # 4桁市外（0120等フリーダイヤル含む）
                elif digits[:4] in ("0120", "0800"):
                    if len(digits) == 10:
                        pretty = f"{digits[:4]}-{digits[4:6]}-{digits[6:]}"
                    else:
                        pretty = f"{digits[:4]}-{digits[4:-4]}-{digits[-4:]}"
            # 後方フォールバック（3-4-4）
            if not pretty:
                if len(digits) >= 8:
                    pretty = f"{digits[:-8]}-{digits[-8:-4]}-{digits[-4:]}" if len(digits) > 10 else f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
                else:
                    pretty = digits
        # 全角ハイフン
        pretty = pretty.replace("-", "－")
        out.append(pretty)
    return ";".join(out)

# 住所：Streetから建物っぽい語尾を切り出す
_BUILDING_KEYWORDS = (
    "ビル", "タワー", "レジデンス", "マンション", "ハイツ", "荘", "寮",
    "号室", "号", "階", "Ｆ", "F", "#", "＃", "棟", "室"
)
def split_street_and_building(street: str):
    if not street:
        return "", ""
    # 半角→全角、スペース正規化（単一スペース→全角空白）
    s = to_zenkaku(street.strip())
    # スペースで分割して建物語尾を後半へ寄せる
    tokens = re.split(r"[　\s]+", s)
    if len(tokens) == 1:
        # キーワードが末尾にあればそこから建物扱い
        for kw in _BUILDING_KEYWORDS:
            m = re.search(rf"{re.escape(kw)}[　\s]*.*$", s)
            if m:
                base = s[:m.start()].strip()
                bld = s[m.start():].strip()
                return base, bld
        return s, ""
    # 先頭から建物キーワードが出る手前までをベース、それ以降を建物
    base_parts = []
    bld_parts = []
    found = False
    for t in tokens:
        if not found and any(kw in t for kw in _BUILDING_KEYWORDS):
            found = True
        (bld_parts if found else base_parts).append(t)
    return ("　".join(base_parts).strip(), "　".join(bld_parts).strip())

# 住所行を構成（都道府県+市区町村+丁目番地を1行、建物は2行目）
def compose_address(region, city, street, ext):
    # region/pref, city/ward, street（丁目番地＋建物）、ext(Extended)
    region = to_zenkaku(region or "")
    city = to_zenkaku(city or "")
    street = to_zenkaku(street or "")
    ext = to_zenkaku(ext or "")
    # Streetから建物分離
    base, bld = split_street_and_building(street)
    line1 = "".join([region, city, base]).strip()
    # 建物は優先（bld or ext）
    line2 = (bld or ext).strip()
    return line1, line2

# メモ／Notes 収集
_MEMO_SLOT_NAMES = ["メモ1", "メモ2", "メモ3", "メモ4", "メモ5"]
def normalize_memo_label(lbl: str) -> str:
    if not lbl:
        return ""
    s = lbl.strip()
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("　", " ").lower()
    # "memo", "メモ", "ﾒﾓ" いずれでも
    if "memo" in s or "メモ" in lbl or "ﾒﾓ" in lbl:
        m = re.search(r"(?:memo|メモ|ﾒﾓ)\s*([1-5１-５])", s, re.I)
        if m:
            d = m.group(1)
            d = unicodedata.normalize("NFKC", d)
            if d in "12345":
                return f"メモ{d}"
    # 完全一致（ゆるめ）
    for i in range(1, 6):
        if s.replace(" ", "") in (f"メモ{i}", f"memo{i}".lower()):
            return f"メモ{i}"
    # Notes/備考系
    if "notes" in s or "note" in s or "備考" in s:
        return "備考"
    return ""

def collect_memos_and_notes(row: dict):
    memos = {k: "" for k in _MEMO_SLOT_NAMES}
    notes_acc = []

    # Google CSV の Relation i - Label / Value に「メモ」類が入ることが多い
    for i in range(1, 10):
        lbl = row.get(f"Relation {i} - Label", "") or row.get(f"Relation{ i } - Label", "")
        val = row.get(f"Relation {i} - Value", "") or row.get(f"Relation{ i } - Value", "")
        t = normalize_memo_label(lbl)
        if t in memos and val:
            memos[t] = (memos[t] + "；" if memos[t] else "") + val.strip()
        elif t == "備考" and val:
            notes_acc.append(val.strip())

    # 行内に「メモ1」「メモ 1」みたいなカラムが紛れている場合（安全側）
    for k, v in row.items():
        nk = normalize_memo_label(k)
        if nk in memos and v:
            memos[nk] = (memos[nk] + "；" if memos[nk] else "") + v.strip()
        elif nk == "備考" and v:
            notes_acc.append(v.strip())

    # Notes 列
    notes_field = row.get("Notes", "") or row.get("Note", "")
    if notes_field:
        notes_acc.append(notes_field.strip())

    # 3つの備考欄にローリング格納
    notes_joined = "｜".join([n for n in notes_acc if n])
    notes1, notes2, notes3 = "", "", ""
    if notes_joined:
        # 雑に3分割（長さで切る）
        chunks = re.split(r"[｜\n]{1,}", notes_joined)
        bucket = ["", "", ""]
        b = 0
        for c in chunks:
            if not c.strip():
                continue
            if not bucket[b]:
                bucket[b] = c.strip()
            else:
                bucket[b] += "｜" + c.strip()
            # 次へ
            if len(bucket[b]) > 120 and b < 2:
                b += 1
        notes1, notes2, notes3 = bucket
    return memos, notes1, notes2, notes3

# 会社名かな：法人格除去→例外辞書→漢字置換→英字読み上げ→記号削除→全角
_ABC2KANA = {
    "A":"エー","B":"ビー","C":"シー","D":"ディー","E":"イー","F":"エフ","G":"ジー","H":"エイチ","I":"アイ",
    "J":"ジェー","K":"ケー","L":"エル","M":"エム","N":"エヌ","O":"オー","P":"ピー","Q":"キュー","R":"アール",
    "S":"エス","T":"ティー","U":"ユー","V":"ブイ","W":"ダブリュー","X":"エックス","Y":"ワイ","Z":"ズィー"
}
def strip_corp_terms(name: str) -> str:
    s = name or ""
    s = s.strip()
    # 前後の法人格を削る（複合・全角半角・スペースあり対応）
    for term in sorted(CORP_TERMS, key=len, reverse=True):
        s = re.sub(rf"^\s*{re.escape(term)}\s*", "", s)
        s = re.sub(rf"\s*{re.escape(term)}\s*$", "", s)
    return s.strip()

def ascii_run_to_kana(run: str) -> str:
    res = []
    for ch in run.upper():
        res.append(_ABC2KANA.get(ch, ch))
    return "".join(res)

def roman_to_kana_blocks(text: str) -> str:
    # 連続英字を見つけて読み上げ化（DY→ディーワイ、KADOKAWA→ケーエーディーオーケーエーダブリューエー）
    def repl(m):
        return ascii_run_to_kana(m.group(0))
    return re.sub(r"[A-Za-z]{2,}", repl, text)

def company_to_kana(orig_company: str) -> str:
    if not orig_company:
        return ""
    s = orig_company.strip()
    # 例外完全一致
    if s in COMPANY_EXCEPT:
        kana = COMPANY_EXCEPT[s]
    else:
        # 法人格除去
        core = strip_corp_terms(s)
        # 漢字→カタカナ（辞書順に長い語から）
        work = core
        for k in sorted(KANJI_WORD_MAP.keys(), key=len, reverse=True):
            work = work.replace(k, KANJI_WORD_MAP[k])
        # 英字連続を読み上げ変換
        work = roman_to_kana_blocks(work)
        # 記号削除：「・ . ，」
        work = re.sub(r"[・\.\,，]", "", work)
        kana = work
    # 全角化
    kana = to_zenkaku(kana)
    return kana

# CSV 読み取り（区切り自動判定）
def read_csv_rows(text: str):
    bio = io.StringIO(text)
    sample = bio.read(2048)
    bio.seek(0)
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
    except Exception:
        # デフォルトはカンマ
        dialect = csv.excel
    reader = csv.DictReader(bio, dialect=dialect)
    return list(reader)

# ========= 変換本体 =========

# 出力カラム順（理想型）
OUTPUT_HEADERS = [
    "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな","敬称","ニックネーム","旧姓","宛先",
    "自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
    "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social",
    "その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
    "会社名かな","会社名","部署名1","部署名2","役職名","連名","連名ふりがな","連名敬称","連名誕生日",
    "メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3",
    "誕生日","性別","血液型","趣味","性格"
]

def convert_address_block(row: dict, idx: int):
    """
    Address i - Label/Formated/Street/City/Region/Postal Code/Country/Extended Address
    → ラベルにより 会社/自宅/その他 に振り分け、住所1/2/3 + 郵便
    """
    prefix = f"Address {idx} - "
    label = (row.get(prefix+"Label","") or "").strip().lower()
    street = row.get(prefix+"Street","") or ""
    city = row.get(prefix+"City","") or ""
    region = row.get(prefix+"Region","") or ""
    postal = row.get(prefix+"Postal Code","") or ""
    country = row.get(prefix+"Country","") or ""
    ext = row.get(prefix+"Extended Address","") or ""
    formatted = row.get(prefix+"Formatted","") or ""

    # 基本は個別項目優先、無ければ Formatted を行分割して推定
    if not (street or city or region or postal or ext) and formatted:
        # フォーマット想定：
        # <Street>\n<City>\n<Region>\n<Postal>\n<Country>
        lines = [l.strip() for l in formatted.splitlines() if l.strip()]
        if len(lines) >= 1: street = lines[0]
        if len(lines) >= 2: city = lines[1]
        if len(lines) >= 3: region = lines[2]
        if len(lines) >= 4: postal = lines[3]
        if len(lines) >= 5: country = lines[4]

    adr1, adr2 = compose_address(region, city, street, ext)
    # adr3 は今回は空（必要なら拡張可）
    adr3 = ""

    # 郵便番号
    post = normalize_postal(postal)

    # どこへ入れるか
    # Google側は "work","home","other" など
    target = "会社"
    if "home" in label:
        target = "自宅"
    elif "other" in label:
        target = "その他"
    elif "work" in label:
        target = "会社"

    return target, post, adr1, adr2, adr3

def put_address(target, post, adr1, adr2, adr3, outrow: dict):
    if target == "自宅":
        outrow["自宅〒"] = post or outrow["自宅〒"]
        outrow["自宅住所1"] = adr1 or outrow["自宅住所1"]
        outrow["自宅住所2"] = adr2 or outrow["自宅住所2"]
        outrow["自宅住所3"] = adr3 or outrow["自宅住所3"]
    elif target == "その他":
        outrow["その他〒"] = post or outrow["その他〒"]
        outrow["その他住所1"] = adr1 or outrow["その他住所1"]
        outrow["その他住所2"] = adr2 or outrow["その他住所2"]
        outrow["その他住所3"] = adr3 or outrow["その他住所3"]
    else:  # 会社
        outrow["会社〒"] = post or outrow["会社〒"]
        outrow["会社住所1"] = adr1 or outrow["会社住所1"]
        outrow["会社住所2"] = adr2 or outrow["会社住所2"]
        outrow["会社住所3"] = adr3 or outrow["会社住所3"]

def convert_row(row: dict) -> dict:
    # 初期化（全カラムを空で用意）
    out = {k: "" for k in OUTPUT_HEADERS}

    # 氏名・ふりがな
    first = row.get("First Name","") or ""
    last = row.get("Last Name","") or ""
    middle = row.get("Middle Name","") or ""
    pf = row.get("Phonetic First Name","") or ""
    pl = row.get("Phonetic Last Name","") or ""
    pm = row.get("Phonetic Middle Name","") or ""

    out["姓"] = (last or "").strip()
    out["名"] = (first or "").strip()
    out["ミドルネーム"] = (middle or "").strip()
    out["姓かな"] = (pl or "").strip()
    out["名かな"] = (pf or "").strip()
    out["ミドルネームかな"] = (pm or "").strip()

    # 姓名/姓名かな（スペースは全角）
    joiner = "　"
    out["姓名"] = (out["姓"] + joiner + out["名"]).strip(joiner)
    out["姓名かな"] = (out["姓かな"] + joiner + out["名かな"]).strip(joiner)

    # 敬称・ニックネーム・旧姓・宛先
    out["敬称"] = "様"
    out["ニックネーム"] = row.get("Nickname","") or ""
    out["旧姓"] = row.get("Maiden Name","") or ""
    out["宛先"] = "会社"

    # 住所（Address i）
    for i in range(1, 4):
        t, post, a1, a2, a3 = convert_address_block(row, i)
        put_address(t, post, a1, a2, a3, out)

    # 電話（Phone i）
    for i in range(1, 5):
        lbl = (row.get(f"Phone {i} - Label","") or "").lower()
        val = row.get(f"Phone {i} - Value","") or ""
        if not val:
            continue
        pretty = normalize_phone(val)
        if "home" in lbl:
            out["自宅電話"] = ";".join([x for x in [out["自宅電話"], pretty] if x])
        elif "other" in lbl:
            out["その他電話"] = ";".join([x for x in [out["その他電話"], pretty] if x])
        else:  # work / mobile 等は会社へ（要求に明記なし→Work優先、Mobileは会社か自宅か仕様に依らず今回は会社へ集約）
            out["会社電話"] = ";".join([x for x in [out["会社電話"], pretty] if x])

    # IM / E-mail / URL / Social（Home/Work/Other 雑に集約：ラベルで投げ分け）
    for i in range(1, 6):
        # E-mail
        elbl = (row.get(f"E-mail {i} - Label","") or "").lower()
        eval_ = row.get(f"E-mail {i} - Value","") or ""
        if eval_:
            if "home" in elbl:
                out["自宅E-mail"] = ";".join([x for x in [out["自宅E-mail"], eval_.strip()] if x])
            elif "other" in elbl:
                out["その他E-mail"] = ";".join([x for x in [out["その他E-mail"], eval_.strip()] if x])
            else:
                out["会社E-mail"] = ";".join([x for x in [out["会社E-mail"], eval_.strip()] if x])
        # URL
        uval = row.get(f"Website {i} - Value","") or ""
        ulbl = (row.get(f"Website {i} - Label","") or "").lower()
        if uval:
            if "home" in ulbl:
                out["自宅URL"] = ";".join([x for x in [out["自宅URL"], uval.strip()] if x])
            elif "other" in ulbl:
                out["その他URL"] = ";".join([x for x in [out["その他URL"], uval.strip()] if x])
            else:
                out["会社URL"] = ";".join([x for x in [out["会社URL"], uval.strip()] if x])

    # IM/Social は CSV によって列名が異なるため、代表的フィールド名を拾って集約
    for key, dest_home, dest_work, dest_other in [
        ("IM 1 - Value", "自宅IM ID", "会社IM ID", "その他IM ID"),
        ("Social 1 - Value", "自宅Social", "会社Social", "その他Social"),
    ]:
        v = row.get(key, "") or ""
        if v:
            # ラベルがあれば振り分け、なければ会社へ
            lblk = key.replace("Value", "Label")
            lb = (row.get(lblk, "") or "").lower()
            if "home" in lb:
                out[dest_home] = ";".join([x for x in [out[dest_home], v.strip()] if x])
            elif "other" in lb:
                out[dest_other] = ";".join([x for x in [out[dest_other], v.strip()] if x])
            else:
                out[dest_work] = ";".join([x for x in [out[dest_work], v.strip()] if x])

    # 会社名／部署／役職
    org = row.get("Organization Name","") or ""
    dep = row.get("Organization Department","") or ""
    title = row.get("Organization Title","") or ""
    out["会社名"] = org
    out["会社名かな"] = company_to_kana(org)

    # 部署名・役職名は全角化（スペースも全角）
    out["部署名1"] = to_zenkaku(dep)
    # 仕様：部署名2 は空のまま（必要に応じて分割実装も可）
    out["部署名2"] = ""
    out["役職名"] = to_zenkaku(title)

    # 連名系（今回は未使用）
    out["連名"] = ""
    out["連名ふりがな"] = ""
    out["連名敬称"] = ""
    out["連名誕生日"] = ""

    # メモ/Notes
    memos, notes1, notes2, notes3 = collect_memos_and_notes(row)
    for k in _MEMO_SLOT_NAMES:
        out[k] = memos.get(k, "")
    out["備考1"] = notes1
    out["備考2"] = notes2
    out["備考3"] = notes3

    # 誕生日/性別/血液型/趣味/性格（存在すればマッピング）
    out["誕生日"] = row.get("Birthday","") or ""
    out["性別"] = row.get("Gender","") or ""
    out["血液型"] = row.get("Blood Type","") or ""
    out["趣味"] = row.get("Hobby","") or ""
    out["性格"] = row.get("Personality","") or ""

    return out

def convert_google_to_atena(text: str):
    rows = read_csv_rows(text)
    out_rows = []
    for r in rows:
        out_rows.append(convert_row(r))
    return out_rows

# ========= Web UI =========

HTML = """
<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <title>google2atena v3.9.15 aligned no-pandas</title>
  <style>
    body{font-family: system-ui, -apple-system, "Segoe UI", Roboto, "Hiragino Sans", "Noto Sans JP", "Yu Gothic", "Meiryo", sans-serif; padding: 24px;}
    .wrap{max-width: 840px; margin: 0 auto;}
    h1{font-size: 20px;}
    .card{border:1px solid #ddd; border-radius:8px; padding:16px;}
    .hint{color:#666; font-size: 12px;}
    button{padding:10px 16px; font-size:14px;}
  </style>
</head>
<body>
  <div class="wrap">
    <h1>google2atena v3.9.15 aligned full-format (no pandas)</h1>
    <div class="card">
      <form method="post" enctype="multipart/form-data">
        <p><input type="file" name="file" accept=".csv,.txt"></p>
        <p><button type="submit">変換開始</button></p>
      </form>
      <p class="hint">Google 連絡先の CSV またはタブ区切りファイルを選択してください。区切り文字は自動判定します。</p>
    </div>
  </div>
</body>
</html>
"""

@app.route("/", methods=["GET","POST"])
def upload():
    if request.method == "GET":
        return render_template_string(HTML)
    f = request.files.get("file")
    if not f:
        return "ファイルがありません。", 400
    text = f.read().decode("utf-8", errors="ignore")
    rows = convert_google_to_atena(text)

    # CSV 出力
    sio = io.StringIO()
    writer = csv.DictWriter(sio, fieldnames=OUTPUT_HEADERS, lineterminator="\n")
    writer.writeheader()
    for r in rows:
        writer.writerow(r)
    csv_bytes = sio.getvalue().encode("utf-8", errors="ignore")
    return Response(
        csv_bytes,
        headers={
            "Content-Type": "text/csv; charset=utf-8",
            "Content-Disposition": 'attachment; filename="atena_v3_9_15.csv"'
        }
    )

# Render 用
def app_factory():
    return app

if __name__ == "__main__":
    # ローカルデバッグ用
    app.run(host="0.0.0.0", port=10000, debug=True)
