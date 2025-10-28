# google2atena.py
# Google連絡先CSV → 宛名職人CSV 変換（v3.9.4 full-format no-pandas 安定版）
# - pandas不使用（標準csvモジュールのみ）
# - 住所分割：v3.9-pre相当を復活（都道府県+市区町村+番地 / 建物名+階は別欄）
# - 郵便番号：半角のまま（3-4固定、記号は半角ハイフン）
# - 会社住所1/2：全角化（数字・英字・記号"#" "-" " " 等を全角に）
# - 電話：NTT風フォーマット（先頭0欠落の救済＋ハイフン付与）
#   11桁(携帯)：3-4-4、10桁(固定)：03/06 → 2-4-4、その他 → 3-3-4
#   9桁/10桁で先頭が0でない場合は、先頭に"0"を付与してから整形
# - 電話の優先順位：Work > Mobile > Home（Workが複数ある時も、入力順を維持）
# - メール：分割（::: / ; / , / 空白）、重複除去、`;`連結
# - メモ："メモ|memo + 数字(1-5)" で柔軟にマッピング。Notes → 備考1
# - 会社名かな：簡易の表音化辞書（NHK, LAB, WORKS, Office 等）を適用し全角カタカナ出力
# - 出力カラム順：ご指定の順に固定
# - 文字コード：UTF-8(BOM付)でダウンロード

from flask import Flask, request, render_template_string, send_file
import csv
import io
import re
import sys
from collections import OrderedDict

app = Flask(__name__)

TITLE = "Google連絡先CSV → 宛名職人CSV 変換（v3.9.4 full-format no-pandas 安定版）"

HTML = f"""
<!doctype html>
<meta charset="utf-8">
<title>{TITLE}</title>
<h2>{TITLE}</h2>
<form method="post" enctype="multipart/form-data">
  <p><input type="file" name="file" accept=".csv">
     <input type="submit" value="変換開始">
</form>
<p style="font-size:12px;color:#666;">
  ・Google連絡先のUTF-8 CSVに対応（BOMあり/なし）<br>
  ・「住所分割」「郵便番号は半角」「会社住所は全角」「電話NTT整形」「メモ柔軟対応」「Notes→備考1」対応済み
</p>
"""

# ========= ユーティリティ =========

ASCII_TO_FULLWIDTH = str.maketrans(
    # 英数字
    {chr(i): chr(i + 0xFEE0) for i in range(0x21, 0x7F)}
)
# ただし、記号のうち、住所で半角にしたいものは後で個別指定

def to_fullwidth_address(s: str) -> str:
    """住所用：英数字・主要記号を全角化。スペースは全角スペースへ。ハイフンは全角'－'、'#'は全角'＃'。"""
    if not s:
        return ""
    s = str(s)
    # まず全角化（!"#$%&'()*+,-./0-9:;<=>?@A-Z[\]^_`a-z{|}~）
    s = s.translate(ASCII_TO_FULLWIDTH)
    # スペース（半角）→ 全角
    s = s.replace(" ", "　")
    # 半角ハイフン（-）は全角の横線へ置換（U+FF0D）
    s = s.replace("-", "－")
    # # を全角へ
    s = s.replace("＃", "＃")  # translateで既に全角化されるが明示
    # 全角コロンやスラッシュ等はそのままでOK
    return s

def to_halfwidth_postal(s: str) -> str:
    """郵便番号：半角数字と半角ハイフンのみを残し、XXX-XXXX 形式に正規化。"""
    if not s:
        return ""
    s = str(s)
    # 全ての数字を半角に、各種ダッシュを半角'-'に
    # いったん数字以外を除去し、再配置でも良いが、桁が7想定
    digits = re.sub(r"\D", "", s)
    if len(digits) == 7:
        return f"{digits[0:3]}-{digits[3:7]}"
    elif len(digits) == 8 and digits[3] == '0':  # まれな表記の救済（不要なら削る）
        return f"{digits[0:3]}-{digits[4:8]}"
    # 桁不明は、数字のみ返す（安全策）
    return digits

def normalize_phone_raw(num: str) -> str:
    """電話生値を数字だけに（＋先頭の0欠落救済用に判定で使う）"""
    if not num:
        return ""
    return re.sub(r"\D", "", str(num))

def format_jp_phone(num: str) -> str:
    """
    日本の電話整形（NTT風）
    - 携帯（11桁で 070/080/090/060/050?）：3-4-4
    - 固定（10桁）
        03/06 → 2-4-4
        その他 → 3-3-4
    - 9桁/10桁で先頭が0でない場合は先頭に0を付与
    例：
      9065090629 → 090-6509-0629
      112615331  → 011-261-5331
      467839111  → 046-783-9111
      364419772  → 03-6441-9772
    """
    if not num:
        return ""
    digits = normalize_phone_raw(num)

    # 先頭0が無い救済：9桁 or 10桁なら 0 を先頭に付ける
    if digits and digits[0] != "0" and len(digits) in (9, 10):
        digits = "0" + digits

    # 11桁（多くは携帯/050IP等）
    if len(digits) == 11:
        return f"{digits[0:3]}-{digits[3:7]}-{digits[7:11]}"

    # 10桁（固定）
    if len(digits) == 10:
        if digits.startswith(("03", "06")):
            return f"{digits[0:2]}-{digits[2:6]}-{digits[6:10]}"
        else:
            return f"{digits[0:3]}-{digits[3:6]}-{digits[6:10]}"

    # 9桁（まれ）→ 先頭0補完済みでここに来ない想定だが、保険
    if len(digits) == 9:
        return f"{digits[0:2]}-{digits[2:5]}-{digits[5:9]}"

    # それ以外は数字をそのまま返す（安全策）
    return digits

def split_emails(value: str):
    """メール列を ':::', ';', ',', 空白 などで分割し、下準備用に返す。"""
    if not value:
        return []
    # 区切りを統一
    s = str(value).replace(":::", ";").replace("；", ";").replace("，", ",")
    # セミコロン / カンマ / 空白で分割
    parts = re.split(r"[;,\s]+", s)
    # メールらしいものだけ（簡易）
    out = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if "@" in p and "." in p:
            out.append(p.lower())
    return out

def label_rank(label: str) -> int:
    """電話ラベル優先度：Work(0) < Mobile(1) < Home(2) < その他(3)"""
    if not label:
        return 3
    t = label.strip().lower()
    if "work" in t:
        return 0
    if "mobile" in t:
        return 1
    if "home" in t:
        return 2
    return 3

def pick_company_phones(row: dict) -> str:
    """Phone i - Label / Value からWork>Mobile>Homeの順で抽出（カテゴリ内は入力順）"""
    buckets = {0: [], 1: [], 2: [], 3: []}
    for i in range(1, 8):  # 余裕を見て7本分見る
        lbl = row.get(f"Phone {i} - Label", "")
        val = row.get(f"Phone {i} - Value", "")
        if not (lbl or val):
            continue
        # 値が "0334190265 ::: 0368046001" のような複数入りに対応
        raw_list = re.split(r"::+|;|,|\s+", str(val))
        for raw in raw_list:
            raw = raw.strip()
            if not raw:
                continue
            fmt = format_jp_phone(raw)
            if not fmt:
                continue
            buckets[label_rank(lbl)].append(fmt)

    # Work(0) → Mobile(1) → Home(2) → その他(3) の順で結合
    seen = set()
    result = []
    for prio in (0, 1, 2, 3):
        for p in buckets[prio]:
            if p not in seen:
                seen.add(p)
                result.append(p)
    return ";".join(result)

# ---- 住所分割（v3.9-pre 相当の復活）----

def coalesce(*vals):
    for v in vals:
        if v and str(v).strip():
            return str(v).strip()
    return ""

def parse_formatted_address(row: dict):
    """
    Google連絡先の住所系カラムから
      - 郵便番号（半角）
      - 会社住所1（都道府県＋市区町村＋番地）※全角化
      - 会社住所2（建物名・階等）※全角化
      - 会社住所3（未使用/将来拡張）
    を返す。
    優先：Street + (Formattedからの補助) / City / Region / Postal Code
    Street に「宇田川7-13 第二共同ビル 5F」のように入っている場合、
    最初のスペースで「番地」と「建物」を分割。
    """
    postal_raw = coalesce(row.get("Address 1 - Postal Code"), row.get("Address 1 - PO Box"))
    region = coalesce(row.get("Address 1 - Region"))
    city = coalesce(row.get("Address 1 - City"))
    street = coalesce(row.get("Address 1 - Street"))

    # Formattedの補助（空欄が多いときの保険）
    if not (region and city and street):
        formatted = coalesce(row.get("Address 1 - Formatted"))
        if formatted:
            lines = [l.strip() for l in re.split(r"[\r\n]+", formatted) if l.strip()]
            # 典型例：
            # 1行目：street（番地+建物込み）
            # 2行目：市区町村（例：渋谷区 / 札幌市中央区）
            # 3行目：都道府県（例：東京都 / 北海道）
            # 4行目：郵便（例：150-0042）
            if len(lines) >= 1 and not street:
                street = lines[0]
            if len(lines) >= 2 and not city:
                city = lines[1]
            # 3行目と2行目が入れ替わる例があるため簡易補正
            if len(lines) >= 3 and not region:
                # 都道府県らしさ判定（末尾が都/道/府/県）
                if re.search(r"[都道府県]$", lines[2]):
                    region = lines[2]
                elif re.search(r"[都道府県]$", lines[1]) and not re.search(r"[都道府県]$", city):
                    region, city = lines[1], lines[2]
            if len(lines) >= 4 and not postal_raw:
                postal_raw = lines[3]

    # 郵便番号（半角）
    postal = to_halfwidth_postal(postal_raw)

    # Street から「番地」/「建物」を分離（最初の半角 or 全角スペースで分割）
    addr_num = street
    bldg = ""
    if street:
        # 半角スペース/全角スペースのいずれか最初で分割
        m = re.search(r"[ 　]", street)
        if m:
            idx = m.start()
            addr_num = street[:idx].strip()
            bldg = street[idx:].strip()
        else:
            addr_num = street.strip()

    # 会社住所1 = 都道府県 + 市区町村 + 番地（全角化）
    addr1 = "".join([region, city, addr_num]).strip()
    addr1 = to_fullwidth_address(addr1)
    # 会社住所2 = 建物（全角化、前のスペースは取り除き、間は全角スペース1つ）
    addr2 = ""
    if bldg:
        # 先頭の空白類は削除
        bldg = re.sub(r"^[ 　]+", "", bldg)
        # 半角スペースを全角スペースへ揃え、複数スペースは1つに
        bldg = re.sub(r"[ ]+", " ", bldg).replace(" ", "　")
        addr2 = to_fullwidth_address(bldg)

    # 会社住所3 は今回未使用（必要なら今後拡張）
    addr3 = ""

    return postal, addr1, addr2, addr3

# ---- 会社名かな（簡易辞書）----

COMPANY_KANA_MAP = [
    # 長い語や固有名を先に（前方一致で拾う）
    (r"\bKADOKAWA\b", "カドカワ"),
    (r"\bMARUZEN\b", "マルゼン"),
    (r"\bNHK\b", "エヌエイチケー"),
    (r"\bWORKS\b", "ワークス"),
    (r"\bLAB\b", "ラボ"),
    (r"\bOffice\b", "オフィス"),
    (r"\bStand\b", "スタンド"),
    (r"\bNeo\b", "ネオ"),
    (r"\bReal\b", "リアル"),
    (r"\bYADOKARI\b", "ヤドカリ"),
    (r"\btoi\b", "トイ"),
    (r"\bPlan\b", "プラン"),
    (r"\bAll\b", "オール"),
    (r"\bReviews\b", "レビューズ"),
    (r"\bcounter\b", "カウンター"),
    (r"\bodd\b", "オッド"),
    # よく使う出版社等（例）
    (r"PHP研究所", "ピーエイチピーケンキュウショ"),
    (r"夢眠社", "ユメミシャ"),
    # 追加したいものがあればここに追記
]

def to_company_kana(org_name: str) -> str:
    """
    会社名かな（全角カタカナ）を簡易生成。
    - 既知パターンは辞書で置換。
    - 英字は読みを辞書でカバー（英字自体の機械読みは行わない）。
    - 変換なしの場合は、そのまま返す（将来、形態素/辞書連携を検討）。
    """
    if not org_name:
        return ""
    s = str(org_name)
    out = None
    # 英字/語句の辞書置換を優先（大文字小文字を無視）
    for pat, kana in COMPANY_KANA_MAP:
        if re.search(pat, s, flags=re.IGNORECASE):
            out = kana
            break
    if out:
        return out
    # 会社名がすでにカタカナならそれを返す（簡易判定）
    if re.fullmatch(r"[ァ-ヶー　・＝Ａ-Ｚａ-ｚ０-９]+", to_fullwidth_address(s)):
        return to_fullwidth_address(s)
    # それ以外は未変換（要件次第で将来強化）
    return s

# ---- メモ/Notes ----

def collect_memos_and_notes(row: dict):
    """
    入力の「メモ/ memo + 数字」で柔軟にマッピングし、出力「メモ1〜5」へ。
    - ラベルの表記ゆれ：「メモ1」「メモ 1」「メモ１」「memo 1」等
    - Google CSVのカラム名としてどこに出現しても拾う
    また、Notes → 備考1 に格納。
    """
    memos = {"メモ1": "", "メモ2": "", "メモ3": "", "メモ4": "", "メモ5": ""}
    notes_out = row.get("Notes", "") or ""

    memo_pat = re.compile(r"(?:メモ|memo)\s*([１２３４５12345])", re.IGNORECASE)
    z2h = str.maketrans("１２３４５", "12345")

    for key, val in row.items():
        if not val:
            continue
        m = memo_pat.search(str(key))
        if m:
            idx = m.group(1).translate(z2h)
            if idx in "12345":
                memos[f"メモ{idx}"] = str(val)

    return memos, notes_out

# ---- メール ----

def collect_emails(row: dict) -> str:
    """E-mail i - Value を総なめし、分割/重複除去して連結。"""
    allm = []
    for i in range(1, 8):
        v = row.get(f"E-mail {i} - Value", "")
        if v:
            allm.extend(split_emails(v))
    # 重複除去（順序維持）
    seen = set()
    out = []
    for m in allm:
        if m not in seen:
            seen.add(m)
            out.append(m)
    return ";".join(out)

# ---- 変換コア ----

OUT_COLUMNS = [
    "姓","名","姓かな","名かな","姓名","姓名かな",
    "ミドルネーム","ミドルネームかな",
    "敬称","ニックネーム","旧姓","宛先",
    "自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
    "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social",
    "その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
    "会社名かな","会社名","部署名1","部署名2","役職名",
    "連名","連名ふりがな","連名敬称","連名誕生日",
    "メモ1","メモ2","メモ3","メモ4","メモ5",
    "備考1","備考2","備考3",
    "誕生日","性別","血液型","趣味","性格"
]

def convert_row(row: dict) -> OrderedDict:
    first = row.get("First Name","") or ""
    last = row.get("Last Name","") or ""
    first_kana = row.get("Phonetic First Name","") or ""
    last_kana  = row.get("Phonetic Last Name","") or ""
    middle = row.get("Middle Name","") or ""
    middle_kana = row.get("Phonetic Middle Name","") or ""
    nickname = row.get("Nickname","") or ""
    birthday = row.get("Birthday","") or ""
    org = row.get("Organization Name","") or ""
    dept = row.get("Organization Department","") or ""
    title = row.get("Organization Title","") or ""

    # 住所（会社側）— v3.9-pre仕様復活
    postal, addr1, addr2, addr3 = parse_formatted_address(row)

    # 電話（会社）
    company_phones = pick_company_phones(row)

    # メール（会社）
    emails = collect_emails(row)

    # メモ & Notes
    memo_map, notes1 = collect_memos_and_notes(row)

    # 会社名かな（簡易辞書）
    org_kana = to_company_kana(org)

    # 出力整形
    full_name = (last + "　" + first).strip("　 ")
    full_name_kana = (last_kana + "　" + first_kana).strip("　 ")

    out = OrderedDict()
    out["姓"] = last
    out["名"] = first
    out["姓かな"] = last_kana
    out["名かな"] = first_kana
    out["姓名"] = full_name
    out["姓名かな"] = full_name_kana
    out["ミドルネーム"] = middle
    out["ミドルネームかな"] = middle_kana
    out["敬称"] = "様"
    out["ニックネーム"] = nickname
    out["旧姓"] = ""
    out["宛先"] = "会社"

    # 自宅系は今回未展開（必要なら将来対応）
    out["自宅〒"] = ""
    out["自宅住所1"] = ""
    out["自宅住所2"] = ""
    out["自宅住所3"] = ""
    out["自宅電話"] = ""
    out["自宅IM ID"] = ""
    out["自宅E-mail"] = ""
    out["自宅URL"] = ""
    out["自宅Social"] = ""

    # 会社側
    out["会社〒"] = postal  # 半角のまま
    out["会社住所1"] = addr1  # 全角化済み
    out["会社住所2"] = addr2  # 全角化済み
    out["会社住所3"] = addr3
    out["会社電話"] = company_phones
    out["会社IM ID"] = ""
    out["会社E-mail"] = emails
    out["会社URL"] = ""
    out["会社Social"] = ""

    # その他
    out["その他〒"] = ""
    out["その他住所1"] = ""
    out["その他住所2"] = ""
    out["その他住所3"] = ""
    out["その他電話"] = ""
    out["その他IM ID"] = ""
    out["その他E-mail"] = ""
    out["その他URL"] = ""
    out["その他Social"] = ""

    # 組織
    out["会社名かな"] = org_kana  # 全角カタカナ想定 or 原文
    out["会社名"] = org
    out["部署名1"] = dept
    out["部署名2"] = ""
    out["役職名"] = title

    # 連名系（未使用）
    out["連名"] = ""
    out["連名ふりがな"] = ""
    out["連名敬称"] = ""
    out["連名誕生日"] = ""

    # メモ
    out["メモ1"] = memo_map["メモ1"]
    out["メモ2"] = memo_map["メモ2"]
    out["メモ3"] = memo_map["メモ3"]
    out["メモ4"] = memo_map["メモ4"]
    out["メモ5"] = memo_map["メモ5"]

    # 備考
    out["備考1"] = notes1
    out["備考2"] = ""
    out["備考3"] = ""

    # 個人属性（未使用）
    out["誕生日"] = birthday
    out["性別"] = "選択なし"
    out["血液型"] = "選択なし"
    out["趣味"] = ""
    out["性格"] = ""

    # カラム順に並べ替え
    ordered = OrderedDict()
    for col in OUT_COLUMNS:
        ordered[col] = out.get(col, "")
    return ordered

def convert_google_to_atena(csv_text: str) -> list[OrderedDict]:
    reader = csv.DictReader(io.StringIO(csv_text))
    out_rows = []
    for row in reader:
        out_rows.append(convert_row(row))
    return out_rows

# ========= Flask =========

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        f = request.files.get("file")
        if not f:
            return render_template_string(HTML)

        # バイト→テキスト（UTF-8優先、読めない文字は置換して続行）
        raw = f.read()
        try:
            text = raw.decode("utf-8-sig", errors="replace")
        except Exception:
            text = raw.decode("utf-8", errors="replace")

        try:
            rows = convert_google_to_atena(text)
        except Exception as e:
            # 失敗時は簡易エラーレポート
            return f"<pre>変換中にエラーが発生しました。\n{e}</pre>", 500

        # 出力CSV（UTF-8 with BOM）
        buf = io.StringIO()
        writer = csv.DictWriter(buf, fieldnames=OUT_COLUMNS, lineterminator="\n")
        writer.writeheader()
        for r in rows:
            writer.writerow(r)

        data = buf.getvalue().encode("utf-8-sig")
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name="google_converted.csv",
            mimetype="text/csv; charset=utf-8"
        )

    return render_template_string(HTML)

if __name__ == "__main__":
    # Render等の標準ポートに合わせて必要なら環境変数で変更
    app.run(host="0.0.0.0", port=10000)
