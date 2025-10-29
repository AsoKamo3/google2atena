# google2atena.py  v3.9.18r7c (Render最終安定版 / no-pandas)
# - 文字コード自動判定（chardet）
# - 区切り自動検出（csv.Sniffer）
# - キー正規化
# - 住所整形：Region+City+Street を連結→全角（数字/英字/記号）→最初の空白で(住所1,住所2)分割／Labelで会社・自宅・その他に振替
# - 電話整形：トークンごとに個別正規化→整形→最後に「;」連結（携帯・特番・固定は最長一致市外局番）
# - メール整形：;連結、重複排除、Home/Work/Otherへ
# - メモ抽出：Relation「メモ1〜5」、Notes→備考
# - 会社かな：外部辞書（company_dicts / kanji_word_map / corp_terms）があれば使用。無ければ最小デフォルトでフェールセーフ。
# - 出力：CSV (UTF-8-SIG) ダウンロード

from flask import Flask, request, send_file, render_template_string
import io, csv, re, codecs
import chardet

# ---- 外部辞書（フェイルセーフ） -------------------------------------------------
try:
    from company_dicts import COMPANY_REPLACE_MAP  # 例: {"（株）": "株式会社", "ＣＯＵＮＴＥＲ": "カウンター", ...}
except Exception:
    COMPANY_REPLACE_MAP = {}

try:
    from kanji_word_map import EN_TO_KATAKANA_MAP  # 例: {"BOOKS":"ブックス","COUNTER":"カウンター",...}
except Exception:
    EN_TO_KATAKANA_MAP = {}

try:
    from corp_terms import CORP_TERMS  # 例: {"株式会社","合同会社","有限会社","Inc.","LLC",...}
except Exception:
    CORP_TERMS = set()

# 市外局番（最長一致ソート済み：長い→短い）
from jp_area_codes import AREA_CODES

app = Flask(__name__)

# ---- UI（最低限） ----------------------------------------------------------------
INDEX_HTML = """<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>google2atena v3.9.18r7c</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
 body{font-family:system-ui,-apple-system,"Segoe UI",Roboto,Helvetica,Arial,"Noto Sans JP","Apple Color Emoji","Segoe UI Emoji";margin:24px;}
 .card{border:1px solid #ddd;border-radius:12px;padding:16px;max-width:780px}
 input[type=file]{margin:.5rem 0}
 button{padding:.6rem 1rem;border-radius:8px;border:1px solid #333;background:#111;color:#fff;cursor:pointer}
 .small{color:#666;font-size:.9em}
</style>
</head>
<body>
<div class="card">
  <h1>google2atena v3.9.18r7c</h1>
  <form action="/convert" method="post" enctype="multipart/form-data">
    <p>Google 連絡先の CSV/TSV をアップロードしてください。</p>
    <input type="file" name="file" accept=".csv,.tsv,.txt" required>
    <div><button type="submit">変換する</button></div>
  </form>
  <p class="small">出力: UTF-8 (BOM付) CSV / 住所は全角（郵便・電話のみ半角ハイフン付）</p>
</div>
</body>
</html>"""

# ---- 汎用：全角変換 --------------------------------------------------------------
# 電話・郵便は半角+ハイフン。それ以外の住所本文は数字/英字/記号を全角に。
def to_zenkaku_except_hyphen(s: str) -> str:
    if not s:
        return s
    # 半角 → 全角（ASCII英数・記号） ※ハイフンだけは全角にせず一旦置換で逃がす
    s = s.replace('-', '\u0000')  # 一時エスケープ
    # 全角化（英数・記号）
    tbl = {ord(c): ord(c) + 0xFEE0 for c in ''.join(chr(i) for i in range(33, 127))}
    s = s.translate(tbl)
    # エスケープ戻し（住所の区切りとしては全角長音のほうが望ましければ 'ー' へ）
    s = s.replace('\u0000', '－')  # 住所中のハイフンは全角長音記号に
    return s

# ---- キー正規化 -----------------------------------------------------------------
def norm_key(k: str) -> str:
    return (k or '').strip().lower().replace(' ', '_').replace('-', '_')

# ---- 値取り出し -----------------------------------------------------------------
def get(row, *keys, default=''):
    for k in keys:
        v = row.get(k)
        if v:
            return v
    return default

# ---- メール正規化 ---------------------------------------------------------------
def normalize_emails(*vals):
    emails = []
    for v in vals:
        if not v:
            continue
        # " ::: " や セミコロン 区切りを尊重
        parts = re.split(r'\s*:::\s*|[;,\s]\s*', str(v))
        for p in parts:
            p = p.strip()
            if not p:
                continue
            # 簡易妥当性
            if '@' in p and '.' in p.split('@')[-1]:
                emails.append(p)
    # 重複排除（順序維持）
    seen = set()
    out = []
    for e in emails:
        el = e.lower()
        if el not in seen:
            seen.add(el)
            out.append(e)
    return ';'.join(out)

# ---- 電話整形 -------------------------------------------------------------------
MOBILE_PREFIX = ('070', '080', '090')
SPECIAL_0120 = '0120'
SPECIAL_0800 = '0800'
SPECIAL_0570 = '0570'
SPECIAL_050  = '050'

def only_digits(s: str) -> str:
    return re.sub(r'\D', '', s or '')

def _format_mobile(n: str) -> str:
    # 070/080/090 の11桁
    if len(n) == 11 and n.startswith(MOBILE_PREFIX):
        return f"{n[:3]}-{n[3:7]}-{n[7:]}"
    return n

def _format_ip050(n: str) -> str:
    # 050-xxxx-xxxx の11桁
    if len(n) == 11 and n.startswith(SPECIAL_050):
        return f"{n[:3]}-{n[3:7]}-{n[7:]}"
    return n

def _format_special(n: str) -> str:
    # 0120/0800 は 4-3-3、0570 は 4-2-4 が一般的
    if n.startswith(SPECIAL_0120) and len(n) == 10:
        return f"{n[:4]}-{n[4:7]}-{n[7:]}"
    if n.startswith(SPECIAL_0800) and len(n) == 10:
        return f"{n[:4]}-{n[4:7]}-{n[7:]}"
    if n.startswith(SPECIAL_0570) and len(n) == 10:
        return f"{n[:4]}-{n[4:6]}-{n[6:]}"
    return n

def _format_fixed_by_areacode(n: str) -> str:
    # 10桁固定電話 前提。市外局番は最長一致（AREA_CODES は長い順）
    if len(n) != 10 or not n.startswith('0'):
        return n
    for code in AREA_CODES:
        if n.startswith(code):
            rest = n[len(code):]
            # 典型パターン
            if len(code) == 2 and len(rest) == 8:
                return f"{code}-{rest[:4]}-{rest[4:]}"   # 03/06 → 4-4
            if len(code) == 3 and len(rest) == 7:
                return f"{code}-{rest[:3]}-{rest[3:]}"   # 3桁 → 3-4
            if len(code) == 4 and len(rest) == 6:
                return f"{code}-{rest[:2]}-{rest[2:]}"   # 4桁 → 2-4
            # 安全フォールバック
            if 6 <= len(rest) <= 8:
                mid = len(rest) - 4
                return f"{code}-{rest[:mid]}-{rest[mid:]}"
            return n
    # 未ヒット（想定外10桁）
    return f"{n[:2]}-{n[2:6]}-{n[6:]}"

def normalize_single_phone(raw: str) -> str:
    if not raw:
        return ''
    n = only_digits(raw)

    # 9桁の固定電話が「先頭0欠落」のケース例：364419772 → 03 + 64419772
    # ルール：先頭が0でなく9桁なら 先頭に '0' を付与（国内固定の典型）
    if len(n) == 9 and not n.startswith('0'):
        n = '0' + n

    # 携帯
    if len(n) == 11 and n.startswith(MOBILE_PREFIX):
        return _format_mobile(n)

    # IP (050)
    if len(n) == 11 and n.startswith(SPECIAL_050):
        return _format_ip050(n)

    # 特番
    if len(n) == 10 and (n.startswith(SPECIAL_0120) or n.startswith(SPECIAL_0800) or n.startswith(SPECIAL_0570)):
        return _format_special(n)

    # 固定（10桁）
    if len(n) == 10 and n.startswith('0'):
        return _format_fixed_by_areacode(n)

    # 8桁や内線、国際、未知は無加工（最後にまとめて連結）
    return n

def normalize_phones(*vals):
    tokens = []
    # 値をばらす（" ::: " / セミコロン / カンマ / スペース / 全角セミコロン）
    for v in vals:
        if not v:
            continue
        parts = re.split(r'\s*:::\s*|[;,、；]\s*|\s+', str(v))
        for p in parts:
            p = p.strip()
            if p:
                tokens.append(p)

    out = []
    seen = set()
    for t in tokens:
        fmt = normalize_single_phone(t)
        if not fmt:
            continue
        if fmt not in seen:
            seen.add(fmt)
            out.append(fmt)
    return ';'.join(out)

# ---- 住所分配（Work/Home/Other） ------------------------------------------------
def build_address(line):
    # Region + City + Street を連結 → 全角化 → 最初の空白で 住所1/住所2
    region = (line.get('address_1__region') or '').strip()
    city   = (line.get('address_1__city') or '').strip()
    street = (line.get('address_1__street') or '').strip()

    body_raw = ' '.join([x for x in [region, city, street] if x])
    body_zen = to_zenkaku_except_hyphen(body_raw)

    addr1, addr2 = body_zen, ''
    m = re.search(r'\s+', body_zen)
    if m:
        idx = m.start()
        addr1 = body_zen[:idx]
        addr2 = body_zen[idx:].strip()

    # 郵便番号は Google Export の "address_1__postal_code"
    pcode = (line.get('address_1__postal_code') or '').strip()
    pcode = only_digits(pcode)
    if len(pcode) == 7:  # 郵便は半角ハイフン
        pcode = f"{pcode[:3]}-{pcode[3:]}"
    elif pcode:
        # 7桁以外は一応そのまま返す（変換しない）
        pass

    return pcode, addr1, addr2

def assign_address(outrow, label, pcode, addr1, addr2):
    label = (label or '').strip().lower()
    if label == 'work':
        outrow['会社〒'] = pcode
        outrow['会社住所1'] = addr1
        outrow['会社住所2'] = addr2
        outrow['会社住所3'] = ''
    elif label == 'home':
        outrow['自宅〒'] = pcode
        outrow['自宅住所1'] = addr1
        outrow['自宅住所2'] = addr2
        outrow['自宅住所3'] = ''
    else:
        outrow['その他〒'] = pcode
        outrow['その他住所1'] = addr1
        outrow['その他住所2'] = addr2
        outrow['その他住所3'] = ''

# ---- 会社かな -------------------------------------------------------------------
def simple_en_to_katakana(s: str) -> str:
    # 超簡易：EN_TO_KATAKANA_MAP → FALLBACKで英大文字の単語を置換
    if not s:
        return s
    t = s
    if EN_TO_KATAKANA_MAP:
        for en, ka in EN_TO_KATAKANA_MAP.items():
            # 全角大文字キーにも対応
            t = re.sub(rf'\b{re.escape(en)}\b', ka, t, flags=re.IGNORECASE)
    return t

def apply_company_replace(s: str) -> str:
    t = s or ''
    if COMPANY_REPLACE_MAP:
        for k, v in COMPANY_REPLACE_MAP.items():
            t = t.replace(k, v)
    return t

def guess_company_kana(name: str) -> str:
    if not name:
        return ''
    base = apply_company_replace(name)
    base = simple_en_to_katakana(base)
    # さらに、明らかな全角英字は半→全の差分も拾えるように大文字へ
    #（ここでは深追いしない：外部辞書優先）
    return base

# ---- 出力カラム ------------------------------------------------------------------
OUTPUT_HEADERS = [
    '姓','名','姓かな','名かな','姓名','姓名かな','ミドルネーム','ミドルネームかな','敬称','ニックネーム','旧姓','宛先',
    '自宅〒','自宅住所1','自宅住所2','自宅住所3','自宅電話','自宅IM ID','自宅E-mail','自宅URL','自宅Social',
    '会社〒','会社住所1','会社住所2','会社住所3','会社電話','会社IM ID','会社E-mail','会社URL','会社Social',
    'その他〒','その他住所1','その他住所2','その他住所3','その他電話','その他IM ID','その他E-mail','その他URL','その他Social',
    '会社名かな','会社名','部署名1','部署名2','役職名','連名','連名ふりがな','連名敬称','連名誕生日',
    'メモ1','メモ2','メモ3','メモ4','メモ5','備考1','備考2','備考3','誕生日','性別','血液型','趣味','性格'
]

# ---- 変換本体 -------------------------------------------------------------------
def convert(rows):
    out = []

    for line in rows:
        r = {h: '' for h in OUTPUT_HEADERS}

        # --- 名前系（そのまま：ユーザーの既存仕様を踏襲）
        r['姓'] = get(line, 'last_name')
        r['名'] = get(line, 'first_name')
        r['姓かな'] = get(line, 'phonetic_last_name')
        r['名かな'] = get(line, 'phonetic_first_name')
        r['姓名'] = (r['姓'] + '　' + r['名']).strip()
        r['姓名かな'] = (r['姓かな'] + '　' + r['名かな']).strip()
        r['敬称'] = '様'
        r['ニックネーム'] = get(line, 'nickname')
        r['旧姓'] = ''
        r['宛先'] = '会社'

        # --- 会社名・かな・部署・役職
        org_name = get(line, 'organization_name')
        org_dept = get(line, 'organization_department')
        org_title = get(line, 'organization_title')

        r['会社名'] = org_name
        r['会社名かな'] = guess_company_kana(org_name)
        if org_dept:
            # 部署1/2にすでに「局 名」のように2語ある場合でも安全に片側へ
            parts = re.split(r'\s+', org_dept.strip(), maxsplit=1)
            r['部署名1'] = parts[0]
            r['部署名2'] = parts[1] if len(parts) > 1 else ''
        r['役職名'] = org_title

        # --- メール（Home/Work/Other）
        home_mails = normalize_emails(get(line, 'e_mail_1___value') if get(line, 'e_mail_1___label') == 'Home' else '',
                                      get(line, 'e_mail_2___value') if get(line, 'e_mail_2___label') == 'Home' else '',
                                      get(line, 'e_mail_3___value') if get(line, 'e_mail_3___label') == 'Home' else '')
        work_mails = normalize_emails(get(line, 'e_mail_1___value') if get(line, 'e_mail_1___label') == 'Work' else '',
                                      get(line, 'e_mail_2___value') if get(line, 'e_mail_2___label') == 'Work' else '',
                                      get(line, 'e_mail_3___value') if get(line, 'e_mail_3___label') == 'Work' else '')
        other_mails = normalize_emails(get(line, 'e_mail_1___value') if get(line, 'e_mail_1___label') not in ('Home','Work') else '',
                                       get(line, 'e_mail_2___value') if get(line, 'e_mail_2___label') not in ('Home','Work') else '',
                                       get(line, 'e_mail_3___value') if get(line, 'e_mail_3___label') not in ('Home','Work') else '')
        r['自宅E-mail'] = home_mails
        r['会社E-mail'] = work_mails
        r['その他E-mail'] = other_mails

        # --- 電話（各ラベル領域へ）
        # Work/Home/Other/Mobile などをまとめてラベル単位で
        def phones_by_label(lbl):
            vals = []
            if get(line, 'phone_1___label') == lbl:
                vals.append(get(line, 'phone_1___value'))
            if get(line, 'phone_2___label') == lbl:
                vals.append(get(line, 'phone_2___value'))
            return normalize_phones(*vals)

        # 「Mobile」は会社か自宅どちらに入れるか運用次第だが、従来仕様なら会社電話へ優先併合
        # 会社：Work + Mobile
        r['会社電話'] = normalize_phones(
            get(line, 'phone_1___value') if get(line, 'phone_1___label') == 'Work' else '',
            get(line, 'phone_2___value') if get(line, 'phone_2___label') == 'Work' else '',
            get(line, 'phone_1___value') if get(line, 'phone_1___label') == 'Mobile' else '',
            get(line, 'phone_2___value') if get(line, 'phone_2___label') == 'Mobile' else '',
        )
        # 自宅
        r['自宅電話'] = phones_by_label('Home')
        # その他
        r['その他電話'] = ''
        if not r['会社電話']:
            # Work未指定で Mobile のみなら会社電話でなくその他へ入れたい場合はここを切替
            pass

        # --- 住所：Work/Home/Other の Address 1 のみ対象（Google形式）
        # Work
        if get(line, 'address_1__label') == 'Work':
            p, a1, a2 = build_address(line)
            assign_address(r, 'work', p, a1, a2)
        elif get(line, 'address_1__label') == 'Home':
            p, a1, a2 = build_address(line)
            assign_address(r, 'home', p, a1, a2)
        elif get(line, 'address_1__label'):
            p, a1, a2 = build_address(line)
            assign_address(r, 'other', p, a1, a2)
        # 追加の Address 2 以降があるCSVにも対応したい場合は同様のブロックを増やす

        # --- メモ（Relation: メモ1〜5, Notes→備考1）
        rel_keys = []
        for i in range(1, 6):
            rel_keys.append((f'relation_{i}__label', f'relation_{i}__value'))
        for i, (lk, vk) in enumerate(rel_keys, start=1):
            if (line.get(lk) or '').strip().startswith('メモ'):
                r[f'メモ{i}'] = (line.get(vk) or '').strip()
        r['備考1'] = (get(line, 'notes') or '').strip()

        # --- 誕生日
        r['誕生日'] = (get(line, 'birthday') or '').strip()

        # --- 会社URL / 会社Social 等（現状空埋め・保持）
        # 必要に応じて拡張

        out.append(r)

    return out

# ---- 取込（自動判定） -----------------------------------------------------------
def sniff_dialect_and_read(fb: bytes):
    # 文字コード
    det = chardet.detect(fb)
    enc = det.get('encoding') or 'utf-8'
    # デコード
    txt = fb.decode(enc, errors='replace')

    # 区切り推定（TSV想定も多い）
    try:
        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(txt[:4000], delimiters=[',','\t',';'])
    except Exception:
        # フォールバック：タブ優先
        class Dialect(csv.Dialect):
            delimiter = '\t'
            quotechar = '"'
            escapechar = None
            doublequote = True
            lineterminator = '\n'
            quoting = csv.QUOTE_MINIMAL
        dialect = Dialect()

    reader = csv.reader(io.StringIO(txt), dialect)
    rows = list(reader)
    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    body = rows[1:]

    # ヘッダ正規化
    norm_header = [norm_key(h) for h in header]
    dicts = []
    for rec in body:
        row = {}
        for i, val in enumerate(rec):
            key = norm_header[i] if i < len(norm_header) else f'col{i}'
            row[key] = val
        dicts.append(row)
    return dicts

# ---- ルーティング ----------------------------------------------------------------
@app.route('/', methods=['GET'])
def index():
    return render_template_string(INDEX_HTML)

@app.route('/convert', methods=['POST'])
def do_convert():
    file = request.files.get('file')
    if not file:
        return "⚠️ エラーが発生しました。CSVの形式や文字コードをご確認ください。", 400
    try:
        fb = file.read()
        rows = sniff_dialect_and_read(fb)
        outrows = convert(rows)

        # 出力
        buf = io.StringIO()
        writer = csv.writer(buf, lineterminator='\n')
        writer.writerow(OUTPUT_HEADERS)
        for r in outrows:
            writer.writerow([r.get(h, '') for h in OUTPUT_HEADERS])

        data = buf.getvalue().encode('utf-8-sig')
        return send_file(
            io.BytesIO(data),
            mimetype='text/csv',
            as_attachment=True,
            download_name='google2atena_converted.csv'
        )
    except Exception as e:
        # デバッグしたいときは e を返す
        return "⚠️ エラーが発生しました。CSVの形式や文字コードをご確認ください。", 400

# ---- Render / Gunicorn 用 --------------------------------------------------------
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
