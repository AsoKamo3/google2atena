# google2atena.py  v3.9.18r7d (Render安定版 / no-pandas)
# 住所／電話／メモ抽出などは v3.9.18r7c と同等。
# この版ではアプリUIにバージョン名を明示表示。

from flask import Flask, request, send_file, render_template_string
import io, csv, re, codecs
import chardet

# ---- 外部辞書（フェイルセーフ） -------------------------------------------------
try:
    from company_dicts import COMPANY_REPLACE_MAP
except Exception:
    COMPANY_REPLACE_MAP = {}

try:
    from kanji_word_map import EN_TO_KATAKANA_MAP
except Exception:
    EN_TO_KATAKANA_MAP = {}

try:
    from corp_terms import CORP_TERMS
except Exception:
    CORP_TERMS = set()

# jp_area_codes は安定版（旧短縮）
from jp_area_codes import AREA_CODES

app = Flask(__name__)

# ---- UI（バージョン表示付き） -----------------------------------------------------
INDEX_HTML = """<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>google2atena v3.9.18r7d</title>
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
  <h1>google2atena v3.9.18r7d</h1>
  <p>（Render安定版 / no-pandas）</p>
  <form action="/convert" method="post" enctype="multipart/form-data">
    <p>Google 連絡先の CSV/TSV をアップロードしてください。</p>
    <input type="file" name="file" accept=".csv,.tsv,.txt" required>
    <div><button type="submit">変換する</button></div>
  </form>
  <p class="small">出力: UTF-8 (BOM付) CSV / 住所は全角（郵便・電話のみ半角ハイフン付）</p>
</div>
</body>
</html>"""

# ---- 共通関数群 ---------------------------------------------------------------
def to_zenkaku_except_hyphen(s: str) -> str:
    if not s:
        return s
    s = s.replace('-', '\u0000')
    tbl = {ord(c): ord(c) + 0xFEE0 for c in ''.join(chr(i) for i in range(33, 127))}
    s = s.translate(tbl)
    s = s.replace('\u0000', '－')
    return s

def norm_key(k: str) -> str:
    return (k or '').strip().lower().replace(' ', '_').replace('-', '_')

def get(row, *keys, default=''):
    for k in keys:
        v = row.get(k)
        if v:
            return v
    return default

def only_digits(s: str) -> str:
    return re.sub(r'\D', '', s or '')

# ---- 電話正規化 ---------------------------------------------------------------
MOBILE_PREFIX = ('070', '080', '090')
SPECIALS = ('0120', '0800', '0570', '050')

def normalize_single_phone(raw: str) -> str:
    if not raw:
        return ''
    n = only_digits(raw)
    if len(n) == 9 and not n.startswith('0'):
        n = '0' + n
    # 携帯
    if len(n) == 11 and n.startswith(MOBILE_PREFIX):
        return f"{n[:3]}-{n[3:7]}-{n[7:]}"
    # IP
    if len(n) == 11 and n.startswith('050'):
        return f"{n[:3]}-{n[3:7]}-{n[7:]}"
    # 特番
    if len(n) == 10 and n.startswith('0120'):
        return f"{n[:4]}-{n[4:7]}-{n[7:]}"
    if len(n) == 10 and n.startswith('0800'):
        return f"{n[:4]}-{n[4:7]}-{n[7:]}"
    if len(n) == 10 and n.startswith('0570'):
        return f"{n[:4]}-{n[4:6]}-{n[6:]}"
    # 固定
    if len(n) == 10 and n.startswith('0'):
        for code in AREA_CODES:
            if n.startswith(code):
                rest = n[len(code):]
                mid = len(rest) - 4
                return f"{code}-{rest[:mid]}-{rest[mid:]}"
    return n

def normalize_phones(*vals):
    parts = []
    for v in vals:
        if not v:
            continue
        for p in re.split(r'\s*:::\s*|[;,、；]\s*|\s+', str(v)):
            if p.strip():
                parts.append(p.strip())
    out = []
    seen = set()
    for p in parts:
        f = normalize_single_phone(p)
        if f and f not in seen:
            seen.add(f)
            out.append(f)
    return ';'.join(out)

# ---- メール正規化 --------------------------------------------------------------
def normalize_emails(*vals):
    out = []
    seen = set()
    for v in vals:
        if not v:
            continue
        for p in re.split(r'\s*:::\s*|[;,\s]\s*', str(v)):
            if p and '@' in p and '.' in p.split('@')[-1]:
                pl = p.lower()
                if pl not in seen:
                    seen.add(pl)
                    out.append(p)
    return ';'.join(out)

# ---- 住所 ----------------------------------------------------------------------
def build_address(line):
    # Region / City / Street を順に探索（Work/Home/Other いずれにも対応）
    reg = get(line,
        'address_1__region', 'address_home__region', 'address_work__region', 'address_other__region'
    ).strip()
    city = get(line,
        'address_1__city', 'address_home__city', 'address_work__city', 'address_other__city'
    ).strip()
    street = get(line,
        'address_1__street', 'address_home__street', 'address_work__street', 'address_other__street'
    ).strip()
    body = ' '.join([x for x in [reg, city, street] if x])

    # 全角変換
    zbody = to_zenkaku_except_hyphen(body)

    # 建物分割（最初のスペースで分割）
    addr1, addr2 = zbody, ''
    m = re.search(r'\s+', zbody)
    if m:
        i = m.start()
        addr1, addr2 = zbody[:i], zbody[i:].strip()

    # 郵便番号（7桁→3-4形式）
    pc = get(line,
        'address_1__postal_code', 'address_home__postal_code', 'address_work__postal_code', 'address_other__postal_code'
    ).strip()
    pc = only_digits(pc)
    if len(pc) == 7:
        pc = f"{pc[:3]}-{pc[3:]}"
    return pc, addr1, addr2

# ---- 会社名かな ---------------------------------------------------------------
def simple_en_to_katakana(s: str) -> str:
    if not s:
        return s
    t = s
    if EN_TO_KATAKANA_MAP:
        for en, ka in EN_TO_KATAKANA_MAP.items():
            t = re.sub(rf'\b{re.escape(en)}\b', ka, t, flags=re.IGNORECASE)
    return t

def apply_company_replace(s: str) -> str:
    t = s or ''
    for k, v in COMPANY_REPLACE_MAP.items():
        t = t.replace(k, v)
    return t

def guess_company_kana(name: str) -> str:
    if not name:
        return ''
    return simple_en_to_katakana(apply_company_replace(name))

# ---- 出力カラム ---------------------------------------------------------------
OUTPUT_HEADERS = [
    '姓','名','姓かな','名かな','姓名','姓名かな','ミドルネーム','ミドルネームかな','敬称','ニックネーム','旧姓','宛先',
    '自宅〒','自宅住所1','自宅住所2','自宅住所3','自宅電話','自宅IM ID','自宅E-mail','自宅URL','自宅Social',
    '会社〒','会社住所1','会社住所2','会社住所3','会社電話','会社IM ID','会社E-mail','会社URL','会社Social',
    'その他〒','その他住所1','その他住所2','その他住所3','その他電話','その他IM ID','その他E-mail','その他URL','その他Social',
    '会社名かな','会社名','部署名1','部署名2','役職名','連名','連名ふりがな','連名敬称','連名誕生日',
    'メモ1','メモ2','メモ3','メモ4','メモ5','備考1','備考2','備考3','誕生日','性別','血液型','趣味','性格'
]

# ---- 変換本体 ---------------------------------------------------------------
def convert(rows):
    out = []
    for line in rows:
        r = {h: '' for h in OUTPUT_HEADERS}
        r['姓'] = get(line, 'last_name')
        r['名'] = get(line, 'first_name')
        r['姓かな'] = get(line, 'phonetic_last_name')
        r['名かな'] = get(line, 'phonetic_first_name')
        r['姓名'] = (r['姓'] + '　' + r['名']).strip()
        r['姓名かな'] = (r['姓かな'] + '　' + r['名かな']).strip()
        r['敬称'] = '様'
        r['宛先'] = '会社'

        org = get(line, 'organization_name')
        dept = get(line, 'organization_department')
        title = get(line, 'organization_title')
        r['会社名'] = org
        r['会社名かな'] = guess_company_kana(org)
        if dept:
            parts = re.split(r'\s+', dept.strip(), 1)
            r['部署名1'] = parts[0]
            r['部署名2'] = parts[1] if len(parts) > 1 else ''
        r['役職名'] = title

        r['会社電話'] = normalize_phones(
            get(line, 'phone_1___value') if get(line, 'phone_1___label') == 'Work' else '',
            get(line, 'phone_2___value') if get(line, 'phone_2___label') == 'Work' else '',
            get(line, 'phone_1___value') if get(line, 'phone_1___label') == 'Mobile' else '',
            get(line, 'phone_2___value') if get(line, 'phone_2___label') == 'Mobile' else ''
        )
        r['自宅電話'] = normalize_phones(
            get(line, 'phone_1___value') if get(line, 'phone_1___label') == 'Home' else '',
            get(line, 'phone_2___value') if get(line, 'phone_2___label') == 'Home' else ''
        )

        r['会社E-mail'] = normalize_emails(
            get(line, 'e_mail_1___value') if get(line, 'e_mail_1___label') == 'Work' else '',
            get(line, 'e_mail_2___value') if get(line, 'e_mail_2___label') == 'Work' else ''
        )
        r['自宅E-mail'] = normalize_emails(
            get(line, 'e_mail_1___value') if get(line, 'e_mail_1___label') == 'Home' else '',
            get(line, 'e_mail_2___value') if get(line, 'e_mail_2___label') == 'Home' else ''
        )

        pc, a1, a2 = build_address(line)
        r['会社〒'], r['会社住所1'], r['会社住所2'] = pc, a1, a2

        r['備考1'] = get(line, 'notes')
        r['誕生日'] = get(line, 'birthday')
        out.append(r)
    return out

# ---- 取込・出力 ---------------------------------------------------------------
def sniff_dialect_and_read(fb: bytes):
    det = chardet.detect(fb)
    enc = det.get('encoding') or 'utf-8'
    txt = fb.decode(enc, errors='replace')
    try:
        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(txt[:4000], delimiters=[',','\t',';'])
    except Exception:
        class D(csv.Dialect):
            delimiter='\t';quotechar='"';escapechar=None;doublequote=True
            lineterminator='\n';quoting=csv.QUOTE_MINIMAL
        dialect=D()
    reader = csv.reader(io.StringIO(txt), dialect)
    rows = list(reader)
    if not rows:
        raise ValueError
    header = [norm_key(h) for h in rows[0]]
    dicts = []
    for rec in rows[1:]:
        d = {}
        for i,v in enumerate(rec):
            if i<len(header): d[header[i]]=v
        dicts.append(d)
    return dicts

@app.route('/', methods=['GET'])
def index():
    return render_template_string(INDEX_HTML)

@app.route('/convert', methods=['POST'])
def convert_csv():
    f = request.files.get('file')
    if not f: return "ファイルを選択してください",400
    try:
        fb = f.read()
        rows = sniff_dialect_and_read(fb)
        outrows = convert(rows)
        buf = io.StringIO()
        w = csv.writer(buf, lineterminator='\n')
        w.writerow(OUTPUT_HEADERS)
        for r in outrows: w.writerow([r.get(h,'') for h in OUTPUT_HEADERS])
        data = buf.getvalue().encode('utf-8-sig')
        return send_file(io.BytesIO(data),mimetype='text/csv',as_attachment=True,download_name='google2atena_converted.csv')
    except Exception as e:
        return "⚠️ エラーが発生しました。CSVの形式や文字コードをご確認ください。",400

if __name__=='__main__':
    app.run(host='0.0.0.0',port=10000)
