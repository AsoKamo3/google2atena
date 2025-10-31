# google2atena.py  v3.9.18r7d+addrfix+webapp  (Render安定版 / no-pandas)
# - v3.9.18r7d をベースに、住所部分のみ修正（Address 1 - Label 判定）
# - Flaskアプリ構成を維持
# - 電話／メール／メモ／誕生日／会社かなロジックは r7d と同一
# - 出力：CSV (UTF-8-SIG) ダウンロード対応

import csv, io, re, sys, os, chardet, unicodedata, tempfile
from flask import Flask, request, send_file

# ========== 共通関数 ==========

def detect_encoding(data):
    return chardet.detect(data)['encoding'] or 'utf-8-sig'

def norm_key(s):
    return re.sub(r'[^a-z0-9]+', '_', s.strip().lower())

def only_digits(s):
    return re.sub(r'\D', '', s or '')

def to_zenkaku_except_hyphen(s):
    out = []
    for ch in s or '':
        if ch == '-':
            out.append(ch)
        else:
            out.append(unicodedata.normalize('NFKC', ch))
    return ''.join(out)

# ========== jp_area_codes 読み込み（フェイルセーフ） ==========

try:
    from jp_area_codes import AREA_CODES, SPECIAL
except Exception:
    from jp_area_codes import AREA_CODES
    SPECIAL = {"0120", "0800", "0570", "050"}

# ========== 電話番号整形（r7dそのまま） ==========

def format_phone(num):
    num = only_digits(num)
    if not num:
        return ''
    if num[0] != '0':
        num = '0' + num
    for code in sorted(AREA_CODES, key=len, reverse=True):
        if num.startswith(code):
            rest = num[len(code):]
            if code in SPECIAL:
                if len(rest) > 4:
                    return f"{code}-{rest[:4]}-{rest[4:]}"
                return f"{code}-{rest}"
            if len(rest) >= 7:
                return f"{code}-{rest[:-4]}-{rest[-4:]}"
            elif len(rest) >= 4:
                return f"{code}-{rest[:len(rest)-4]}-{rest[-4:]}"
            else:
                return f"{code}-{rest}"
    if len(num) == 10:
        return f"{num[:3]}-{num[3:6]}-{num[6:]}"
    elif len(num) == 11:
        return f"{num[:3]}-{num[3:7]}-{num[7:]}"
    return num

def merge_phones(*vals):
    phones = []
    for v in vals:
        if not v:
            continue
        parts = re.split(r'[;:/\s]+', v)
        for p in parts:
            f = format_phone(p)
            if f and f not in phones:
                phones.append(f)
    return ';'.join(phones)

# ========== 住所ユーティリティ（修正版） ==========

def _format_postal(pc: str) -> str:
    d = only_digits(pc)
    return f"{d[:3]}-{d[3:]}" if len(d) == 7 else (pc or '')

def _concat_region_city_street(line: dict) -> str:
    region = (line.get('address_1__region') or '').strip()
    city   = (line.get('address_1__city') or '').strip()
    street = (line.get('address_1__street') or '').strip()
    parts = [p for p in [region, city, street] if p]
    return ' '.join(parts)

def _split_building_first_space(full_zenkaku: str):
    m = re.search(r'\s+', full_zenkaku or '')
    if m:
        i = m.start()
        return full_zenkaku[:i], full_zenkaku[i:].strip()
    return full_zenkaku, ''

def build_address_from_address1_fields(line: dict):
    """Address 1 - Label に応じて Home/Work/Other を判定"""
    label_raw = (line.get('address_1__label') or '').strip().lower()
    postal = _format_postal(line.get('address_1__postal_code') or '')
    body_raw = _concat_region_city_street(line)
    body_z = to_zenkaku_except_hyphen(body_raw)
    addr1, addr2 = _split_building_first_space(body_z)

    if 'work' in label_raw:
        label = 'work'
    elif 'home' in label_raw:
        label = 'home'
    else:
        label = 'other'
    return label, postal, addr1, addr2

# ========== メール整形（r7dそのまま） ==========

def merge_emails(*vals):
    emails = []
    for v in vals:
        if not v:
            continue
        parts = re.split(r'[;:/\s]+', v)
        for p in parts:
            p = p.strip()
            if '@' in p and p not in emails:
                emails.append(p)
    return ';'.join(emails)

# ========== メモ整形（r7dそのまま） ==========

def collect_memos(line):
    notes = (line.get('notes') or '').strip()
    memos = []
    for i in range(1, 6):
        label = f'relation_{i}__label'
        value = f'relation_{i}__value'
        if 'メモ' in (line.get(label) or '') and line.get(value):
            memos.append(line[value])
    return memos, notes

# ========== 会社かな生成（r7dそのまま） ==========

def company_kana(name):
    if not name:
        return ''
    try:
        from company_dicts import company_kana_map
        for k, v in company_kana_map.items():
            if k in name:
                return v
    except Exception:
        pass
    name = re.sub(r'[A-Za-z]', lambda m: chr(ord(m.group(0)) + 0xFEE0), name)
    return name

# ========== メイン変換処理 ==========

def convert(input_bytes):
    enc = detect_encoding(input_bytes)
    text = input_bytes.decode(enc, errors='ignore')
    sniffer = csv.Sniffer()
    dialect = sniffer.sniff(text.splitlines()[0])
    reader = csv.DictReader(io.StringIO(text), dialect=dialect)
    rows = []

    for line in reader:
        line = {norm_key(k): v for k, v in line.items()}

        r = {k: '' for k in [
            '姓','名','姓かな','名かな','姓名','姓名かな','ミドルネーム','ミドルネームかな','敬称',
            'ニックネーム','旧姓','宛先','自宅〒','自宅住所1','自宅住所2','自宅住所3','自宅電話',
            '会社〒','会社住所1','会社住所2','会社住所3','会社電話',
            'その他〒','その他住所1','その他住所2','その他住所3','その他電話',
            '会社名かな','会社名','部署名1','部署名2','役職名',
            'メモ1','メモ2','メモ3','メモ4','メモ5','備考1','誕生日'
        ]}

        r['姓'] = (line.get('last_name') or '').strip()
        r['名'] = (line.get('first_name') or '').strip()
        r['姓かな'] = (line.get('phonetic_last_name') or '').strip()
        r['名かな'] = (line.get('phonetic_first_name') or '').strip()
        r['姓名'] = f"{r['姓']}　{r['名']}".strip()
        r['姓名かな'] = f"{r['姓かな']}　{r['名かな']}".strip()
        r['敬称'] = '様'
        r['会社名'] = (line.get('organization_name') or '').strip()
        r['会社名かな'] = company_kana(r['会社名'])
        r['部署名1'] = (line.get('organization_department') or '').strip()
        r['役職名'] = (line.get('organization_title') or '').strip()

        # --- 住所部分（修正版） ---
        label, pc, a1, a2 = build_address_from_address1_fields(line)
        if label == 'work':
            r['会社〒'], r['会社住所1'], r['会社住所2'] = pc, a1, a2
        elif label == 'home':
            r['自宅〒'], r['自宅住所1'], r['自宅住所2'] = pc, a1, a2
        else:
            r['その他〒'], r['その他住所1'], r['その他住所2'] = pc, a1, a2

        # 電話
        r['会社電話'] = merge_phones(
            line.get('phone_1__value'),
            line.get('phone_2__value'),
            line.get('phone_3__value')
        )

        # メール
        r['備考1'] = merge_emails(
            line.get('e_mail_1__value'),
            line.get('e_mail_2__value'),
            line.get('e_mail_3__value')
        )

        # メモ＋備考
        memos, notes = collect_memos(line)
        for i, m in enumerate(memos[:5], 1):
            r[f'メモ{i}'] = m
        if notes:
            r['備考1'] = notes
        r['誕生日'] = (line.get('birthday') or '').strip()

        rows.append(r)

    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=list(rows[0].keys()), lineterminator='\n')
    writer.writeheader()
    writer.writerows(rows)
    return output.getvalue()

# ========== Flaskアプリ部分 ==========

app = Flask(__name__)

@app.route("/", methods=["GET"])
def index():
    return f"""
    <h2>Google → Atena 変換ツール</h2>
    <p>Version: <b>v3.9.18r7d+addrfix+webapp</b></p>
    <form action="/convert" method="post" enctype="multipart/form-data">
      <input type="file" name="file" accept=".csv" required>
      <button type="submit">変換実行</button>
    </form>
    """

@app.route("/convert", methods=["POST"])
def convert_route():
    f = request.files.get("file")
    if not f:
        return "⚠️ ファイルが指定されていません。", 400
    data = f.read()
    try:
        result = convert(data)
    except Exception as e:
        return f"⚠️ エラーが発生しました: {e}", 500

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
    tmp.write(result.encode("utf-8-sig"))
    tmp.seek(0)
    return send_file(tmp.name, as_attachment=True, download_name="converted.csv", mimetype="text/csv")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
