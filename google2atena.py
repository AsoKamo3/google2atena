# google2atena.py  v3.9.18r7e  (Render安定版 / no-pandas)
# - 文字コード自動判定（chardet）
# - 区切り自動検出（csv.Sniffer）
# - キー正規化
# - 住所整形：郵便番号=半角ハイフン、住所本文=全角（数字/英字/記号）、建物分離（先頭スペース/空白でスプリット）
# - Address 1 - Label に応じて Home / Work / Other に振り分け
# - 電話整形：全国市外局番プリセット + 050/0120/0800/0570、;連結、重複排除
# - メール整形：;連結、重複排除、Home/Work/Otherへ
# - メモ抽出：Relation「メモ1〜5」、Notes→備考
# - 会社かな：外部辞書（company_dicts / kanji_word_map / corp_terms）があれば使用。無ければ最小デフォルトでフェールセーフ。
# - 出力：CSV (UTF-8-SIG) ダウンロード
# - バージョン表記付き出力

import csv, io, re, sys, os, chardet

# ========== 共通関数 ==========

def detect_encoding(data):
    return chardet.detect(data)['encoding'] or 'utf-8-sig'

def norm_key(s):
    return re.sub(r'[^a-z0-9]+', '_', s.strip().lower())

def only_digits(s):
    return re.sub(r'\D', '', s)

def to_zenkaku_except_hyphen(s):
    import unicodedata
    result = []
    for ch in s:
        if ch == '-':
            result.append(ch)
        else:
            result.append(unicodedata.normalize('NFKC', ch))
    return ''.join(result)

def get(d, *keys):
    for k in keys:
        if k in d and d[k]:
            return d[k]
    return ''

# ========== 電話番号整形 ==========

from jp_area_codes import AREA_CODES, SPECIAL

def format_phone(num):
    num = only_digits(num)
    if not num:
        return ''
    if num.startswith('0') is False:
        num = '0' + num
    for code in sorted(AREA_CODES, key=len, reverse=True):
        if num.startswith(code):
            rest = num[len(code):]
            if code in SPECIAL:
                return f"{code}-{rest[:4]}-{rest[4:]}" if len(rest) > 4 else f"{code}-{rest}"
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
        if not v: continue
        parts = re.split(r'[;:/\s]+', v)
        for p in parts:
            p = format_phone(p)
            if p and p not in phones:
                phones.append(p)
    return ';'.join(phones)

# ========== 住所整形（Address 1 - Label対応） ==========

def build_address_grouped(line):
    """Address 1 を Home/Work/Other に振り分け"""
    label = (line.get('address_1__label') or '').strip().lower()
    reg = (line.get('address_1__region') or '').strip()
    city = (line.get('address_1__city') or '').strip()
    street = (line.get('address_1__street') or '').strip()
    pc = (line.get('address_1__postal_code') or '').strip()
    body = ' '.join([x for x in [reg, city, street] if x])

    # 郵便番号整形
    pc = only_digits(pc)
    if len(pc) == 7:
        pc = f"{pc[:3]}-{pc[3:]}"

    # 全角変換＋建物分割
    zbody = to_zenkaku_except_hyphen(body)
    addr1, addr2 = zbody, ''
    m = re.search(r'\s+', zbody)
    if m:
        i = m.start()
        addr1, addr2 = zbody[:i], zbody[i:].strip()

    return label, pc, addr1, addr2

# ========== メール整形 ==========

def merge_emails(*vals):
    emails = []
    for v in vals:
        if not v: continue
        parts = re.split(r'[;:/\s]+', v)
        for p in parts:
            p = p.strip()
            if p and '@' in p and p not in emails:
                emails.append(p)
    return ';'.join(emails)

# ========== メモ整形 ==========

def collect_memos(line):
    notes = (line.get('notes') or '').strip()
    memos = []
    for i in range(1, 6):
        label = f'relation_{i}__label'
        value = f'relation_{i}__value'
        if 'メモ' in (line.get(label) or '') and line.get(value):
            memos.append(line[value])
    return memos, notes

# ========== 会社かな生成（フェールセーフ） ==========

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
    name = re.sub(r'[Ａ-Ｚａ-ｚA-Za-z]', lambda m: chr(ord(m.group(0)) + 0xFEE0) if 'A' <= m.group(0) <= 'Z' or 'a' <= m.group(0) <= 'z' else m.group(0), name)
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
            '姓', '名', '姓かな', '名かな', '姓名', '姓名かな',
            'ミドルネーム', 'ミドルネームかな', '敬称', 'ニックネーム', '旧姓',
            '宛先', '自宅〒', '自宅住所1', '自宅住所2', '自宅住所3', '自宅電話',
            '会社〒', '会社住所1', '会社住所2', '会社住所3', '会社電話',
            'その他〒', 'その他住所1', 'その他住所2', 'その他住所3', 'その他電話',
            '会社名かな', '会社名', '部署名1', '部署名2', '役職名',
            'メモ1', 'メモ2', 'メモ3', 'メモ4', 'メモ5', '備考1', '誕生日'
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

        label, pc, a1, a2 = build_address_grouped(line)
        if 'work' in label:
            r['会社〒'], r['会社住所1'], r['会社住所2'] = pc, a1, a2
        elif 'home' in label:
            r['自宅〒'], r['自宅住所1'], r['自宅住所2'] = pc, a1, a2
        else:
            r['その他〒'], r['その他住所1'], r['その他住所2'] = pc, a1, a2

        phones = merge_phones(
            line.get('phone_1__value'),
            line.get('phone_2__value'),
            line.get('phone_3__value')
        )
        r['会社電話'] = phones

        emails = merge_emails(
            line.get('e_mail_1__value'),
            line.get('e_mail_2__value'),
            line.get('e_mail_3__value')
        )
        r['備考1'] = emails

        memos, notes = collect_memos(line)
        for i, m in enumerate(memos[:5], 1):
            r[f'メモ{i}'] = m
        r['備考1'] = notes or r['備考1']
        r['誕生日'] = (line.get('birthday') or '').strip()

        rows.append(r)

    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=list(rows[0].keys()), lineterminator='\n')
    writer.writeheader()
    writer.writerows(rows)
    return output.getvalue()

# ========== Render出力 ==========
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python google2atena.py input.csv > output.csv")
        sys.exit(1)
    input_path = sys.argv[1]
    with open(input_path, 'rb') as f:
        input_bytes = f.read()
    result = convert(input_bytes)
    print("# Google → Atena 変換結果（v3.9.18r7e）\n")
    print(result)
