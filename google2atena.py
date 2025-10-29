# -*- coding: utf-8 -*-
"""
google2atena v3.9.16 aligned full-format no-pandas (import済・統合版)
- Flask + 標準csvのみ（pandas不使用）
- 住所整形：自宅/その他/会社すべてに適用（建物分割・全角化）
- 電話：全国市外局番 + 050/0120/0800/0570 等の特番帯をプリセット、半角数字+半角ハイフンに正規化
- 郵便番号：半角数字+半角ハイフン
- メール：区切り正規化・重複排除・;結合
- 誕生日：YYYY/MM/DD に統一
- 会社名かな：辞書/ルールベース（会社かなは全角、英字・記号除去、法人格除去）
- 出力列：理想型に完全一致
"""

import io
import re
import csv
import html
import unicodedata
from collections import OrderedDict

from flask import Flask, request, Response

# ---- imports (外部ファイル) -------------------------------------------------
from company_dicts import COMPANY_EXCEPT
from kanji_word_map import KANJI_WORD_MAP
from corp_terms import CORP_TERMS
from jp_area_codes import AREA_CODES  # 市外局番(0xx, 0xxx, 0xxxx) + 特番帯: 0120/0800/0570/050 等

app = Flask(__name__)

# ==== ユーティリティ ==========================================================

def to_fullwidth(s: str) -> str:
    """半角英数記号→全角（スペースやASCII記号も含む）。
    ※電話番号/郵便番号の正規化には使わないこと（別関数で半角化）。
    """
    if not s:
        return s
    out = []
    for ch in s:
        code = ord(ch)
        # 半角スペース
        if ch == ' ':
            out.append('　')
            continue
        # ASCII 33-126 を全角へ（ハイフンも全角へ）
        if 33 <= code <= 126:
            out.append(chr(code + 0xFEE0))
        else:
            out.append(ch)
    return ''.join(out)

def zenkaku_numbers_and_ascii(s: str) -> str:
    """アドレス系で「数字・英字・記号」を含めて全角化（住所専用）。
    ハイフンも全角へ。"""
    return to_fullwidth(s)

def to_half_digits(s: str) -> str:
    """全角数字を半角数字へ。他はそのまま。"""
    if not s:
        return s
    return ''.join(unicodedata.normalize('NFKC', ch) if ch.isdigit() else ch for ch in s)

def compact_spaces(s: str) -> str:
    return re.sub(r'[ \t\u3000]+', ' ', s or '').strip()

def split_multi(value: str) -> list:
    """メール等の複数区切り： ::: / ; / ， / 、 / , / 空白 を許容"""
    if not value:
        return []
    value = value.replace(':::', ';').replace('，', ',').replace('、', ',')
    parts = re.split(r'[;\s,]+', value)
    return [p.strip() for p in parts if p.strip()]

def uniq_preserve(seq):
    seen = set()
    out = []
    for x in seq:
        if x not in seen:
            out.append(x)
            seen.add(x)
    return out

# ==== 郵便番号 ===============================================================

def normalize_postal(raw: str) -> str:
    """郵便番号：半角数字と半角ハイフン、7桁→ NNN-NNNN"""
    if not raw:
        return ''
    digits = re.sub(r'\D', '', raw)
    if len(digits) == 7:
        return f"{digits[:3]}-{digits[3:]}"
    if len(digits) == 3 or len(digits) == 4:  # 稀に3-4で来る
        # 可能なら 3-4 に整形
        return '-'.join([digits[:3], digits[3:]]) if len(digits) == 7 else raw
    # それ以外は原文の数字並びを返す（無理に壊さない）
    return digits or ''

# ==== 電話番号（全国＋特番）==================================================

_SPECIAL_PREFIXES = ("0120", "0800", "0570", "050")
_MOBILE_PREFIXES = ("070", "080", "090")

def _split_to_digits(s: str) -> str:
    return re.sub(r'\D', '', s or '')

def _format_with_hyphens(digits: str) -> str:
    """市外局番辞書 + 特番帯 + 携帯帯で適切に 0A-XXXX-XXXX 等に整形（半角）"""
    if not digits:
        return ''
    # まず特番帯（固定長）
    for p in _SPECIAL_PREFIXES:
        if digits.startswith(p):
            # 0120/0800 -> 4-3-3、0570 -> 4-2-4（一般的）
            if p in ("0120", "0800"):
                if len(digits) == 10:
                    return f"{digits[:4]}-{digits[4:7]}-{digits[7:]}"
                # 非10桁はそのままの数字列
                return digits
            if p == "0570":
                if len(digits) == 10:
                    return f"{digits[:4]}-{digits[4:6]}-{digits[6:]}"
                return digits
            if p == "050":
                # VoIP 050-XXXX-XXXX （11桁想定）
                if len(digits) == 11:
                    return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
                return digits

    # 携帯 070/080/090 -> 3-4-4（11桁）
    for m in _MOBILE_PREFIXES:
        if digits.startswith(m) and len(digits) == 11:
            return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"

    # 固定電話：市外局番辞書（4/3/2桁の優先でマッチ）
    # AREA_CODES: set of strings like "011","015","0174",...
    # 4桁コード優先
    for k in (4, 3, 2):  # '0'込みで 4/3/2 桁を想定（例: 011,03）
        prefix = digits[:k]
        if prefix in AREA_CODES:
            rest = digits[k:]
            # 残り桁数に応じてハイフン位置（基本 3-4 or 2-4 等）
            if len(rest) == 8:  # 2-4-2 のような異常は稀、ここは通常 4+4 と判断
                # 03 + 1234 + 5678
                return f"{prefix}-{rest[:4]}-{rest[4:]}"
            if len(rest) == 7:
                # 011 + 123 + 4567 / 048 + 123 + 4567 等：3-4
                return f"{prefix}-{rest[:3]}-{rest[3:]}"
            if len(rest) == 6:
                # 099 + 12 + 3456：2-4
                return f"{prefix}-{rest[:2]}-{rest[2:]}"
            # 想定外長さは一旦 prefix-残り で返す
            return f"{prefix}-{rest}" if rest else prefix

    # ここまで来たら一般的な長さでのフォールバック
    if len(digits) == 10:
        return f"{digits[:2]}-{digits[2:6]}-{digits[6:]}"
    if len(digits) == 9:
        return f"{digits[:2]}-{digits[2:5]}-{digits[5:]}"
    if len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    return digits  # 変に壊さない

def normalize_phone_field(raw: str) -> list:
    """複数値を ; / ::: / , / 空白 などから抽出 → 整形（半角数字+半角ハイフン）"""
    if not raw:
        return []
    chunks = split_multi(raw)
    out = []
    for ch in chunks:
        digits = _split_to_digits(ch)
        if not digits:
            continue
        out.append(_format_with_hyphens(digits))
    return uniq_preserve([x for x in out if x])

# ==== メール ================================================================

def normalize_emails(*values) -> str:
    """複数カラム値をまとめ、区切り正規化・小文字化・重複排除 → ; 連結"""
    all_parts = []
    for v in values:
        all_parts.extend(split_multi(v))
    # トリム・小文字
    all_parts = [p.strip().lower() for p in all_parts if p.strip()]
    # 余計な全角記号を半角に（メールは半角ASCIIで）
    normalized = []
    for p in all_parts:
        p = unicodedata.normalize('NFKC', p)
        normalized.append(p)
    return ';'.join(uniq_preserve(normalized))

# ==== 誕生日 ================================================================

def normalize_birthday(raw: str) -> str:
    """YYYY/MM/DD へ。ISO/和暦/区切り揺れをできる範囲で吸収。"""
    if not raw:
        return ''
    s = raw.strip()
    s = s.replace('年', '/').replace('月', '/').replace('日', '')
    s = re.sub(r'[\.\-]', '/', s)
    s = re.sub(r'\s+', '', s)
    m = re.match(r'^(\d{4})/(\d{1,2})/(\d{1,2})$', s)
    if m:
        y, mm, dd = m.groups()
        return f"{int(y):04d}/{int(mm):02d}/{int(dd):02d}"
    # 他形式（MM/DD/YYYY等）は最小限対応
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', s)
    if m:
        mm, dd, y = m.groups()
        return f"{int(y):04d}/{int(mm):02d}/{int(dd):02d}"
    return s  # 無理に壊さない

# ==== 会社名かな ============================================================

_PUNCT_TO_DROP = '・．。,，.'
def strip_corp_terms(name: str) -> str:
    s = name or ''
    s = s.strip()
    # 全角/半角スペース正規化
    s = re.sub(r'[\u3000 ]+', ' ', s)
    # 法人格を削除（語頭/語中の独立語として）
    for term in sorted(CORP_TERMS, key=len, reverse=True):
        s = s.replace(term, ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def company_to_kana(name: str) -> str:
    """会社名→会社名かな
    - 全角化
    - 「・ . ，」削除
    - 法人格除去
    - 例外辞書 → KANJI_WORD_MAP 置換 → ローマ字/英字の読み（簡易）
    """
    if not name:
        return ''
    # 例外（フル一致）先に
    nm = name.strip()
    if nm in COMPANY_EXCEPT:
        base = COMPANY_EXCEPT[nm]
    else:
        base = nm

    # 法人格除去
    base = strip_corp_terms(base)

    # 句読点/中黒/ピリオド等の削除
    for p in _PUNCT_TO_DROP:
        base = base.replace(p, '')

    # 英字は読み上げ（A→エー, B→ビー...）にし、最終的には全角で
    def roman_to_reading(m):
        token = m.group(0)
        # 既に全角の英字は NFKCで半角→読みへ
        token_ascii = unicodedata.normalize('NFKC', token).upper()
        table = {
            'A':'エー','B':'ビー','C':'シー','D':'ディー','E':'イー',
            'F':'エフ','G':'ジー','H':'エイチ','I':'アイ','J':'ジェイ',
            'K':'ケー','L':'エル','M':'エム','N':'エヌ','O':'オー',
            'P':'ピー','Q':'キュー','R':'アール','S':'エス','T':'ティー',
            'U':'ユー','V':'ブイ','W':'ダブリュー','X':'エックス','Y':'ワイ','Z':'ゼット'
        }
        return ''.join(table.get(ch, ch) for ch in token_ascii if ch.isalpha())

    base = re.sub(r'[A-Za-zＡ-Ｚａ-ｚ]+', roman_to_reading, base)

    # KANJI_WORD_MAP で置換（長い語から）
    if KANJI_WORD_MAP:
        for k in sorted(KANJI_WORD_MAP.keys(), key=len, reverse=True):
            base = base.replace(k, KANJI_WORD_MAP[k])

    # 最後に全角化（会社かなは全角）
    base = to_fullwidth(base)

    # 全角の & / ハイフンなど余分は基本残すが、中黒は前段で除去済み
    # 仕上げ：全角スペースのトリム
    base = re.sub(r'[\u3000 ]+', ' ', base).strip()
    return base

# ==== 住所（建物分割・全角化）===============================================

_BUILDING_TOKENS = [
    'ビル','タワー','マンション','ハイツ','コーポ','レジデンス','荘','寮','アパート',
    '建物','館','棟','号棟','#','＃','階','F','１F','２F','３F','4F','5F','６F',
    'フロア','室','号','丁目','番地','番','号室'
]

def split_building(addr1: str) -> (str, str):
    """住所の通り/番地 と 建物系を分割。簡易ルール（壊さない前提）。
    - 最初に現れる建物トークン以降を建物側へ
    """
    s = addr1 or ''
    s = s.strip()
    if not s:
        return '', ''
    # 半角→全角化（住所は全角が要件）
    s_fw = zenkaku_numbers_and_ascii(s)
    # 分割位置探索（先に #/＃）
    for token in ['＃', '#']:
        if token in s_fw:
            i = s_fw.index(token)
            return s_fw[:i].strip(), s_fw[i:].strip()
    # 建物トークン
    for tk in _BUILDING_TOKENS:
        if tk in s_fw:
            i = s_fw.index(tk)
            return s_fw[:i].strip(), s_fw[i:].strip()
    return s_fw, ''

def compose_address(city: str, region: str, street: str, ext: str) -> (str, str):
    """Googleエクスポートの分割要素から、
       住所1=「都道府県+市区町村+番地（全角化）」、住所2=「建物（全角化）」 を作る
    """
    parts = []
    for p in (region, city, street):
        p = (p or '').strip()
        if p:
            parts.append(p)
    joined = ' '.join(parts).strip()

    # 住所本体→建物分割
    a1, bld = split_building(joined)

    # Address 1 - Extended Address（ext）は建物の補助として住所2に寄せる
    bld2 = ' '.join([x for x in [bld, (ext or '').strip()] if x]).strip()

    # 最終：住所1/2 は全角化済。住所3は空。
    return a1, bld2

# ==== 行変換 ================================================================

OUTPUT_COLUMNS = [
    "姓","名","姓かな","名かな","姓名","姓名かな","ミドルネーム","ミドルネームかな",
    "敬称","ニックネーム","旧姓","宛先",
    "自宅〒","自宅住所1","自宅住所2","自宅住所3","自宅電話","自宅IM ID","自宅E-mail","自宅URL","自宅Social",
    "会社〒","会社住所1","会社住所2","会社住所3","会社電話","会社IM ID","会社E-mail","会社URL","会社Social",
    "その他〒","その他住所1","その他住所2","その他住所3","その他電話","その他IM ID","その他E-mail","その他URL","その他Social",
    "会社名かな","会社名","部署名1","部署名2","役職名",
    "連名","連名ふりがな","連名敬称","連名誕生日",
    "メモ1","メモ2","メモ3","メモ4","メモ5","備考1","備考2","備考3",
    "誕生日","性別","血液型","趣味","性格"
]

def collect_memos(row: dict):
    """Relation i - Label/Value + Notes から メモ/備考抽出。欠落なし優先。"""
    memo = ['','','','','']  # メモ1-5
    biko = ['','','']        # 備考1-3

    # Notes 全体は優先的に メモ1 に入れる（既にメモ1が埋まっている場合は空いている先頭へ）
    notes = (row.get('Notes','') or '').strip()
    if notes:
        placed = False
        for i in range(5):
            if not memo[i]:
                memo[i] = notes
                placed = True
                break
        if not placed:
            # あふれたら備考1へ
            if not biko[0]:
                biko[0] = notes

    # Relation i - Label/Value
    for i in range(1, 10):
        lbl = row.get(f"Relation {i} - Label", "") or ""
        val = row.get(f"Relation {i} - Value", "") or ""
        if not (lbl or val):
            continue
        lbl = lbl.strip()
        if not val:
            val = ''
        # 「メモn」「備考n」など
        m = re.match(r'メモ\s*([１-５1-5])', lbl)
        if m:
            idx = int(unicodedata.normalize('NFKC', m.group(1)))
            if 1 <= idx <= 5 and not memo[idx-1]:
                memo[idx-1] = val
            else:
                # 次の空きへ
                for k in range(5):
                    if not memo[k]:
                        memo[k] = val
                        break
            continue
        b = re.match(r'備考\s*([１-３1-3])', lbl)
        if b:
            idx = int(unicodedata.normalize('NFKC', b.group(1)))
            if 1 <= idx <= 3 and not biko[idx-1]:
                biko[idx-1] = val
            else:
                for k in range(3):
                    if not biko[k]:
                        biko[k] = val
                        break
            continue
        # ラベルがメモ系でなくても内容は拾ってどこかに置く（欠落防止）
        for k in range(5):
            if not memo[k]:
                memo[k] = f"{lbl}　{val}".strip()
                break

    return memo, biko

def normalize_company_fields(org_name: str, dept: str, title: str):
    """会社名かな/会社名/部署名/役職名 整形。部署・役職は全角化（スペースも全角）。"""
    org = (org_name or '').strip()
    org_kana = company_to_kana(org) if org else ''
    # 部署は「部署名1」「部署名2」に分割（全角化・空白は全角スペース）
    d1, d2 = '', ''
    d = compact_spaces(dept)
    if d:
        d = to_fullwidth(d)
        # スペースで2分割（過剰分割はしない）
        if ' ' in d:
            d1, d2 = d.split(' ', 1)
        else:
            d1 = d
    # 役職は全角化
    t = compact_spaces(title)
    t = to_fullwidth(t) if t else ''
    return org_kana, org, d1, d2, t

def normalize_address_block(row: dict, label: str):
    """Address 1 - * を取り込み、住所1/2/3 + 郵便 整形
       - label が "Home"→自宅, "Other"→その他, それ以外→会社
    """
    addr_label = (row.get('Address 1 - Label','') or '').strip()
    if (label == 'home' and addr_label != 'Home') or \
       (label == 'other' and addr_label != 'Other') or \
       (label == 'work' and addr_label not in ('Work','','Company','Business')):
        return '', '', '', ''  # 該当なし

    street = row.get('Address 1 - Street','') or ''
    city   = row.get('Address 1 - City','') or ''
    region = row.get('Address 1 - Region','') or ''
    ext    = row.get('Address 1 - Extended Address','') or ''
    postal = row.get('Address 1 - Postal Code','') or ''

    # 住所1/2（全角化+建物分割）
    a1, a2 = compose_address(city=city, region=region, street=street, ext=ext)
    # 住所1/2とも全角に（念押し）
    a1 = zenkaku_numbers_and_ascii(a1)
    a2 = zenkaku_numbers_and_ascii(a2)
    # 住所3は空
    a3 = ''

    # 郵便：半角数字+半角ハイフン
    pz = normalize_postal(postal)

    return pz, a1, a2, a3

def collect_phones(row: dict):
    """Google フォーマットの Phone i - Label/Value を走査し、Work/Mobile は会社電話、Home は自宅電話、Other はその他電話へ。
       ※ご希望仕様に合わせ、Mobile も会社電話に寄せます。
    """
    home_list, work_list, other_list = [], [], []
    for i in range(1, 10):
        lbl = (row.get(f"Phone {i} - Label", '') or '').strip()
        val = row.get(f"Phone {i} - Value", '') or ''
        if not (lbl or val):
            continue
        numbers = normalize_phone_field(val)
        if not numbers:
            continue
        if lbl in ('Work', 'Mobile', 'Main', 'Company', 'Business'):
            work_list.extend(numbers)
        elif lbl == 'Home':
            home_list.extend(numbers)
        else:
            other_list.extend(numbers)
    # ユニーク & ; 結合（半角）
    home = ';'.join(uniq_preserve(home_list))
    work = ';'.join(uniq_preserve(work_list))
    other = ';'.join(uniq_preserve(other_list))
    return home, work, other

def collect_emails(row: dict):
    e1 = row.get('E-mail 1 - Value','') or ''
    e2 = row.get('E-mail 2 - Value','') or ''
    e3 = row.get('E-mail 3 - Value','') or ''
    return normalize_emails(e1, e2, e3)

def convert_row(row: dict) -> OrderedDict:
    # 氏名
    last = row.get('Last Name','') or ''
    first = row.get('First Name','') or ''
    last_k = row.get('Phonetic Last Name','') or ''
    first_k = row.get('Phonetic First Name','') or ''
    middle = row.get('Middle Name','') or ''
    middle_k = row.get('Phonetic Middle Name','') or ''
    nickname = row.get('Nickname','') or ''
    prefix = row.get('Name Prefix','') or ''
    suffix = row.get('Name Suffix','') or ''

    sei = last.strip()
    mei = first.strip()
    seikana = last_k.strip()
    meikana = first_k.strip()
    fullname = f"{sei}　{mei}".strip('　')
    fullname_k = f"{seikana}　{meikana}".strip('　')

    # 会社系
    org = row.get('Organization Name','') or ''
    title = row.get('Organization Title','') or ''
    dept = row.get('Organization Department','') or ''

    company_kana, company_name, dept1, dept2, yakushoku = normalize_company_fields(org, dept, title)

    # 住所：自宅/会社/その他
    home_pz, home_a1, home_a2, home_a3 = normalize_address_block(row, 'home')
    work_pz, work_a1, work_a2, work_a3 = normalize_address_block(row, 'work')
    other_pz, other_a1, other_a2, other_a3 = normalize_address_block(row, 'other')

    # 電話
    home_tel, work_tel, other_tel = collect_phones(row)

    # メール
    emails = collect_emails(row)

    # 誕生日
    birthday = normalize_birthday(row.get('Birthday','') or '')

    # メモ/備考
    memos, biko = collect_memos(row)

    # 出力行（理想型順に完全整列）
    out = OrderedDict()
    for col in OUTPUT_COLUMNS:
        out[col] = ''

    out["姓"] = sei
    out["名"] = mei
    out["姓かな"] = seikana
    out["名かな"] = meikana
    out["姓名"] = fullname
    out["姓名かな"] = fullname_k
    out["ミドルネーム"] = middle
    out["ミドルネームかな"] = middle_k
    out["敬称"] = "様" if True else ""
    out["ニックネーム"] = nickname
    out["旧姓"] = suffix  # 旧姓は Name Suffix が近い（必要なら調整）
    out["宛先"] = "会社"

    # 自宅
    out["自宅〒"] = home_pz
    out["自宅住所1"] = home_a1
    out["自宅住所2"] = home_a2
    out["自宅住所3"] = home_a3
    out["自宅電話"] = home_tel
    out["自宅IM ID"] = ""
    out["自宅E-mail"] = ""   # 自宅専用メールは分けない仕様
    out["自宅URL"] = ""
    out["自宅Social"] = ""

    # 会社
    out["会社〒"] = work_pz
    out["会社住所1"] = work_a1
    out["会社住所2"] = work_a2
    out["会社住所3"] = work_a3
    out["会社電話"] = work_tel
    out["会社IM ID"] = ""
    out["会社E-mail"] = emails
    out["会社URL"] = ""
    out["会社Social"] = ""

    # その他
    out["その他〒"] = other_pz
    out["その他住所1"] = other_a1
    out["その他住所2"] = other_a2
    out["その他住所3"] = other_a3
    out["その他電話"] = other_tel
    out["その他IM ID"] = ""
    out["その他E-mail"] = ""
    out["その他URL"] = ""
    out["その他Social"] = ""

    # 会社名等
    out["会社名かな"] = company_kana
    out["会社名"] = company_name
    out["部署名1"] = dept1
    out["部署名2"] = dept2
    out["役職名"] = yakushoku

    # 連名（今回は未使用）
    out["連名"] = ""
    out["連名ふりがな"] = ""
    out["連名敬称"] = ""
    out["連名誕生日"] = ""

    # メモ/備考
    out["メモ1"] = memos[0]
    out["メモ2"] = memos[1]
    out["メモ3"] = memos[2]
    out["メモ4"] = memos[3]
    out["メモ5"] = memos[4]
    out["備考1"] = biko[0]
    out["備考2"] = biko[1]
    out["備考3"] = biko[2]

    out["誕生日"] = birthday
    out["性別"] = ""
    out["血液型"] = ""
    out["趣味"] = ""
    out["性格"] = ""

    return out

# ==== CSV I/O ================================================================

def sniff_delimiter(sample: str) -> str:
    # タブ/カンマ自動判定
    if '\t' in sample and sample.count('\t') >= sample.count(','):
        return '\t'
    return ','

def read_contacts_text(text: str) -> list[dict]:
    # 1行目で区切り推定
    head = text.splitlines()[0] if text else ''
    delim = sniff_delimiter(head)
    f = io.StringIO(text)
    reader = csv.DictReader(f, delimiter=delim)
    return [r for r in reader]

def convert_google_to_atena(text: str) -> list[OrderedDict]:
    rows = read_contacts_text(text)
    return [convert_row(r) for r in rows]

def write_tsv(rows: list[OrderedDict]) -> str:
    output = io.StringIO()
    writer = csv.writer(output, delimiter='\t', lineterminator='\n')
    writer.writerow(OUTPUT_COLUMNS)
    for r in rows:
        writer.writerow([r.get(c, '') for c in OUTPUT_COLUMNS])
    return output.getvalue()

# ==== Web UI =================================================================

_HTML = """<!doctype html>
<meta charset="utf-8">
<title>google2atena v3.9.16 aligned no-pandas</title>
<style>
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,"Apple Color Emoji","Segoe UI Emoji";padding:24px;line-height:1.6}
input[type=file]{margin:12px 0}
button{padding:8px 16px;border:1px solid #999;border-radius:8px;background:#fafafa;cursor:pointer}
pre{white-space:pre-wrap;background:#f6f6f6;padding:12px;border-radius:8px}
.small{color:#666;font-size:12px}
</style>
<h1>google → atena 変換（v3.9.16 aligned / no-pandas）</h1>
<form method="POST" enctype="multipart/form-data">
  <input type="file" name="file" accept=".csv,.tsv,text/csv,text/tab-separated-values" required>
  <button type="submit">変換開始</button>
</form>
<p class="small">・Google 連絡先エクスポート（CSV/TSV）を読み込みます。<br>
・住所は自宅/会社/その他すべて、建物分割＋全角化。郵便は半角NNN-NNNN。<br>
・電話は全国市外局番＋特番（0120/0800/0570/050）に対応し <b>半角</b> で整形。<br>
・メールは複数区切りを統一し <code>;</code> 結合、重複排除。<br>
・誕生日は YYYY/MM/DD に統一。<br>
・会社名かなは全角、法人格削除、「・ . ，」除去、辞書＋規則で変換。</p>
"""

@app.get("/")
def index_get():
    return Response(_HTML, mimetype="text/html; charset=utf-8")

@app.post("/")
def index_post():
    file = request.files.get('file')
    if not file:
        return Response("ファイルが選択されていません。", status=400)
    text = file.stream.read().decode('utf-8-sig', errors='ignore')
    try:
        rows = convert_google_to_atena(text)
        tsv = write_tsv(rows)
        return Response(tsv, mimetype="text/tab-separated-values; charset=utf-8",
                        headers={"Content-Disposition": "attachment; filename=atena.tsv"})
    except Exception as e:
        return Response(f"Internal Server Error\n\n{html.escape(str(e))}", status=500, mimetype="text/plain; charset=utf-8")

# ==== main ===================================================================
if __name__ == "__main__":
    # ローカル実行用
    app.run(host="0.0.0.0", port=10000, debug=True)
