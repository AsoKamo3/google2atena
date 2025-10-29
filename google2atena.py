# -*- coding: utf-8 -*-
"""
google2atena.py v3.9.14
----------------------------------------
Googleスプレッドシートなどから抽出した企業情報を整理・整形し、
宛名帳（CSV/Excel形式）を生成するスクリプト。

対応機能：
- 法人格（株式会社など）自動除去
- 住所と建物の分割
- 電話番号の正規化（ハイフン整形）
- 「会社かな」生成（全角カタカナ化、句読点除去）
- 外部辞書：COMPANY_EXCEPT / KANJI_WORD_MAP 対応
----------------------------------------
"""

import re
import unicodedata
import pandas as pd
from company_dicts import COMPANY_EXCEPT
from kanji_word_map import KANJI_WORD_MAP


# ----------------------------------------
# 共通設定
# ----------------------------------------
CORP_TERMS = [
    "株式会社", "有限会社", "合同会社", "一般社団法人", "一般財団法人",
    "公益社団法人", "公益財団法人", "特定非営利活動法人",
    "学校法人", "医療法人", "宗教法人", "社会福祉法人",
    "公立大学法人", "独立行政法人", "地方独立行政法人"
]

# ----------------------------------------
# Utility functions
# ----------------------------------------

def remove_corp_terms(name: str) -> str:
    """法人格の除去"""
    if not isinstance(name, str):
        return ""
    for term in CORP_TERMS:
        name = re.sub(term, "", name)
    return name.strip()


def normalize_phone(phone: str) -> str:
    """電話番号をハイフン区切りに整形"""
    if not isinstance(phone, str):
        return ""
    digits = re.sub(r"\D", "", phone)
    # 日本の市外局番・携帯番号・フリーダイヤルなどを考慮
    if digits.startswith("0120"):
        return f"{digits[:4]}-{digits[4:7]}-{digits[7:]}"
    elif digits.startswith("080") or digits.startswith("090") or digits.startswith("070"):
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    elif len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    elif len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    else:
        return phone


def split_address(address: str):
    """住所を「住所」と「建物」に分割"""
    if not isinstance(address, str):
        return "", ""
    pattern = r"(.*?)(?:　| )(.*[ビル|マンション|ハイツ|号|室|階|棟].*)"
    match = re.match(pattern, address)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return address.strip(), ""


def to_fullwidth_katakana(text: str) -> str:
    """すべて全角カタカナに統一"""
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[・.,，]", "", text)
    text = text.translate(str.maketrans({
        "ﾞ": "゛", "ﾟ": "゜",
    }))
    return text


# ----------------------------------------
# 会社名かな変換
# ----------------------------------------

def company_to_kana(name: str) -> str:
    """会社名からカナ変換を生成"""
    if not isinstance(name, str) or name.strip() == "":
        return ""

    # ① 辞書に完全一致している場合
    if name in COMPANY_EXCEPT:
        return to_fullwidth_katakana(COMPANY_EXCEPT[name])

    # ② 法人格を除去して検索
    name_clean = remove_corp_terms(name)
    if name_clean in COMPANY_EXCEPT:
        return to_fullwidth_katakana(COMPANY_EXCEPT[name_clean])

    # ③ 後方一致（KANJI_WORD_MAP）
    for key, val in KANJI_WORD_MAP.items():
        if name_clean.endswith(key):
            prefix = name_clean[: -len(key)]
            kana_prefix = company_to_kana(prefix) if prefix else ""
            return to_fullwidth_katakana(kana_prefix + val)

    # ④ 自動カタカナ化（英字・数字なども含む）
    name_kana = re.sub(r"[A-Za-z]", lambda m: chr(ord(m.group(0).upper()) + 0xFEE0), name_clean)
    return to_fullwidth_katakana(name_kana)


# ----------------------------------------
# DataFrame 変換
# ----------------------------------------

def convert_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Googleスプレッドシート形式のデータを宛名帳用に整形"""
    df = df.copy()

    df["会社名"] = df["会社名"].fillna("").astype(str).str.strip()
    df["会社かな"] = df["会社名"].apply(company_to_kana)

    df["住所"], df["建物"] = zip(*df["住所"].map(split_address))
    df["電話番号"] = df["電話番号"].apply(normalize_phone)

    return df


# ----------------------------------------
# メイン処理
# ----------------------------------------

def main(input_path: str, output_path: str):
    df = pd.read_csv(input_path)
    df_converted = convert_dataframe(df)
    df_converted.to_excel(output_path, index=False)
    print(f"✅ 出力完了: {output_path}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("使い方: python google2atena.py input.csv output.xlsx")
    else:
        main(sys.argv[1], sys.argv[2])
