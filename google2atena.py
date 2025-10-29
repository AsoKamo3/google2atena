# -*- coding: utf-8 -*-
"""
google2atena_nopandas.py v3.9.14-noPandas
----------------------------------------
pandasを使わずに、CSV→CSV変換を行う軽量版。
Renderなど軽量環境対応。外部辞書を使用。
----------------------------------------
"""

import csv
import re
import unicodedata
from company_dicts import COMPANY_EXCEPT
from kanji_word_map import KANJI_WORD_MAP

CORP_TERMS = [
    "株式会社", "有限会社", "合同会社", "合資会社", "相互会社", "一般社団法人", "一般財団法人",
    "公益社団法人", "公益財団法人", "特定非営利活動法人", "ＮＰＯ法人",
    "学校法人", "医療法人", "宗教法人", "社会福祉法人",
    "公立大学法人", "独立行政法人", "地方独立行政法人"
]


def remove_corp_terms(name: str) -> str:
    if not isinstance(name, str):
        return ""
    for term in CORP_TERMS:
        name = re.sub(term, "", name)
    return name.strip()


def normalize_phone(phone: str) -> str:
    if not isinstance(phone, str):
        return ""
    digits = re.sub(r"\D", "", phone)
    if digits.startswith("0120"):
        return f"{digits[:4]}-{digits[4:7]}-{digits[7:]}"
    elif digits.startswith(("070", "080", "090")):
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    elif len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    elif len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    else:
        return phone


def split_address(address: str):
    if not isinstance(address, str):
        return "", ""
    pattern = r"(.*?)(?:　| )(.*[ビル|マンション|ハイツ|号|室|階|棟].*)"
    match = re.match(pattern, address)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return address.strip(), ""


def to_fullwidth_katakana(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[・.,，]", "", text)
    return text


def company_to_kana(name: str) -> str:
    if not isinstance(name, str) or name.strip() == "":
        return ""

    if name in COMPANY_EXCEPT:
        return to_fullwidth_katakana(COMPANY_EXCEPT[name])

    name_clean = remove_corp_terms(name)
    if name_clean in COMPANY_EXCEPT:
        return to_fullwidth_katakana(COMPANY_EXCEPT[name_clean])

    for key, val in KANJI_WORD_MAP.items():
        if name_clean.endswith(key):
            prefix = name_clean[: -len(key)]
            kana_prefix = company_to_kana(prefix) if prefix else ""
            return to_fullwidth_katakana(kana_prefix + val)

    name_kana = re.sub(r"[A-Za-z]", lambda m: chr(ord(m.group(0).upper()) + 0xFEE0), name_clean)
    return to_fullwidth_katakana(name_kana)


def convert_csv(input_path: str, output_path: str):
    with open(input_path, newline='', encoding='utf-8') as infile, \
         open(output_path, 'w', newline='', encoding='utf-8') as outfile:

        reader = csv.DictReader(infile)
        fieldnames = reader.fieldnames + ["会社かな", "住所1", "住所2", "電話番号整形"]
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()

        for row in reader:
            name = row.get("会社名", "")
            addr = row.get("住所", "")
            phone = row.get("電話番号", "")

            kana = company_to_kana(name)
            addr1, addr2 = split_address(addr)
            phone_fmt = normalize_phone(phone)

            row.update({
                "会社かな": kana,
                "住所1": addr1,
                "住所2": addr2,
                "電話番号整形": phone_fmt
            })
            writer.writerow(row)

    print(f"✅ 出力完了: {output_path}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("使い方: python google2atena_nopandas.py input.csv output.csv")
    else:
        convert_csv(sys.argv[1], sys.argv[2])
