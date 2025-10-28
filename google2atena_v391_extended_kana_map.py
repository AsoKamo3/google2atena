# -*- coding: utf-8 -*-
"""
Google連絡先CSV → 宛名職人CSV 変換（v3.9.1 extended kana map）
・会社名かな：包括的音訳 + 全角カタカナ統一 + 法人格除去
・電話番号：Work > Mobile > Home 優先、半角ハイフン整形
・住所：スペースで分割（全角化）
・メモ1〜5保持
"""

import pandas as pd
import re
import unicodedata
from flask import Flask, request, send_file, render_template_string
from io import BytesIO

app = Flask(__name__)

# ===============================
# 全角カタカナ変換ヘルパ
# ===============================
def to_fullwidth_katakana(text):
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = text.translate(str.maketrans({chr(i): chr(i + 0x60) for i in range(ord("ぁ"), ord("ゖ"))}))
    text = "".join(unicodedata.normalize("NFKC", c) for c in text)
    return text

# ===============================
# 会社名 → カタカナ変換ロジック
# ===============================
def company_to_kana(name: str) -> str:
    if not name:
        return ""

    name = unicodedata.normalize("NFKC", name)
    name = re.sub(r"(株式会社|有限会社|合同会社|一般社団法人|公益社団法人|公益財団法人|財団法人|社団法人)", "", name)
    name = re.sub(r"[・　\s]", "", name)
    name_upper = name.upper()

    # --- 包括的音訳辞書 ---
    romaji_map = {
        # 基本構成語
        "MAIN": "メイン", "BASE": "ベース", "CORE": "コア", "SYSTEM": "システム",
        "DIGITAL": "デジタル", "CREATIVE": "クリエイティブ", "GLOBAL": "グローバル",
        "FUTURE": "フューチャー", "NEXT": "ネクスト",

        # 業種・一般名詞
        "PRESS": "プレス", "BOOK": "ブック", "BOOKS": "ブックス", "STUDIO": "スタジオ",
        "DESIGN": "デザイン", "WORKS": "ワークス", "LAB": "ラボ", "LABO": "ラボ",
        "PROJECT": "プロジェクト", "PRODUCTION": "プロダクション",
        "COMMUNICATIONS": "コミュニケーションズ", "MEDIA": "メディア",
        "INC": "インク", "LTD": "リミテッド", "TECH": "テック",
        "ENGINEERING": "エンジニアリング", "CONSULTING": "コンサルティング",
        "SERVICE": "サービス", "GROUP": "グループ", "HOLDINGS": "ホールディングス",
        "FOUNDATION": "ファウンデーション", "ASSOCIATION": "アソシエーション",
        "COMPANY": "カンパニー", "ENTERPRISE": "エンタープライズ",
        "PARTNERS": "パートナーズ", "WORKSHOP": "ワークショップ",

        # 固有略語・特殊単語
        "PHP": "ピーエイチピー", "AI": "エーアイ", "ART": "アート",
        "CENTER": "センター", "INSTITUTE": "インスティテュート",
        "UNION": "ユニオン", "BANK": "バンク", "SOCIETY": "ソサエティ",
        "JAPAN": "ジャパン",

        # 新規追加語（ユーザー指定）
        "OFFICE": "オフィス", "NHK": "エヌエイチケー", "KADOKAWA": "カドカワ",
        "STAND": "スタンド", "NEO": "ネオ", "REAL": "リアル", "MARUZEN": "マルゼン",
        "YADOKARI": "ヤドカリ", "TOI": "トイ", "PLAN": "プラン",
        "ALL": "オール", "REVIEWS": "レビューズ", "COUNTER": "カウンター", "ODD": "オッド",
    }

    for eng, kana in romaji_map.items():
        if eng in name_upper:
            name_upper = name_upper.replace(eng, kana)

    # --- 単文字変換 ---
    letter_map = {
        "A": "エー","B": "ビー","C": "シー","D": "ディー","E": "イー",
        "F": "エフ","G": "ジー","H": "エイチ","I": "アイ","J": "ジェイ",
        "K": "ケー","L": "エル","M": "エム","N": "エヌ","O": "オー",
        "P": "ピー","Q": "キュー","R": "アール","S": "エス","T": "ティー",
        "U": "ユー","V": "ブイ","W": "ダブリュー","X": "エックス",
        "Y": "ワイ","Z": "ゼット"
    }
    name_upper = "".join(letter_map.get(ch, ch) for ch in name_upper)
    return to_fullwidth_katakana(name_upper)

# ===============================
# 住所分割（スペース基準）
# ===============================
def split_address(addr):
    if not isinstance(addr, str) or addr.strip() == "":
        return "", ""
    parts = addr.strip().split(" ", 1)
    if len(parts) == 1:
        return to_fullwidth_katakana(parts[0]), ""
    return to_fullwidth_katakana(parts[0]), to_fullwidth_katakana(parts[1])

# ===============================
# 電話番号整形＋優先順位ロジック
# ===============================
def format_phone(phone):
    if not isinstance(phone, str):
        return ""
    phone = re.sub(r"[^\d]", "", phone)
    if len(phone) < 7:
        return ""
    if len(phone) == 10:
        return f"{phone[0:2]}-{phone[2:6]}-{phone[6:]}"
    if len(phone) == 11:
        return f"{phone[0:3]}-{phone[3:7]}-{phone[7:]}"
    return phone

def merge_phones(row):
    phones = []
    for label in ["Work", "Mobile", "Home"]:
        for i in range(1, 6):
            value = row.get(f"Phone {i} - Value", "")
            lbl = row.get(f"Phone {i} - Label", "")
            if label.lower() in lbl.lower():
                f = format_phone(value)
                if f:
                    phones.append((label, f))
    phones.sort(key=lambda x: {"Work": 0, "Mobile": 1, "Home": 2}.get(x[0], 99))
    unique = []
    for _, num in phones:
        if num not in unique:
            unique.append(num)
    return ";".join(unique)

# ===============================
# メイン変換処理
# ===============================
def convert_csv(input_csv):
    df = pd.read_csv(input_csv, sep="\t")
    rows = []

    for _, row in df.iterrows():
        org_name = str(row.get("Organization Name", "")).strip()
        org_kana = company_to_kana(org_name)

        addr1, addr2 = split_address(str(row.get("Address 1 - Street", "")))

        phones = merge_phones(row)
        emails = ";".join(str(row.get(col, "")).strip() for col in df.columns if "E-mail" in col and "Value" in col and pd.notna(row.get(col)))

        rows.append({
            "姓": row.get("Last Name", ""),
            "名": row.get("First Name", ""),
            "姓かな": row.get("Phonetic Last Name", ""),
            "名かな": row.get("Phonetic First Name", ""),
            "姓名": f"{row.get('Last Name','')}　{row.get('First Name','')}",
            "姓名かな": f"{row.get('Phonetic Last Name','')}　{row.get('Phonetic First Name','')}",
            "敬称": "様",
            "宛先": "会社",
            "会社〒": row.get("Address 1 - Postal Code", ""),
            "会社住所1": addr1,
            "会社住所2": addr2,
            "会社電話1〜10": phones,
            "会社E-mail1〜5": emails,
            "会社名かな": org_kana,
            "会社名": org_name,
            "部署名1": row.get("Organization Department", ""),
            "役職名": row.get("Organization Title", ""),
            "メモ1": row.get("メモ1", ""),
            "メモ2": row.get("メモ2", ""),
            "メモ3": row.get("メモ3", ""),
            "メモ4": row.get("メモ4", ""),
            "メモ5": row.get("メモ5", ""),
        })

    out_df = pd.DataFrame(rows)
    buffer = BytesIO()
    out_df.to_csv(buffer, index=False, encoding="utf-8-sig")
    buffer.seek(0)
    return buffer

# ===============================
# Flaskルート
# ===============================
@app.route("/")
def index():
    return render_template_string("""
    <h2>Google連絡先CSV → 宛名職人CSV 変換（v3.9.1 extended kana map）</h2>
    <form action="/convert" method="post" enctype="multipart/form-data">
        <p><input type="file" name="file" accept=".csv" required></p>
        <p><input type="submit" value="変換する"></p>
    </form>
    """)

@app.route("/convert", methods=["POST"])
def convert():
    file = request.files["file"]
    buffer = convert_csv(file)
    return send_file(buffer, as_attachment=True, download_name="converted_v391.csv", mimetype="text/csv")

if __name__ == "__main__":
    app.run(debug=True)
