import logging
import warnings
import numpy as np
import google.generativeai as genai
from gspread_dataframe import get_as_dataframe
import os


# ============================================
# Gemini 初期化
# ============================================
def init_gemini():
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("環境変数 GEMINI_API_KEY が設定されていません")
    genai.configure(api_key=api_key)
    return genai.GenerativeModel("gemini-3.5-flash")


gemini_model = None


# ============================================
# PVQ 10項目
# ============================================
pvq_traits = [
    "自己方向性", "刺激", "享楽", "達成", "権力",
    "安全", "順応", "伝統", "博愛", "普遍主義"
]
pvq_columns = [f"PVQ_{t}" for t in pvq_traits]


# ============================================
# Gemini による PVQ 推定
# ============================================
def extract_pvq_scores(value_text):
    global gemini_model
    if gemini_model is None:
        gemini_model = init_gemini()

    prompt = f"""
        あなたは心理学の専門家です。
        以下の文章は、ある企業の「バリュー」または「行動指針」を要約したものです。
        
        Schwartzの10価値観（PVQ）理論に基づいて、この文章が各価値観をどの程度重視しているかを、1〜7で推定してください。
        
        出力形式（順番厳守）：
        自己方向性: 数値
        刺激: 数値
        享楽: 数値
        達成: 数値
        権力: 数値
        安全: 数値
        順応: 数値
        伝統: 数値
        博愛: 数値
        普遍主義: 数値
        
        ---
        {value_text}
        """

    try:
        res = gemini_model.generate_content(prompt)
        text = res.text.strip()
        lines = text.splitlines()

        scores = {}
        for line in lines:
            if ":" not in line:
                continue
            key, val = line.split(":", 1)
            key = key.strip()
            val = val.strip()

            if key in pvq_traits:
                try:
                    scores[f"PVQ_{key}"] = int(val)
                except:
                    scores[f"PVQ_{key}"] = ""

        return scores

    except Exception as e:
        warnings.warn(f"Gemini PVQ推定エラー: {e}")
        return {}


# ============================================
# update_co個人価値観（メイン処理）
# ============================================
def update_co個人価値観(worksheet):
    logging.info("🧭 update_co個人価値観 開始")

    df = get_as_dataframe(worksheet)
    df.fillna("", inplace=True)

    # PVQ列がなければ作成
    for col in pvq_columns:
        if col not in df.columns:
            df[col] = ""

    update_count = 0

    for idx, row in df.iterrows():
        company = row.get("会社名", "")
        value_text = row.get("バリュー", "")

        # すでに埋まっている行はスキップ（ログは出さない）
        if all(str(row.get(col, "")).strip() not in ["", "対象外"] for col in pvq_columns):
            continue

        # 「対象外」処理（ログを出さない）
        if company == "対象外" or value_text in ["対象外", "取得失敗", ""]:
            for col in pvq_columns:
                df.at[idx, col] = "対象外"
            update_count += 1
            continue

        # ---------- PVQ 推定 ----------
        scores = extract_pvq_scores(value_text)

        if scores and any(scores.values()):
            # 正常にスコアが返った場合
            for col in pvq_columns:
                df.at[idx, col] = scores.get(col, "")
            update_count += 1
            logging.info(f"📝 PVQ推定: {company}")
        else:
            # Gemini の推定が失敗した場合 → すべて「対象外」
            for col in pvq_columns:
                df.at[idx, col] = "対象外"
            update_count += 1
            logging.warning(f"⚠️ 推定失敗 → 対象外に設定: {company}")

    df.replace([np.nan, np.inf, -np.inf], "", inplace=True)

    # ============================================
    # 列全体一括更新
    # ============================================
    def col_to_letter(index):
        letters = ""
        while index >= 0:
            index, rem = divmod(index, 26)
            letters = chr(65 + rem) + letters
            index -= 1
        return letters

    for col in pvq_columns:
        col_index = df.columns.get_loc(col)
        col_letter = col_to_letter(col_index)

        worksheet.update(
            f"{col_letter}2:{col_letter}{len(df)+1}",
            [[v] for v in df[col].tolist()]
        )

    logging.info(f"📝 {update_count} 件のPVQスコアを更新しました")
    return f"{update_count} 件更新", 200


# ============================================
# Big Five 推定
# ============================================
big5_traits = [
    "Extraversion",
    "Agreeableness",
    "Conscientiousness",
    "Neuroticism",
    "Openness"
]

big5_columns = big5_traits[:]   # そのまま列名に使う


def extract_big_five_from_value(value_text):
    """バリュー文からBig Fiveを推定（2〜14の整数）"""
    global gemini_model
    if gemini_model is None:
        gemini_model = init_gemini()

    prompt = f"""
        あなたは心理学者です。
        以下は企業文化を示す「バリュー」または「行動指針」の要約です。
        
        これを書いた人物の Big Five（性格5因子）を推定し、
        各因子を 2〜14 の整数値で返してください。
        
        出力形式（順番厳守）：
        Extraversion: 数値
        Agreeableness: 数値
        Conscientiousness: 数値
        Neuroticism: 数値
        Openness: 数値
        
        ---
        {value_text}
        """

    try:
        res = gemini_model.generate_content(prompt)
        text = res.text.strip()
        lines = text.splitlines()

        scores = {}
        for line in lines:
            if ":" not in line:
                continue
            key, val = line.split(":", 1)
            key = key.strip()
            val = val.strip()

            if key in big5_traits:
                try:
                    scores[key] = int(val)
                except:
                    scores[key] = ""

        return scores

    except Exception as e:
        warnings.warn(f"Gemini Big5推定エラー: {e}")
        return {}


# ============================================================
# update_cobig5（メイン処理）
# ============================================================
def update_cobig5(worksheet):
    logging.info("🧭 update_cobig5 開始")

    df = get_as_dataframe(worksheet)
    df.fillna("", inplace=True)

    # Big5列がなければ作成
    for col in big5_columns:
        if col not in df.columns:
            df[col] = ""

    update_count = 0

    for idx, row in df.iterrows():
        company = row.get("会社名", "")
        value_text = row.get("バリュー", "")

        # 既に全項目が埋まっている行はスキップ
        if all(str(row.get(col, "")).strip() not in ["", "対象外"] for col in big5_columns):
            continue

        # 対象外の処理（ログ出さない）
        if company == "対象外" or value_text in ["対象外", "取得失敗", ""]:
            for col in big5_columns:
                df.at[idx, col] = "対象外"
            update_count += 1
            continue

        # Gemini 推定
        scores = extract_big_five_from_value(value_text)

        if scores and any(scores.values()):
            for col in big5_columns:
                df.at[idx, col] = scores.get(col, "")
            update_count += 1
            logging.info(f"📝 Big5推定: {company}")

        else:
            # 推定失敗 → 全て対象外
            for col in big5_columns:
                df.at[idx, col] = "対象外"
            update_count += 1
            logging.warning(f"⚠️ Big5推定失敗 → 対象外に設定: {company}")

    # 欠損・無限を除去
    df.replace([np.nan, np.inf, -np.inf], "", inplace=True)

    # Excel一括更新（PVQと同じ関数を利用）
    def col_to_letter(index):
        letters = ""
        while index >= 0:
            index, rem = divmod(index, 26)
            letters = chr(65 + rem) + letters
            index -= 1
        return letters

    for col in big5_columns:
        col_index = df.columns.get_loc(col)
        col_letter = col_to_letter(col_index)

        worksheet.update(
            f"{col_letter}2:{col_letter}{len(df) + 1}",
            [[v] for v in df[col].tolist()]
        )

    logging.info(f"📝 {update_count} 件のBig Fiveスコアを更新しました")
    return f"{update_count} 件更新", 200

