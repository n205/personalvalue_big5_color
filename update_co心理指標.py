# update_co個人価値観.py

import os
import warnings
import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter
from google.cloud import secretmanager
import google.generativeai as genai


# -------------------------------
# Secret Manager → Gemini APIキー
# -------------------------------
def load_gemini_key():
    client = secretmanager.SecretManagerServiceClient()
    secret_name = os.environ["GEMINI_API_KEY_SECRET"]
    response = client.access_secret_version(request={"name": secret_name})
    return response.payload.data.decode("utf-8")


# -------------------------------
# Geminiモデル生成（リーク防止）
# -------------------------------
def get_gemini_model():
    api_key = load_gemini_key()
    genai.configure(api_key=api_key)
    return genai.GenerativeModel("gemini-2.5-flash")   # ←あなた指定の2.5


# -------------------------------
# Schwartz PVQ
# -------------------------------
pvq_traits = [
    '自己方向性', '刺激', '享楽', '達成', '権力',
    '安全', '順応', '伝統', '博愛', '普遍主義'
]
pvq_columns = [f'PVQ_{t}' for t in pvq_traits]


# -------------------------------
# Gemini → PVQスコア推定
# -------------------------------
def extract_pvq_from_value(value_text):

    model = get_gemini_model()

    prompt = f"""
    あなたは心理学の専門家です。この文章を、Schwartzの10価値観（PVQ）に基づいて
    1〜7のスコアで評価してください。

    出力形式は以下に厳密に従ってください：

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
        res = model.generate_content(prompt)
        if not hasattr(res, "text"):
            return {}

        lines = res.text.strip().splitlines()
        scores = {}

        for line in lines:
            if ":" in line:
                key, val = line.split(":", 1)
                key = key.strip().replace(" ", "")
                if key in pvq_traits:
                    try:
                        scores[f"PVQ_{key}"] = int(val.strip())
                    except:
                        continue

        return scores

    except Exception as e:
        warnings.warn(f"Gemini PVQ推定エラー: {e}")
        return {}


# -------------------------------
# メイン処理（worksheetを引数で受け取る）
# -------------------------------
def update_co個人価値観(worksheet, df):

    print("=== 開始: 個人価値観(PVQ) 更新処理 ===")

    df = df.copy()
    df.fillna("", inplace=True)

    update_count = 0

    for idx, row in df.iterrows():

        company = row.get("会社名", "")
        value_text = row.get("バリュー", "")

        # -------- 対象外処理 --------
        if company == "対象外" or value_text in ["対象外", "取得失敗"]:

            # 既にすべて対象外ならスキップ
            if all(str(row.get(col, "")) == "対象外" for col in pvq_columns):
                continue

            print(f"⏭️ 対象外: {company}")

            for col in pvq_columns:
                df.at[idx, col] = "対象外"

            update_count += 1

            # 行ごとに即スプレッドシート更新
            for col in pvq_columns:
                col_idx = df.columns.get_loc(col)
                col_letter = get_column_letter(col_idx + 1)
                worksheet.update(
                    f"{col_letter}{idx+2}:{col_letter}{idx+2}",
                    [[df.at[idx, col]]]
                )

            continue

        # -------- 既に全て埋まっていればスキップ --------
        if all(str(row.get(col, "")) not in ["", "対象外"] for col in pvq_columns):
            continue

        # -------- Gemini 推定 --------
        scores = extract_pvq_from_value(value_text)

        if scores:
            print(f"✅ 推定成功: {company}")
            for col in pvq_columns:
                df.at[idx, col] = scores.get(col, "")
        else:
            print(f"⚠️ 推定失敗: {company}")
            continue

        update_count += 1

        # -------- 行ごとにシート更新（Cloud Run timeout対策）--------
        for col in pvq_columns:
            col_idx = df.columns.get_loc(col)
            col_letter = get_column_letter(col_idx + 1)
            worksheet.update(
                f"{col_letter}{idx+2}:{col_letter}{idx+2}",
                [[df.at[idx, col]]]
            )

    print(f"=== 完了: {update_count} 件更新 ===")

    return {"updated": update_count}

