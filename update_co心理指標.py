# update_co心理指標.py

import os
import warnings
import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter
from google.cloud import secretmanager

import google.generativeai as genai

from read_coデータ import read_coデータ


# -----------------------------------------
# Secret Manager → Gemini APIキー取得
# -----------------------------------------
def load_gemini_key():
    client = secretmanager.SecretManagerServiceClient()
    name = os.environ["GEMINI_API_KEY_SECRET"]
    response = client.access_secret_version(request={"name": name})
    return response.payload.data.decode("utf-8")


# -----------------------------------------
# Geminiモデルの初期化（都度生成：メモリリーク対策）
# -----------------------------------------
def get_gemini_model():
    api_key = load_gemini_key()
    genai.configure(api_key=api_key)
    return genai.GenerativeModel("gemini-2.5-flash")


# -----------------------------------------
# Schwartz PVQ 定義
# -----------------------------------------
pvq_traits = [
    '自己方向性', '刺激', '享楽', '達成', '権力',
    '安全', '順応', '伝統', '博愛', '普遍主義'
]
pvq_columns = [f'PVQ_{t}' for t in pvq_traits]


# -----------------------------------------
# Geminiで企業バリューからPVQ値を推定
# -----------------------------------------
def extract_pvq_from_value(value_text):
    model = get_gemini_model()

    prompt = f"""
    あなたは心理学の専門家です。
    以下の文章は、ある企業の「バリュー」または「行動指針」を要約したものです。

    Schwartzの10価値観（PVQ）理論に基づいて、この文章が各価値観をどの程度重視しているかを、1〜7の範囲で推定してください。
    「全く反映されていない」価値観 → 1
    「非常に強く反映されている」価値観 → 7

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


# -----------------------------------------
# Cloud Run Functions mainエントリポイント
# -----------------------------------------
def update_co心理指標(request):

    print("=== 開始: PVQスコア更新処理 ===")

    worksheet = read_coデータ()
    df = pd.DataFrame(worksheet.get_all_records())
    df.fillna("", inplace=True)

    update_count = 0

    for idx, row in df.iterrows():

        company = row.get("会社名", "")
        value_text = row.get("バリュー", "")

        # 対象外企業処理
        if company == "対象外" or value_text in ["対象外", "取得失敗"]:
            print(f"⏭️ 対象外スキップ: {company}")

            # 既に対象外なら更新不要
            if all(str(row.get(col, "")).strip() == "対象外" for col in pvq_columns):
                continue

            # 対象外で上書き
            for col in pvq_columns:
                df.at[idx, col] = "対象外"

            update_count += 1
            continue

        # すでに全て埋まっていればスキップ
        if all(str(row.get(col, "")).strip() not in ["", "対象外"] for col in pvq_columns):
            continue

        # PVQ推定
        scores = extract_pvq_from_value(value_text)

        if scores:
            for col in pvq_columns:
                df.at[idx, col] = scores.get(col, "")
            update_count += 1
            print(f"✅ 成功: {company}")
        else:
            print(f"⚠️ 推定失敗: {company}")

        # Cloud Run のタイムアウト回避のため、1行ごとに部分更新
        for col in pvq_columns:
            col_idx = df.columns.get_loc(col)
            col_letter = get_column_letter(col_idx + 1)
            worksheet.update(
                f"{col_letter}{idx+2}:{col_letter}{idx+2}",
                [[df.at[idx, col]]]
            )

    print(f"=== 完了: {update_count} 件更新 ===")

    return {"updated": update_count}
