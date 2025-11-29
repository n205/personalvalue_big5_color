import pandas as pd
import numpy as np
from gspread_dataframe import get_as_dataframe
from pdf2image import convert_from_bytes
from sklearn.cluster import KMeans
from PIL import Image
import requests
import warnings
import logging


# ============================================
# グレー判定
# ============================================
def is_near_gray(rgb, threshold=30):
    r, g, b = rgb
    return (
        abs(r - g) < threshold
        and abs(g - b) < threshold
        and abs(r - b) < threshold
    )


# ============================================
# PDF から主要色を抽出
# ============================================
def extract_main_colors_from_pdf(pdf_bytes, num_colors=2):
    try:
        # PDF → 1〜3ページ画像化
        images = convert_from_bytes(
            pdf_bytes,
            dpi=200,
            first_page=1,
            last_page=3
        )

        all_pixels = []

        for img in images:
            img_resized = img.resize((400, 400)).convert("RGB")
            arr = np.array(img_resized).reshape(-1, 3)

            # グレー・白黒付近を除去
            arr = np.array(
                [px for px in arr if not is_near_gray(px)],
                dtype=int
            )

            if len(arr) > 0:
                all_pixels.append(arr)

        if not all_pixels:
            return []

        full_array = np.vstack(all_pixels)

        # KMeans（主要2色）
        kmeans = KMeans(n_clusters=num_colors, random_state=0)
        kmeans.fit(full_array)

        centers = kmeans.cluster_centers_.astype(int)
        hex_colors = [f"#{r:02X}{g:02X}{b:02X}" for r, g, b in centers]

        return hex_colors

    except Exception as e:
        warnings.warn(f"色抽出失敗: {e}")
        return []


# ============================================
# update_色番号（メイン処理）
# ============================================
def update_色番号(worksheet):
    logging.info("🖼️ update_色番号 開始")

    df = get_as_dataframe(worksheet)
    df.fillna("", inplace=True)

    # 色列がなければ作成
    if "色1コード" not in df.columns:
        df["色1コード"] = ""
    if "色2コード" not in df.columns:
        df["色2コード"] = ""

    update_count = 0

    for idx, row in df.iterrows():
        url = row.get("URL", "")
        company = row.get("会社名", "")
        color1 = row.get("色1コード", "")
        color2 = row.get("色2コード", "")

        # URL が空、または両方埋まっている場合はスキップ（ログなし）
        if not url or (color1 and color2):
            continue

        # 対象外処理（ログなし）
        if company == "対象外":
            df.at[idx, "色1コード"] = "対象外"
            df.at[idx, "色2コード"] = "対象外"
            update_count += 1
            continue

        # PDF ダウンロード
        try:
            response = requests.get(
                url,
                headers={"User-Agent": "Mozilla/5.0"},
                timeout=15
            )

            if response.status_code == 200:
                colors = extract_main_colors_from_pdf(response.content)

                if len(colors) >= 2:
                    df.at[idx, "色1コード"] = colors[0]
                    df.at[idx, "色2コード"] = colors[1]
                    update_count += 1
                    logging.info(f"🎨 抽出成功: {url}")
                else:
                    df.at[idx, "色1コード"] = "取得失敗"
                    df.at[idx, "色2コード"] = "取得失敗"
                    update_count += 1
                    logging.warning(f"⚠️ 色抽出失敗: {url}")

            else:
                df.at[idx, "色1コード"] = "取得失敗"
                df.at[idx, "色2コード"] = "取得失敗"
                update_count += 1
                logging.warning(f"⚠️ ダウンロード失敗: {url}")

        except Exception as e:
            df.at[idx, "色1コード"] = "取得失敗"
            df.at[idx, "色2コード"] = "取得失敗"
            update_count += 1
            logging.warning(f"❌ エラー: {e} → {url}")

    # 欠損値補正
    df.replace([np.nan, np.inf, -np.inf], "", inplace=True)

    # A〜ZZ の列対応
    def col_to_letter(index):
        letters = ""
        while index >= 0:
            index, rem = divmod(index, 26)
            letters = chr(65 + rem) + letters
            index -= 1
        return letters

    # スプレッドシート更新
    for col in ["色1コード", "色2コード"]:
        col_index = df.columns.get_loc(col)
        col_letter = col_to_letter(col_index)

        worksheet.update(
            f"{col_letter}2:{col_letter}{len(df) + 1}",
            [[v] for v in df[col].tolist()]
        )

    logging.info(f"📝 {update_count} 件の色コードを更新しました")
    return f"{update_count} 件更新", 200
