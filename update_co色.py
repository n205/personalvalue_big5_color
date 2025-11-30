import pandas as pd
import numpy as np
from gspread_dataframe import get_as_dataframe
from pdf2image import convert_from_bytes
from sklearn.cluster import KMeans
from PIL import Image
import requests
import warnings
import logging
import re
from gspread_formatting import format_cell_ranges, CellFormat, Color


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
        images = convert_from_bytes(
            pdf_bytes,
            dpi=200,
            first_page=1,
            last_page=3,
        )

        all_pixels = []

        for img in images:
            img_resized = img.resize((400, 400)).convert("RGB")
            arr = np.array(img_resized).reshape(-1, 3)

            arr = np.array([px for px in arr if not is_near_gray(px)], dtype=int)

            if len(arr) > 0:
                all_pixels.append(arr)

        if not all_pixels:
            return []

        full_array = np.vstack(all_pixels)
        kmeans = KMeans(n_clusters=num_colors, random_state=0)
        kmeans.fit(full_array)

        centers = kmeans.cluster_centers_.astype(int)
        hex_colors = [f"#{r:02X}{g:02X}{b:02X}" for r, g, b in centers]

        return hex_colors

    except Exception as e:
        warnings.warn(f"色抽出失敗: {e}")
        return []


# ============================================
# 色番号を更新する（色コード抽出）
# ============================================
def update_co色番号(worksheet):
    logging.info("🖼️ update_co色番号 開始")

    df = get_as_dataframe(worksheet)
    df.fillna("", inplace=True)

    # 色列がなければ作成
    if "色1番号" not in df.columns:
        df["色1番号"] = ""
    if "色2番号" not in df.columns:
        df["色2番号"] = ""

    update_count = 0

    for idx, row in df.iterrows():
        url = row.get("URL", "")
        company = row.get("会社名", "")
        color1 = row.get("色1番号", "")
        color2 = row.get("色2番号", "")

        if not url or (color1 and color2):
            continue

        if company == "対象外":
            df.at[idx, "色1番号"] = "対象外"
            df.at[idx, "色2番号"] = "対象外"
            update_count += 1
            continue

        try:
            response = requests.get(
                url,
                headers={"User-Agent": "Mozilla/5.0"},
                timeout=15,
            )

            if response.status_code == 200:
                colors = extract_main_colors_from_pdf(response.content)

                if len(colors) >= 2:
                    df.at[idx, "色1番号"] = colors[0]
                    df.at[idx, "色2番号"] = colors[1]
                    update_count += 1
                    logging.info(f"🎨 抽出成功: {url}")
                else:
                    df.at[idx, "色1番号"] = "取得失敗"
                    df.at[idx, "色2番号"] = "取得失敗"
                    update_count += 1
                    logging.warning(f"⚠️ 色抽出失敗: {url}")

            else:
                df.at[idx, "色1番号"] = "取得失敗"
                df.at[idx, "色2番号"] = "取得失敗"
                update_count += 1
                logging.warning(f"⚠️ ダウンロード失敗: {url}")

        except Exception as e:
            df.at[idx, "色1番号"] = "取得失敗"
            df.at[idx, "色2番号"] = "取得失敗"
            update_count += 1
            logging.warning(f"❌ エラー: {e} → {url}")

    df.replace([np.nan, np.inf, -np.inf], "", inplace=True)

    # 列 index → A1 記法
    def col_to_letter(index):
        letters = ""
        while index >= 0:
            index, rem = divmod(index, 26)
            letters = chr(65 + rem) + letters
            index -= 1
        return letters

    for col in ["色1番号", "色2番号"]:
        col_index = df.columns.get_loc(col)
        col_letter = col_to_letter(col_index)

        worksheet.update(
            f"{col_letter}2:{col_letter}{len(df) + 1}",
            [[v] for v in df[col].tolist()],
        )

    logging.info(f"📝 {update_count} 件の色番号を更新しました")
    return f"{update_count} 件更新", 200


# ============================================
# HEX → 色塗りつぶし用 Color
# ============================================
def hex_to_color(hex_str):
    if (
        not isinstance(hex_str, str)
        or not re.match(r"^#([0-9A-Fa-f]{6})$", hex_str.strip())
    ):
        return None

    r = int(hex_str[1:3], 16) / 255
    g = int(hex_str[3:5], 16) / 255
    b = int(hex_str[5:7], 16) / 255

    return Color(red=r, green=g, blue=b)


# ============================================
# 色番号に応じてセルを塗りつぶす
# ============================================
def update_co色(worksheet):
    logging.info("🎨 update_co色（塗りつぶし）開始")

    df = get_as_dataframe(worksheet)
    df.fillna("", inplace=True)

    start_row = 2

    color_map = {
        "色1番号": "色1",
        "色2番号": "色2",
    }

    # 列 index → A1 記法
    def col_to_letter(n):
        result = ""
        while n >= 0:
            result = chr(n % 26 + ord("A")) + result
            n = n // 26 - 1
        return result

    for code_col, fill_col in color_map.items():
        if code_col not in df.columns or fill_col not in df.columns:
            logging.warning(f"⚠️ 列が存在しません: {code_col} / {fill_col}")
            continue

        fill_index = df.columns.get_loc(fill_col)
        col_letter = col_to_letter(fill_index)

        format_list = []

        for i, hex_code in enumerate(df[code_col]):
            color = hex_to_color(hex_code)
            if color is None:
                continue  # 無効色の場合は塗りつぶしなし

            row_num = start_row + i
            cell_ref = f"{col_letter}{row_num}"

            format_list.append((cell_ref, CellFormat(backgroundColor=color)))

        if format_list:
            format_cell_ranges(worksheet, format_list)
            logging.info(f"🟩 {fill_col}: {len(format_list)} 件に塗りつぶし適用")
        else:
            logging.info(f"ℹ️ {fill_col}: 有効なカラーコードなし")

    return "OK", 200
