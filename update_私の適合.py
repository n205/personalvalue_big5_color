# update_私の適合.py (色列名を '色1番号' / '色2番号' に修正した完全版)

import pandas as pd
import numpy as np
import logging
import re
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from matplotlib.colors import to_rgb
from sklearn.preprocessing import MinMaxScaler
from gspread_formatting import format_cell_ranges, CellFormat, Color


# HEX -> gspread_formatting.Color
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


# 列 index -> A1 記法
def col_to_letter(index):
    letters = ""
    while index >= 0:
        index, rem = divmod(index, 26)
        letters = chr(65 + rem) + letters
        index -= 1
    return letters


def update_私の適合(worksheet, target_ws):
    logging.info("🔍 update_私の適合 開始")

    # ---- ユーザー設定
    my_bigfive = {
        "Extraversion": 3,
        "Agreeableness": 9,
        "Conscientiousness": 12,
        "Neuroticism": 6,
        "Openness": 8,
    }

    my_pvq = {
        "PVQ_自己方向性": 7,
        "PVQ_刺激": 2,
        "PVQ_享楽": 2,
        "PVQ_達成": 4,
        "PVQ_権力": 1,
        "PVQ_安全": 7,
        "PVQ_順応": 6,
        "PVQ_伝統": 1,
        "PVQ_博愛": 2,
        "PVQ_普遍主義": 3,
    }

    favorite_color = "#006400"
    unfavorite_color = "#ff0000"

    bigfive_traits = list(my_bigfive.keys())
    pvq_traits = list(my_pvq.keys())

    my_bigfive_vec = np.array([my_bigfive[t] for t in bigfive_traits])
    my_pvq_vec = np.array([my_pvq[t] for t in pvq_traits])

    # ---- データ読み込み
    df = get_as_dataframe(worksheet)
    df.fillna("", inplace=True)

    # 数値変換（存在しない/数値でないものは NaN になる）
    for col in bigfive_traits + pvq_traits:
        df[col] = pd.to_numeric(df.get(col, pd.Series(dtype=float)), errors="coerce")

    # ---- 有効行フィルタ
    # NOTE: 色列は '色1番号' / '色2番号' を使用する仕様に合わせる
    valid_rows = df[
        (df.get("会社名", "") != "")
        & (df.get("会社名", "") != "対象外")
        & (df.get("バリュー", "") != "")
        & df[bigfive_traits + pvq_traits].notnull().all(axis=1)
        & (df.get("色1番号", "") != "")
        & (df.get("色2番号", "") != "")
    ].copy()

    if len(valid_rows) == 0:
        logging.warning("⚠️ 有効な行がありません")
        return "No valid rows", 200

    # ---- スコア計算
    def compute_bigfive_score(row):
        vec = np.array([row[t] for t in bigfive_traits], dtype=float)
        return 1 / (1 + np.linalg.norm(my_bigfive_vec - vec))

    def compute_pvq_score(row):
        vec = np.array([row[t] for t in pvq_traits], dtype=float)
        return 1 / (1 + np.linalg.norm(my_pvq_vec - vec))

    def compute_color_score(row):
        try:
            c1 = np.array(to_rgb(row["色1番号"]))
            c2 = np.array(to_rgb(row["色2番号"]))
            fav = np.array(to_rgb(favorite_color))
            unfav = np.array(to_rgb(unfavorite_color))

            sim_fav = max(1 - np.linalg.norm(c1 - fav), 1 - np.linalg.norm(c2 - fav))
            sim_unfav = max(1 - np.linalg.norm(c1 - unfav), 1 - np.linalg.norm(c2 - unfav))

            return sim_fav - sim_unfav
        except Exception:
            return 0

    valid_rows["B5相性スコア_そのまま"] = valid_rows.apply(compute_bigfive_score, axis=1)
    valid_rows["PVQ相性スコア_そのまま"] = valid_rows.apply(compute_pvq_score, axis=1)
    valid_rows["色相性スコア_そのまま"] = valid_rows.apply(compute_color_score, axis=1)

    # ---- 正規化
    scaler = MinMaxScaler()
    valid_rows[["B5相性スコア_01", "PVQ相性スコア_01", "色相性スコア_01"]] = scaler.fit_transform(
        valid_rows[["B5相性スコア_そのまま", "PVQ相性スコア_そのまま", "色相性スコア_そのまま"]]
    )

    # ---- 順位
    valid_rows["B5相性スコア_順位"] = valid_rows["B5相性スコア_そのまま"].rank(ascending=False)
    valid_rows["PVQ相性スコア_順位"] = valid_rows["PVQ相性スコア_そのまま"].rank(ascending=False)
    valid_rows["色相性スコア_順位"] = valid_rows["色相性スコア_そのまま"].rank(ascending=False)

    # ---- 総合スコア（重み付け）
    valid_rows["総合スコア"] = (
        valid_rows["B5相性スコア_01"] * 0.35
        + valid_rows["PVQ相性スコア_01"] * 0.45
        + valid_rows["色相性スコア_01"] * 0.20
    )

    # ---- 出力整形（色列は '色1番号' / '色2番号' を表示）
    result_df = valid_rows.sort_values("総合スコア", ascending=False)[
        [
            "会社名",
            "色1",
            "色2",
            "総合スコア",
            "バリュー",
            "URL",
            "B5相性スコア_そのまま",
            "B5相性スコア_01",
            "B5相性スコア_順位",
            "PVQ相性スコア_そのまま",
            "PVQ相性スコア_01",
            "PVQ相性スコア_順位",
            "色相性スコア_そのまま",
            "色相性スコア_01",
            "色相性スコア_順位",
            "色1番号",
            "色2番号",
        ]
    ]

    # ---- スプレッドシート出力
    target_ws.clear()
    set_with_dataframe(target_ws, result_df)

    # ---- 色塗り（出力シートの '色1' / '色2' を塗る）
    df_out = get_as_dataframe(target_ws)
    df_out.fillna("", inplace=True)

    color_map = {
        "色1番号": "色1",
        "色2番号": "色2",
    }

    start_row = 2

    for code_col, fill_col in color_map.items():
        if code_col not in df_out.columns or fill_col not in df_out.columns:
            logging.warning(f"⚠️ 列が見つかりません: {code_col} / {fill_col}")
            continue

        fill_idx = df_out.columns.get_loc(fill_col)
        col_letter = col_to_letter(fill_idx)

        ranges = []
        for i, hex_code in enumerate(df_out[code_col]):
            color = hex_to_color(hex_code)
            if color:
                cell_range = f"{col_letter}{start_row + i}"
                ranges.append((cell_range, CellFormat(backgroundColor=color)))

        if ranges:
            format_cell_ranges(target_ws, ranges)
            logging.info(f"🎨 {fill_col}: {len(ranges)} 件 塗りつぶし適用")

    msg = f"✅ 相性スコア {len(result_df)} 件更新"
    logging.info(msg)
    return msg, 200

