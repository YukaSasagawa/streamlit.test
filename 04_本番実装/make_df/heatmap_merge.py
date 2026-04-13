import pandas as pd
from .heatmap_VR import make_df_heatmap_VR
from .heatmap_TVAL import make_df_heatmap_TVAL
from .heatmap_REVISIO import make_df_heatmap_REVISIO

def make_df_heatmap_all(
    vr_all_file=None,
    vr_target_file=None,
    tval_all_file=None,
    tval_target_file=None,
    revisio_file=None
):
    """
    VR / TVAL / REVISIO を結合する関数（個人全体・ターゲット対応）

    Returns
    -------
    pandas.DataFrame
        時間帯, 月〜日, 局, データ種別, カテゴリー
    """

    df_list = []

    # =========================
    # VR
    # =========================
    if vr_all_file is not None:
        df = make_df_heatmap_VR(vr_all_file).copy()
        df["データ種別"] = "VR"
        df["カテゴリー"] = "個人全体"
        df_list.append(df)

    if vr_target_file is not None:
        df = make_df_heatmap_VR(vr_target_file).copy()
        df["データ種別"] = "VR"
        df["カテゴリー"] = "ターゲット"
        df_list.append(df)

    # =========================
    # TVAL
    # =========================
    if tval_all_file is not None:
        df = make_df_heatmap_TVAL(tval_all_file).copy()
        df["データ種別"] = "TVAL"
        df["カテゴリー"] = "個人全体"
        df_list.append(df)

    if tval_target_file is not None:
        df = make_df_heatmap_TVAL(tval_target_file).copy()
        df["データ種別"] = "TVAL"
        df["カテゴリー"] = "ターゲット"
        df_list.append(df)

    # =========================
    # REVISIO
    # =========================
    if revisio_file is not None:
        df = make_df_heatmap_REVISIO(revisio_file).copy()
        df["データ種別"] = "REVISIO"
        df["カテゴリー"] = "個人全体"  # REVISIOは基本これ
        df_list.append(df)

    # =========================
    # 空チェック
    # =========================
    if not df_list:
        return pd.DataFrame(
            columns=["時間帯", "月", "火", "水", "木", "金", "土", "日", "局", "データ種別", "カテゴリー"]
        )

    # =========================
    # カラム揃え
    # =========================
    base_cols = ["時間帯", "月", "火", "水", "木", "金", "土", "日", "局", "データ種別", "カテゴリー"]

    aligned = []
    for df in df_list:
        for col in base_cols:
            if col not in df.columns:
                df[col] = pd.NA
        aligned.append(df[base_cols])

    # =========================
    # 結合
    # =========================
    df_heatmap_real = pd.concat(aligned, ignore_index=True)

    # =========================
    # 並び順を指定
    # =========================
    data_type_order = ["VR", "TVAL", "REVISIO"]
    category_order = ["個人全体", "ターゲット"]
    station_order = ["NTV", "TBS", "CX", "EX", "TX"]

    # カテゴリ順を設定
    df_heatmap_real["データ種別"] = pd.Categorical(
        df_heatmap_real["データ種別"],
        categories=data_type_order,
        ordered=True
    )

    df_heatmap_real["カテゴリー"] = pd.Categorical(
        df_heatmap_real["カテゴリー"],
        categories=category_order,
        ordered=True
    )

    df_heatmap_real["局"] = pd.Categorical(
        df_heatmap_real["局"],
        categories=station_order,
        ordered=True
    )

    # 時間帯ソート用
    df_heatmap_real["_sort_time"] = (
        df_heatmap_real["時間帯"]
        .astype(str)
        .str.strip()
        .str.replace(":", "", regex=False)
        .astype(int)
    )

    # 並び替え
    df_heatmap_real = (
        df_heatmap_real.sort_values(["データ種別", "カテゴリー", "局", "_sort_time"])
        .drop(columns="_sort_time")
        .reset_index(drop=True)
    )

    return df_heatmap_real