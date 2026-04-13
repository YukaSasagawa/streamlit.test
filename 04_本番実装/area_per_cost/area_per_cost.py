import pandas as pd
import streamlit as st

ZONE_NAMES = ["全日", "ヨの字", "コの字", "逆L", "ATT"]

def sync_from_top(area: str):
    """
    上部のエリア選択 → 各エリア横の✅へ同期
    """
    value = st.session_state.get(f"top_{area}", False)
    st.session_state[f"row_{area}"] = value
    st.session_state["selected_area_map"][area] = value


def sync_from_row(area: str):
    """
    各エリア横の✅ → 上部のエリア選択へ同期
    """
    value = st.session_state.get(f"row_{area}", False)
    st.session_state[f"top_{area}"] = value
    st.session_state["selected_area_map"][area] = value


def init_area_state(areas):
    """
    session_state 初期化
    """
    if "selected_area_map" not in st.session_state:
        st.session_state["selected_area_map"] = {}

    for area in areas:
        if area not in st.session_state["selected_area_map"]:
            st.session_state["selected_area_map"][area] = False

        if f"top_{area}" not in st.session_state:
            st.session_state[f"top_{area}"] = st.session_state["selected_area_map"][area]

        if f"row_{area}" not in st.session_state:
            st.session_state[f"row_{area}"] = st.session_state["selected_area_map"][area]


def read_cost_excel(uploaded_file):
    """
    画像のようなExcelを読み込んで整形する関数
    想定:
      A列: エリア（結合セルあり）
      B列: 局
      C:G  パーコスト
      H:L  ターゲットINDEX
      M:Q  ターゲットコスト
    """
    # ヘッダーが複数行ある想定なので header=None で読む
    raw = pd.read_excel(uploaded_file, header=None)

    # 3行目以降がデータ本体想定（画像上の3行目あたりから）
    df = raw.iloc[2:, :17].copy()
    df.columns = [
        "エリア", "局",
        "パーコスト_全日", "パーコスト_ヨの字", "パーコスト_コの字", "パーコスト_逆L", "パーコスト_ATT",
        "ターゲットINDEX_全日", "ターゲットINDEX_ヨの字", "ターゲットINDEX_コの字", "ターゲットINDEX_逆L", "ターゲットINDEX_ATT",
        "ターゲットコスト_全日", "ターゲットコスト_ヨの字", "ターゲットコスト_コの字", "ターゲットコスト_逆L", "ターゲットコスト_ATT",
    ]

    # 結合セル対策
    df["エリア"] = df["エリア"].ffill()

    # 不要行除外
    df = df[df["局"].notna()].copy()
    df["局"] = df["局"].astype(str).str.strip()
    df["エリア"] = df["エリア"].astype(str).str.strip()

    # 数値化
    num_cols = [c for c in df.columns if c not in ["エリア", "局"]]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    return df


def calc_area_avg_target_cost(df):
    """
    エリアごとのターゲットコスト局平均
    """
    cols = [f"ターゲットコスト_{z}" for z in ZONE_NAMES]
    area_avg = (
        df.groupby("エリア")[cols]
        .mean()
        .reset_index()
    )

    rename_map = {f"ターゲットコスト_{z}": z for z in ZONE_NAMES}
    area_avg = area_avg.rename(columns=rename_map)

    return area_avg


def calc_selected_area_zone_average(area_avg_df):
    """
    ✅したエリアのターゲットコストのゾーン単位局平均を足し上げ × 100
    例:
      （関東の局平均 + 関西の局平均 + 中京の局平均） × 100
    """
    selected_areas = [
        area for area, selected in st.session_state["selected_area_map"].items()
        if selected
    ]

    if len(selected_areas) == 0:
        return pd.Series({z: 0 for z in ZONE_NAMES})

    target = area_avg_df[area_avg_df["エリア"].isin(selected_areas)]

    result = target[ZONE_NAMES].sum() * 100
    return result


def format_num(x):
    if pd.isna(x):
        return "-"
    return f"{x:,.0f}"