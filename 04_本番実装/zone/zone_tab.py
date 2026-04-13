import pandas as pd
import streamlit as st

# 曜日順
WEEKDAYS = ["月", "火", "水", "木", "金", "土", "日"]

# 画面に表示するゾーン名一覧
ZONE_NAMES = ["全日", "ヨの字", "コの字", "逆L", "ATT"]


# =========================================================
# 固定の時間帯一覧
# - ゾーン表で使う時間帯を固定で作成
# - 05:00 ～ 28:00 までを1時間刻みで返す
# =========================================================
def get_default_times():
    return [f"{h:02d}:00" for h in range(5, 29)]


# =========================================================
# 空ゾーン
# - 全セル False の DataFrame を作成
# - index: 時間帯
# - columns: 曜日
# =========================================================
def create_empty_zone_df(times):
    return pd.DataFrame(False, index=times, columns=WEEKDAYS)


# =========================================================
# 指定時間帯をON
# - 指定した曜日 × 指定した時間帯範囲を True にする
# =========================================================
def apply_zone_on(zone_df, days, time_start, time_end):
    target_times = [
        t for t in zone_df.index
        if time_start <= int(str(t).split(":")[0]) <= time_end
    ]
    zone_df.loc[target_times, days] = True


# =========================================================
# ゾーン初期値
# - ゾーン名に応じたデフォルトの ON / OFF パターンを作成
# =========================================================
def build_zone_preset(zone_name, times):
    zone_df = create_empty_zone_df(times)

    if zone_name == "ATT":
        apply_zone_on(zone_df, WEEKDAYS, 5, 28)

    elif zone_name == "全日":
        apply_zone_on(zone_df, WEEKDAYS, 5, 25)

    elif zone_name == "ヨの字":
        apply_zone_on(zone_df, ["月", "火", "水", "木", "金"], 5, 8)
        apply_zone_on(zone_df, ["月", "火", "水", "木", "金"], 12, 13)
        apply_zone_on(zone_df, ["月", "火", "水", "木", "金"], 19, 25)
        apply_zone_on(zone_df, ["土", "日"], 5, 25)

    elif zone_name == "コの字":
        apply_zone_on(zone_df, ["月", "火", "水", "木", "金"], 5, 8)
        apply_zone_on(zone_df, ["月", "火", "水", "木", "金"], 19, 25)
        apply_zone_on(zone_df, ["土", "日"], 5, 25)

    elif zone_name == "逆L":
        apply_zone_on(zone_df, ["月", "火", "水", "木", "金"], 19, 25)
        apply_zone_on(zone_df, ["土", "日"], 5, 25)

    return zone_df


# =========================================================
# checkbox 用 key
# - ゾーン名 × 時間帯 × 曜日 で一意な key を作る
# =========================================================
def zone_checkbox_key(zone_name, time_label, day_label):
    return f"zone_chk_{zone_name}_{time_label}_{day_label}"


# =========================================================
# ゾーン再描画要求フラグ用 key
# - 初期化時に、そのゾーンの checkbox を作り直すために使う
# =========================================================
def zone_reset_request_key(zone_name):
    return f"zone_reset_requested_{zone_name}"


# =========================================================
# session_state 初期化
# - zone_filters : 反映済みの正式データ
# - zone_drafts  : 編集中の一時データ
# - shared_wish_zone / shared_wish_worst_n : 希望枠タブ共有用
# - zone_reset_requested_* : ゾーンごとの再生成フラグ
# =========================================================
def init_zone_state():
    times = get_default_times()

    if "zone_filters" not in st.session_state:
        st.session_state["zone_filters"] = {
            zone_name: build_zone_preset(zone_name, times)
            for zone_name in ZONE_NAMES
        }

    if "zone_drafts" not in st.session_state:
        st.session_state["zone_drafts"] = {
            zone_name: st.session_state["zone_filters"][zone_name].copy()
            for zone_name in ZONE_NAMES
        }

    if "shared_wish_zone" not in st.session_state:
        st.session_state["shared_wish_zone"] = "全日"

    if "shared_wish_worst_n" not in st.session_state:
        st.session_state["shared_wish_worst_n"] = 0

    for zone_name in ZONE_NAMES:
        req_key = zone_reset_request_key(zone_name)
        if req_key not in st.session_state:
            st.session_state[req_key] = False


# =========================================================
# draft -> checkbox 初期値反映
# - zone_drafts の値を checkbox の初期値に流し込む
# - 初期化要求があるときだけ既存 checkbox key を消して再生成する
# =========================================================
def ensure_zone_checkbox_defaults(zone_name):
    draft_df = st.session_state["zone_drafts"][zone_name]
    reset_requested = st.session_state.get(zone_reset_request_key(zone_name), False)

    for time_label in draft_df.index:
        for day_label in WEEKDAYS:
            key = zone_checkbox_key(zone_name, time_label, day_label)

            if reset_requested and key in st.session_state:
                del st.session_state[key]

            if key not in st.session_state:
                st.session_state[key] = bool(draft_df.loc[time_label, day_label])

    if reset_requested:
        st.session_state[zone_reset_request_key(zone_name)] = False


# =========================================================
# checkbox -> draft に取り込み
# - 画面上の checkbox の状態を読み取って zone_drafts に保存する
# =========================================================
def read_zone_draft_from_widgets(zone_name):
    draft_df = st.session_state["zone_drafts"][zone_name]
    times = list(draft_df.index)

    updated_df = create_empty_zone_df(times)

    for time_label in times:
        for day_label in WEEKDAYS:
            updated_df.loc[time_label, day_label] = bool(
                st.session_state.get(
                    zone_checkbox_key(zone_name, time_label, day_label),
                    False,
                )
            )

    st.session_state["zone_drafts"][zone_name] = updated_df


# =========================================================
# 個別初期化
# - そのゾーンの draft だけ preset に戻す
# - 次回描画時に checkbox を再生成させる
# =========================================================
def reset_zone_draft(zone_name):
    times = get_default_times()
    st.session_state["zone_drafts"][zone_name] = build_zone_preset(zone_name, times)
    st.session_state[zone_reset_request_key(zone_name)] = True


# =========================================================
# 個別反映
# - そのゾーンの checkbox 状態を draft に取り込み
# - その内容を正式データ zone_filters に保存する
# =========================================================
def apply_zone_filter(zone_name):
    read_zone_draft_from_widgets(zone_name)
    st.session_state["zone_filters"][zone_name] = st.session_state["zone_drafts"][zone_name].copy()


# =========================================================
# 全体反映
# - すべてのゾーンをまとめて正式データへ保存する
# =========================================================
def apply_all_zone_filters():
    for zone_name in ZONE_NAMES:
        apply_zone_filter(zone_name)


# =========================================================
# 希望枠用ゾーン取得
# - 希望枠タブで使う、反映済みの正式ゾーンを返す
# =========================================================
def get_wish_zone_df(selected_zone_name):
    return st.session_state["zone_filters"][selected_zone_name]


# =========================================================
# エディタ表示
# - 1ゾーン分の checkbox 表を描画
# - 行: 時間帯
# - 列: 曜日
# - ATT は編集不可
# =========================================================
def render_zone_editor_compact(zone_name):
    ensure_zone_checkbox_defaults(zone_name)

    draft_df = st.session_state["zone_drafts"][zone_name]
    times = list(draft_df.index)

    header_cols = st.columns([1.15] + [0.62] * len(WEEKDAYS), gap=None)
    header_cols[0].markdown(" ")
    for i, day in enumerate(WEEKDAYS):
        header_cols[i + 1].markdown(
            f"<div style='font-size:0.66rem; text-align:center'>{day}</div>",
            unsafe_allow_html=True,
        )

    for time_label in times:
        row_cols = st.columns([1.15] + [0.62] * len(WEEKDAYS), gap=None)
        row_cols[0].markdown(
            f"<div style='font-size:0.68rem; padding-top:0.10rem'>{time_label}</div>",
            unsafe_allow_html=True,
        )

        for i, day_label in enumerate(WEEKDAYS):
            row_cols[i + 1].checkbox(
                "",
                key=zone_checkbox_key(zone_name, time_label, day_label),
                label_visibility="collapsed",
                disabled=(zone_name == "ATT"),
            )


# =========================================================
# 1ゾーン分を fragment 化
# - form は使わない
# - checkbox の値は session_state に直接入る
# - 個別ボタンは「初期化」のみ
# =========================================================
@st.fragment
def render_single_zone_fragment(zone_name):
    st.markdown(f"### {zone_name}")

    render_zone_editor_compact(zone_name)

    if st.button("初期化", key=f"reset_zone_{zone_name}", use_container_width=True):
        reset_zone_draft(zone_name)
        st.rerun(scope="fragment")