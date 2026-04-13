# ライブラリの読み込み
import streamlit as st

# =========================================================
# 共通アップロードUI
# - 各種ファイルアップロード欄を共通関数で描画する
# - title        : 見出し
# - file_label   : file_uploader に表示する説明文
# - key          : session_state に保存するキー
# - allowed_types: 許可する拡張子
# =========================================================
def render_upload_box(title, file_label, key, allowed_types):
    # アップロード欄の見出し表示
    st.markdown(f"#### {title}")

    # ファイルアップローダーを表示
    st.file_uploader(
        file_label,
        type=allowed_types,
        key=key,
    )

    # アップロード済みファイルを session_state から取得
    uploaded = st.session_state.get(key)

    # ファイルがアップロードされていればファイル名を表示
    if uploaded is not None:
        st.success(f"アップロード完了: {uploaded.name}")