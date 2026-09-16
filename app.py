import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="出荷ファイル自動整形Webシステム", layout="wide")
st.title("📦 出荷ファイル自動整形Webシステム")
st.caption("ブラウザ上で元データを読み込み、チェックボックスで担当者を選ぶだけで、不要な行の非表示・整形を自動で行います。")

# 1. 元データのアップロード（複数ファイルを許可してエラーを排除）
st.subheader("1. 元データのアップロード")
uploaded_files = st.file_uploader(
    "メルカリ等の売却済みデータ CSV (複数選択・ドラッグ可)",
    type=["csv"],
    accept_multiple_files=True
)

if uploaded_files:
    dfs = []
    for f in uploaded_files:
        try:
            # Shift-JIS / UTF-8 両対応
            try:
                df_temp = pd.read_csv(f, encoding="cp932")
            except Exception:
                f.seek(0)
                df_temp = pd.read_csv(f, encoding="utf-8")
            dfs.append(df_temp)
        except Exception:
            continue

    if dfs:
        df = pd.concat(dfs, ignore_index=True).drop_duplicates()
        st.success(f"ファイルを正常に読み込みました（全 {len(df)} 件）")

        # 識別カラムの探索
        target_col = None
        for col in ["original_product_id", "商品管理番号", "管理番号"]:
            if col in df.columns:
                target_col = col
                break

        if target_col:
            st.caption(f"データ内の識別カラム： {target_col}")

            # 2. 残す担当者の選択
            st.subheader("2. 残す担当者の選択")
            st.write("チェックを入れた担当者のデータのみを残し、他を自動で非表示（除外）します。")

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                chk_ami = st.checkbox("担当: ami", value=True)
            with col2:
                chk_kj = st.checkbox("担当: kj", value=True)
            with col3:
                chk_mb = st.checkbox("担当: mb", value=True)
            with col4:
                chk_sy = st.checkbox("担当: sy", value=True)

            selected_tags = []
            if chk_ami: selected_tags.append("ami")
            if chk_kj: selected_tags.append("kj")
            if chk_mb: selected_tags.append("mb")
            if chk_sy: selected_tags.append("sy")

            # 3. 処理の実行
            st.subheader("3. 処理の実行")
            if st.button("整形データを生成する", type="primary"):
                if not selected_tags:
                    st.warning("担当者を1名以上選択してください。")
                else:
                    pattern = "|".join(selected_tags)
                    filtered_df = df[df[target_col].astype(str).str.contains(pattern, case=False, na=False)]

                    st.write(f"抽出結果: {len(filtered_df)} 件")
                    st.dataframe(filtered_df)

                    csv_buffer = io.BytesIO()
                    filtered_df.to_csv(csv_buffer, index=False, encoding="cp932", errors="replace")
                    st.download_button(
                        label="整形済みCSVをダウンロード",
                        data=csv_buffer.getvalue(),
                        file_name="shipping_formatted.csv",
                        mime="text/csv"
                    )
        else:
            st.error("識別用のカラム（original_product_id等）が見つかりませんでした。")
