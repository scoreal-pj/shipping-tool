import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="出荷ファイル自動整形Webシステム", layout="wide")
st.title("📦 出荷ファイル自動整形Webシステム")
st.caption("メルカリ用・ゆうプリR用の各CSVを個別に整形・ダウンロードします（結合は行いません）。")

# 1. 担当者の選択（共通設定）
st.subheader("1. 残す担当者の選択")
st.write("チェックを入れた担当者の行のみを残し、他の行を除外します。")
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

# 2. ファイルアップロード（複数可）
st.subheader("2. CSVファイルのアップロード")
uploaded_files = st.file_uploader(
    "メルカリ用・ゆうプリR用のCSVを投入（複数まとめて投入可能）",
    type=["csv"],
    accept_multiple_files=True
)

if uploaded_files:
    if not selected_tags:
        st.warning("残す担当者を1名以上選択してください。")
    else:
        pattern = "|".join(selected_tags)
        target_cols = ["original_product_id", "商品管理番号", "管理番号"]

        st.subheader("3. 各ファイルの整形結果・ダウンロード")

        for f in uploaded_files:
            try:
                # 文字コード判定（Shift-JIS / UTF-8）
                try:
                    df = pd.read_csv(f, encoding="cp932")
                except Exception:
                    f.seek(0)
                    df = pd.read_csv(f, encoding="utf-8")
            except Exception as e:
                st.error(f"ファイル読み込み失敗 ({f.name}): {e}")
                continue

            # 識別カラムの特定
            target_col = None
            for col in target_cols:
                if col in df.columns:
                    target_col = col
                    break

            st.markdown(f"---")
            st.markdown(f"#### 📄 元ファイル: `{f.name}`")

            if target_col:
                # 該当担当者のみ抽出
                filtered_df = df[df[target_col].astype(str).str.contains(pattern, case=False, na=False)]
                
                st.write(f"抽出結果: **{len(filtered_df)}** 件 (元データ: {len(df)} 件 / 判定列: `{target_col}`)")
                st.dataframe(filtered_df.head(5))

                # ダウンロード用バッファ（ファイル名ごとに個別生成）
                csv_buffer = io.BytesIO()
                filtered_df.to_csv(csv_buffer, index=False, encoding="cp932", errors="replace")
                
                out_name = f"整形済_{f.name}"
                st.download_button(
                    label=f"⬇️ 「{out_name}」をダウンロード",
                    data=csv_buffer.getvalue(),
                    file_name=out_name,
                    mime="text/csv",
                    key=f"dl_{f.name}"
                )
            else:
                st.error(f"`{f.name}` 内に対象の管理番号カラム（original_product_id等）が見つかりませんでした。")
