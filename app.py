import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="出荷ファイル自動整形Webシステム", layout="wide")
st.title("📦 出荷ファイル自動整形Webシステム")
st.caption("メルカリ用CSV（original_product_id）とゆうプリR用CSV（品名２/AD列）を個別に判定・整形します。")

# 1. 残す担当者の選択
st.subheader("1. 残す担当者の選択")
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

# 2. ファイルアップロード
st.subheader("2. CSVファイルのアップロード")
uploaded_files = st.file_uploader(
    "メルカリ用CSV、ゆうプリR用CSVをまとめて投入（2ファイル同時投入可）",
    type=["csv"],
    accept_multiple_files=True
)

if uploaded_files:
    if not selected_tags:
        st.warning("残す担当者を1名以上選択してください。")
    else:
        pattern = "|".join(selected_tags)
        st.subheader("3. 各ファイルの整形結果・ダウンロード")

        for f in uploaded_files:
            # Shift-JIS / UTF-8 読み込み
            try:
                try:
                    df = pd.read_csv(f, encoding="cp932")
                except Exception:
                    f.seek(0)
                    df = pd.read_csv(f, encoding="utf-8")
            except Exception as e:
                st.error(f"ファイル読み込み失敗 ({f.name}): {e}")
                continue

            target_col = None
            detected_type = ""

            # 判定パターン1: メルカリ側 (original_product_id, 商品管理番号, 管理番号)
            for col in ["original_product_id", "商品管理番号", "管理番号"]:
                if col in df.columns:
                    target_col = col
                    detected_type = "メルカリ用データ"
                    break

            # 判定パターン2: ゆうプリR側 (品名２ または AD列 = 30列目 / 0始まりの29番目)
            if not target_col:
                for col in ["品名２", "品名2"]:
                    if col in df.columns:
                        target_col = col
                        detected_type = "ゆうプリR用データ (品名２)"
                        break

            # カラム名一致がなかった場合、AD列（列インデックス29）を確認
            if not target_col and len(df.columns) >= 30:
                target_col = df.columns[29]
                detected_type = f"ゆうプリR用データ (AD列: {target_col})"

            st.markdown("---")
            st.markdown(f"#### 📄 `{f.name}` 【{detected_type if detected_type else '未判定'}】")

            if target_col:
                # 指定担当者で抽出
                filtered_df = df[df[target_col].astype(str).str.contains(pattern, case=False, na=False)]
                
                st.write(f"判定列: `{target_col}` | 抽出結果: **{len(filtered_df)}** 件 / 元データ: {len(df)} 件")
                st.dataframe(filtered_df.head(3))

                # ダウンロード用バッファ（ファイルごとに分離）
                csv_buffer = io.BytesIO()
                filtered_df.to_csv(csv_buffer, index=False, encoding="cp932", errors="replace")
                
                out_name = f"整形済_{f.name}"
                st.download_button(
                    label=f"⬇️ 「{out_name}」をダウンロード",
                    data=csv_buffer.getvalue(),
                    file_name=out_name,
                    mime="text/csv",
                    key=f"btn_{f.name}"
                )
            else:
                st.error(f"`{f.name}` 内に判定用列（original_product_id、または 品名２/AD列）が見つかりませんでした。")
