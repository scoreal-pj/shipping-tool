import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="出荷ファイル自動整形Webシステム", layout="wide")
st.title("📦 出荷ファイル自動整形Webシステム")
st.caption("投入された各CSVから「syのみ」と「sy以外（除外）」の2種類のデータを自動生成します。")

# ファイルアップロード
st.subheader("1. CSVファイルのアップロード")
uploaded_files = st.file_uploader(
    "メルカリ用CSV、ゆうプリR用CSVを投入（複数同時投入可能）",
    type=["csv"],
    accept_multiple_files=True
)

if uploaded_files:
    st.subheader("2. 自動仕分け・ダウンロード")

    for f in uploaded_files:
        # 文字コード自動判定（Shift-JIS / UTF-8）
        try:
            try:
                df = pd.read_csv(f, encoding="cp932")
            except Exception:
                f.seek(0)
                df = pd.read_csv(f, encoding="utf-8")
        except Exception as e:
            st.error(f"ファイル読み込み失敗 ({f.name}): {e}")
            continue

        # 判定列の特定
        target_col = None
        detected_type = ""

        # 1. メルカリ側
        for col in ["original_product_id", "商品管理番号", "管理番号"]:
            if col in df.columns:
                target_col = col
                detected_type = "メルカリ用データ"
                break

        # 2. ゆうプリR側
        if not target_col:
            for col in ["品名２", "品名2"]:
                if col in df.columns:
                    target_col = col
                    detected_type = "ゆうプリR用データ (品名２)"
                    break

        if not target_col and len(df.columns) >= 30:
            target_col = df.columns[29]
            detected_type = f"ゆうプリR用データ (AD列: {target_col})"

        st.markdown("---")
        st.markdown(f"#### 📄 元ファイル: `{f.name}` 【{detected_type if detected_type else '未判定'}】")

        if target_col:
            # syを含むかどうかの真偽値マスク（大文字・小文字両対応）
            is_sy = df[target_col].astype(str).str.contains(r"sy", case=False, na=False)
            
            df_sy = df[is_sy]         # sy のみ
            df_other = df[~is_sy]     # sy 以外すべて

            st.write(f"判定列: `{target_col}` ｜ 全体: **{len(df)}** 件 ➔ `sy`: **{len(df_sy)}** 件 ／ `sy以外`: **{len(df_other)}** 件")

            col_sy, col_other = st.columns(2)

            # --- sy のみのダウンロード ---
            with col_sy:
                st.markdown("##### 👤 【sy のみ】")
                st.dataframe(df_sy.head(3))
                
                buf_sy = io.BytesIO()
                df_sy.to_csv(buf_sy, index=False, encoding="cp932", errors="replace")
                out_name_sy = f"sy_{f.name}"
                
                st.download_button(
                    label=f"⬇️ 「{out_name_sy}」を保存 ({len(df_sy)}件)",
                    data=buf_sy.getvalue(),
                    file_name=out_name_sy,
                    mime="text/csv",
                    key=f"dl_sy_{f.name}"
                )

            # --- sy 以外のダウンロード ---
            with col_other:
                st.markdown("##### 👥 【sy 以外（除外）】")
                st.dataframe(df_other.head(3))
                
                buf_other = io.BytesIO()
                df_other.to_csv(buf_other, index=False, encoding="cp932", errors="replace")
                out_name_other = f"sy以外_{f.name}"
                
                st.download_button(
                    label=f"⬇️ 「{out_name_other}」を保存 ({len(df_other)}件)",
                    data=buf_other.getvalue(),
                    file_name=out_name_other,
                    mime="text/csv",
                    key=f"dl_other_{f.name}"
                )
        else:
            st.error(f"`{f.name}` 内に判定用列（original_product_id または 品名２/AD列）が見つかりませんでした。")
