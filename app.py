import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="出荷ファイル自動整形Webシステム", layout="wide")
st.title("📦 出荷ファイル自動整形Webシステム")
st.caption("メルカリ用（原本通り非表示・列幅設定済Excel）とゆうプリR用（CSV）を自動判別し、sy・sy以外に仕分けて出力します。")

st.subheader("1. CSVファイルのアップロード")
uploaded_files = st.file_uploader(
    "メルカリ用CSV、ゆうプリR用CSVを投入（複数同時投入可能）",
    type=["csv"],
    accept_multiple_files=True
)

def create_formatted_excel(dataframe):
    """原本Excelの非表示列（A-D, G-H, J-N, P-Z）と列幅（E, F）を完全再現"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Sheet1")
        worksheet = writer.sheets["Sheet1"]

        # 原本で非表示になっている列インデックス（1始まり）
        # A-D(1-4), G-H(7-8), J-N(10-14), P-Z(16-26)
        hidden_indices = set(list(range(1, 5)) + list(range(7, 9)) + list(range(10, 15)) + list(range(16, 27)))

        for idx in range(1, len(dataframe.columns) + 1):
            col_letter = get_column_letter(idx)
            dim = worksheet.column_dimensions[col_letter]
            
            if idx in hidden_indices:
                dim.hidden = True
                dim.width = 0
            elif col_letter == "E":
                dim.hidden = False
                dim.width = 16.125
            elif col_letter == "F":
                dim.hidden = False
                dim.width = 69.5
            else:
                dim.hidden = False
                dim.width = 13.0

    return output.getvalue()

if uploaded_files:
    st.subheader("2. 自動仕分け・ダウンロード")

    for f in uploaded_files:
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
        is_mercari = False

        # メルカリ側判定
        for col in ["original_product_id", "商品管理番号", "管理番号"]:
            if col in df.columns:
                target_col = col
                is_mercari = True
                break

        # ゆうプリR側判定
        if not target_col:
            for col in ["品名２", "品名2"]:
                if col in df.columns:
                    target_col = col
                    break
            if not target_col and len(df.columns) >= 30:
                target_col = df.columns[29]

        file_type_label = "メルカリ用 (Excel出力)" if is_mercari else "ゆうプリR用 (CSV出力)"
        st.markdown("---")
        st.markdown(f"#### 📄 元ファイル: `{f.name}` 【{file_type_label}】")

        if target_col:
            is_sy = df[target_col].astype(str).str.contains(r"sy", case=False, na=False)
            df_sy = df[is_sy]
            df_other = df[~is_sy]

            st.write(f"判定列: `{target_col}` ｜ 全体: **{len(df)}** 件 ➔ `sy`: **{len(df_sy)}** 件 ／ `sy以外`: **{len(df_other)}** 件")

            col_sy, col_other = st.columns(2)
            base_name = f.name.rsplit(".", 1)[0]

            # sy のみ
            with col_sy:
                st.markdown("##### 👤 【sy のみ】")
                st.dataframe(df_sy.head(3))

                if is_mercari:
                    excel_data_sy = create_formatted_excel(df_sy)
                    out_name_sy = f"sy_{base_name}.xlsx"
                    st.download_button(
                        label=f"⬇️ 「{out_name_sy}」を保存 (Excel)",
                        data=excel_data_sy,
                        file_name=out_name_sy,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_sy_{f.name}"
                    )
                else:
                    csv_buf_sy = io.BytesIO()
                    df_sy.to_csv(csv_buf_sy, index=False, encoding="cp932", errors="replace")
                    out_name_sy = f"sy_{base_name}.csv"
                    st.download_button(
                        label=f"⬇️ 「{out_name_sy}」を保存 (CSV)",
                        data=csv_buf_sy.getvalue(),
                        file_name=out_name_sy,
                        mime="text/csv",
                        key=f"dl_sy_{f.name}"
                    )

            # sy 以外
            with col_other:
                st.markdown("##### 👥 【sy 以外】")
                st.dataframe(df_other.head(3))

                if is_mercari:
                    excel_data_other = create_formatted_excel(df_other)
                    out_name_other = f"sy以外_{base_name}.xlsx"
                    st.download_button(
                        label=f"⬇️ 「{out_name_other}」を保存 (Excel)",
                        data=excel_data_other,
                        file_name=out_name_other,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_other_{f.name}"
                    )
                else:
                    csv_buf_other = io.BytesIO()
                    df_other.to_csv(csv_buf_other, index=False, encoding="cp932", errors="replace")
                    out_name_other = f"sy以外_{base_name}.csv"
                    st.download_button(
                        label=f"⬇️ 「{out_name_other}」を保存 (CSV)",
                        data=csv_buf_other.getvalue(),
                        file_name=out_name_other,
                        mime="text/csv",
                        key=f"dl_other_{f.name}"
                    )
        else:
            st.error(f"`{f.name}` 内に判定用列が見つかりませんでした。")
