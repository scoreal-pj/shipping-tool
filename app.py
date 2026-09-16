import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="業務支援システム", layout="wide")

# サイドバーメニュー
menu = st.sidebar.radio("機能を選択", ["📦 出荷CSV整形", "📮 ゆうびん後納計算"])

# ==========================================
# 1. 出荷CSV整形機能
# ==========================================
if menu == "📦 出荷CSV整形":
    st.title("📦 出荷ファイル自動整形")
    st.caption("メルカリ用（原本通り非表示・列幅設定済Excel）とゆうプリR用（CSV）を自動判別し、sy・sy以外に仕分けて出力します。")

    uploaded_files = st.file_uploader(
        "メルカリ用CSV、ゆうプリR用CSVを投入（複数同時投入可能）",
        type=["csv"],
        accept_multiple_files=True
    )

    def create_formatted_excel(dataframe):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            dataframe.to_excel(writer, index=False, sheet_name="Sheet1")
            worksheet = writer.sheets["Sheet1"]

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
        st.subheader("自動仕分け・ダウンロード")
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

            for col in ["original_product_id", "商品管理番号", "管理番号"]:
                if col in df.columns:
                    target_col = col
                    is_mercari = True
                    break

            if not target_col:
                for col in ["品名２", "品名2"]:
                    if col in df.columns:
                        target_col = col
                        break
                if not target_col and len(df.columns) >= 30:
                    target_col = df.columns[29]

            file_type_label = "メルカリ用 (Excel出力)" if is_mercari else "ゆうプリR用 (CSV出力)"
            st.markdown("---")
            st.markdown(f"#### 📄 `{f.name}` 【{file_type_label}】")

            if target_col:
                is_sy = df[target_col].astype(str).str.contains(r"sy", case=False, na=False)
                df_sy = df[is_sy]
                df_other = df[~is_sy]

                st.write(f"判定列: `{target_col}` ｜ 全体: **{len(df)}** 件 ➔ `sy`: **{len(df_sy)}** 件 ／ `sy以外`: **{len(df_other)}** 件")

                col_sy, col_other = st.columns(2)
                base_name = f.name.rsplit(".", 1)[0]

                with col_sy:
                    st.markdown("##### 👤 【sy のみ】")
                    st.dataframe(df_sy.head(3))
                    if is_mercari:
                        excel_data_sy = create_formatted_excel(df_sy)
                        st.download_button(label=f"⬇️ 「sy_{base_name}.xlsx」を保存", data=excel_data_sy, file_name=f"sy_{base_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_sy_{f.name}")
                    else:
                        csv_buf_sy = io.BytesIO()
                        df_sy.to_csv(csv_buf_sy, index=False, encoding="cp932", errors="replace")
                        st.download_button(label=f"⬇️ 「sy_{base_name}.csv」を保存", data=csv_buf_sy.getvalue(), file_name=f"sy_{base_name}.csv", mime="text/csv", key=f"dl_sy_{f.name}")

                with col_other:
                    st.markdown("##### 👥 【sy 以外】")
                    st.dataframe(df_other.head(3))
                    if is_mercari:
                        excel_data_other = create_formatted_excel(df_other)
                        st.download_button(label=f"⬇️ 「sy以外_{base_name}.xlsx」を保存", data=excel_data_other, file_name=f"sy以外_{base_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_other_{f.name}")
                    else:
                        csv_buf_other = io.BytesIO()
                        df_other.to_csv(csv_buf_other, index=False, encoding="cp932", errors="replace")
                        st.download_button(label=f"⬇️ 「sy以外_{base_name}.csv」を保存", data=csv_buf_other.getvalue(), file_name=f"sy以外_{base_name}.csv", mime="text/csv", key=f"dl_other_{f.name}")
            else:
                st.error(f"`{f.name}` 内に判定用列が見つかりませんでした。")

# ==========================================
# 2. ゆうびん後納計算機能（エクセルの完全再現）
# ==========================================
elif menu == "📮 ゆうびん後納計算":
    st.title("📮 ゆうびん後納計算システム")
    st.caption("エクセルの計算式に基づき、通数・小計・合計・正規料金・粗利を自動計算します。")

    tab1, tab2 = st.tabs(["📦 パケット専用（粗利計算）", "📑 メール・パケット・定形外（総合）"])

    # ----------------------------------------------------
    # タブ1: パケット専用（シート「パケット」の再現）
    # ----------------------------------------------------
    with tab1:
        st.subheader("ゆうパケット 運賃・粗利計算")
        st.write("個数を入力すると合計運賃と正規料金との差額（粗利）が自動計算されます。")

        col1, col2, col3 = st.columns(3)
        with col1:
            p_cnt_1cm = st.number_input("1cm 個数 (後納173円 / 正規250円)", min_value=0, value=0, step=1, key="p1")
        with col2:
            p_cnt_2cm = st.number_input("2cm 個数 (後納204円 / 正規310円)", min_value=0, value=0, step=1, key="p2")
        with col3:
            p_cnt_3cm = st.number_input("3cm 個数 (後納280円 / 正規360円)", min_value=0, value=0, step=1, key="p3")

        # 計算
        tot_1cm = p_cnt_1cm * 173
        reg_1cm = p_cnt_1cm * 250
        prof_1cm = reg_1cm - tot_1cm

        tot_2cm = p_cnt_2cm * 204
        reg_2cm = p_cnt_2cm * 310
        prof_2cm = reg_2cm - tot_2cm

        tot_3cm = p_cnt_3cm * 280
        reg_3cm = p_cnt_3cm * 360
        prof_3cm = reg_3cm - tot_3cm

        all_cnt = p_cnt_1cm + p_cnt_2cm + p_cnt_3cm
        all_tot = tot_1cm + tot_2cm + tot_3cm
        all_prof = prof_1cm + prof_2cm + prof_3cm

        # サマリーカード表示
        m1, m2, m3 = st.columns(3)
        m1.metric("総個数", f"{all_cnt} 個")
        m2.metric("後納運賃合計", f"{all_tot:,} 円")
        m3.metric("粗利合計", f"{all_prof:,} 円")

        # 詳細テーブル
        df_packet = pd.DataFrame([
            {"区分": "1cm", "運賃": 173, "個数": p_cnt_1cm, "合計": tot_1cm, "正規料金": 250, "正規合計": reg_1cm, "粗利": prof_1cm},
            {"区分": "2cm", "運賃": 204, "個数": p_cnt_2cm, "合計": tot_2cm, "正規料金": 310, "正規合計": reg_2cm, "粗利": prof_2cm},
            {"区分": "3cm", "運賃": 280, "個数": p_cnt_3cm, "合計": tot_3cm, "正規料金": 360, "正規合計": reg_3cm, "粗利": prof_3cm},
            {"区分": "【合計】", "運賃": "-", "個数": all_cnt, "合計": all_tot, "正規料金": "-", "正規合計": reg_1cm + reg_2cm + reg_3cm, "粗利": all_prof},
        ])
        st.dataframe(df_packet, use_container_width=True, hide_index=True)

    # ----------------------------------------------------
    # タブ2: メール・パケット・定形外（総合集計シートの再現）
    # ----------------------------------------------------
    with tab2:
        st.subheader("メール・パケット・定形外 総合計算")
        
        c_mail, c_pack, c_teikei = st.columns(3)

        # ゆうメール
        with c_mail:
            st.markdown("#### ✉️ ゆうメール")
            m_500 = st.number_input("500g (136円)", min_value=0, value=0, step=1, key="m_500")
            m_1k = st.number_input("1kg (193円)", min_value=0, value=0, step=1, key="m_1k")
            m_2k = st.number_input("2kg (286円)", min_value=0, value=0, step=1, key="m_2k")
            m_3k = st.number_input("3kg (422円)", min_value=0, value=0, step=1, key="m_3k")

        # ゆうパケット
        with c_pack:
            st.markdown("#### 📦 ゆうパケット")
            pk_1 = st.number_input("1cm (173円)", min_value=0, value=0, step=1, key="pk_1")
            pk_2 = st.number_input("2cm (204円)", min_value=0, value=0, step=1, key="pk_2")
            pk_3 = st.number_input("3cm (280円)", min_value=0, value=0, step=1, key="pk_3")

        # 定形外
        with c_teikei:
            st.markdown("#### 📮 定形外")
            t_510 = st.number_input("510円", min_value=0, value=0, step=1, key="t_510")
            t_350 = st.number_input("350円", min_value=0, value=0, step=1, key="t_350")
            t_220 = st.number_input("220円", min_value=0, value=0, step=1, key="t_220")

        # 計算
        sub_m = (m_500 * 136) + (m_1k * 193) + (m_2k * 286) + (m_3k * 422)
        cnt_m = m_500 + m_1k + m_2k + m_3k

        sub_pk = (pk_1 * 173) + (pk_2 * 204) + (pk_3 * 280)
        cnt_pk = pk_1 + pk_2 + pk_3

        sub_t = (t_510 * 510) + (t_350 * 350) + (t_220 * 220)
        cnt_t = t_510 + t_350 + t_220

        grand_cnt = cnt_m + cnt_pk + cnt_t
        grand_total = sub_m + sub_pk + sub_t

        st.markdown("---")
        st.markdown("### 📊 総合計")
        g1, g2, g3, g4 = st.columns(4)
        g1.metric("ゆうメール小計", f"{cnt_m} 個 / {sub_m:,} 円")
        g2.metric("ゆうパケット小計", f"{cnt_pk} 個 / {sub_pk:,} 円")
        g3.metric("定形外小計", f"{cnt_t} 個 / {sub_t:,} 円")
        g4.metric("総計（総合計）", f"{grand_cnt} 個 / {grand_total:,} 円")

        # 計算結果Excelのダウンロード
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
            df_all = pd.DataFrame([
                {"種別": "ゆうメール", "規格": "500g", "運賃": 136, "個数": m_500, "小計": m_500 * 136},
                {"種別": "ゆうメール", "規格": "1kg", "運賃": 193, "個数": m_1k, "小計": m_1k * 193},
                {"種別": "ゆうメール", "規格": "2kg", "運賃": 286, "個数": m_2k, "小計": m_2k * 286},
                {"種別": "ゆうメール", "規格": "3kg", "運賃": 422, "個数": m_3k, "小計": m_3k * 422},
                {"種別": "ゆうパケット", "規格": "1cm", "運賃": 173, "個数": pk_1, "小計": pk_1 * 173},
                {"種別": "ゆうパケット", "規格": "2cm", "運賃": 204, "個数": pk_2, "小計": pk_2 * 204},
                {"種別": "ゆうパケット", "規格": "3cm", "運賃": 280, "個数": pk_3, "小計": pk_3 * 280},
                {"種別": "定形外", "規格": "510円", "運賃": 510, "個数": t_510, "小計": t_510 * 510},
                {"種別": "定形外", "規格": "350円", "運賃": 350, "個数": t_350, "小計": t_350 * 350},
                {"種別": "定形外", "規格": "220円", "運賃": 220, "個数": t_220, "小計": t_220 * 220},
                {"種別": "【総合計】", "規格": "-", "運賃": "-", "個数": grand_cnt, "小計": grand_total},
            ])
            df_all.to_excel(writer, index=False, sheet_name="後納計算結果")

        st.download_button(
            label="⬇️ この計算結果をExcelとして保存",
            data=excel_buf.getvalue(),
            file_name="後納計算結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
