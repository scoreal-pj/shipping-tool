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
# 2. ゆうびん後納計算機能
# ==========================================
elif menu == "📮 ゆうびん後納計算":
    st.title("📮 ゆうびん後納計算（ゆうパケット）")
    st.caption("1cm / 2cm / 3cm の個数を入力すると、後納運賃・正規料金および各合計を自動計算します。")

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

    tot_2cm = p_cnt_2cm * 204
    reg_2cm = p_cnt_2cm * 310

    tot_3cm = p_cnt_3cm * 280
    reg_3cm = p_cnt_3cm * 360

    all_cnt = p_cnt_1cm + p_cnt_2cm + p_cnt_3cm
    all_tot = tot_1cm + tot_2cm + tot_3cm
    all_reg = reg_1cm + reg_2cm + reg_3cm

    # サマリー表示
    st.markdown("---")
    m1, m2, m3 = st.columns(3)
    m1.metric("合計個数", f"{all_cnt} 通")
    m2.metric("後納運賃 合計", f"{all_tot:,} 円")
    m3.metric("正規料金 合計", f"{all_reg:,} 円")

    # 計算結果一覧表
    df_packet = pd.DataFrame([
        {"区分": "1cm", "運賃(後納)": "173 円", "個数": f"{p_cnt_1cm} 通", "合計(後納)": f"{tot_1cm:,} 円", "正規料金": "250 円", "正規合計": f"{reg_1cm:,} 円"},
        {"区分": "2cm", "運賃(後納)": "204 円", "個数": f"{p_cnt_2cm} 通", "合計(後納)": f"{tot_2cm:,} 円", "正規料金": "310 円", "正規合計": f"{reg_2cm:,} 円"},
        {"区分": "3cm", "運賃(後納)": "280 円", "個数": f"{p_cnt_3cm} 通", "合計(後納)": f"{tot_3cm:,} 円", "正規料金": "360 円", "正規合計": f"{reg_3cm:,} 円"},
        {"区分": "【合計】", "運賃(後納)": "-", "個数": f"{all_cnt} 通", "合計(後納)": f"{all_tot:,} 円", "正規料金": "-", "正規合計": f"{all_reg:,} 円"},
    ])
    st.dataframe(df_packet, use_container_width=True, hide_index=True)

    # ==========================================
    # 送料早見表の参照セクション
    # ==========================================
    st.markdown("---")
    with st.expander("📋 【送料早見表】ゆうパック・レターパック・集荷時間（クリックで開閉）", expanded=True):
        col_yp, col_other = st.columns([3, 2])

        with col_yp:
            st.markdown("#### 📦 ゆうパック運賃（兵庫発 / 1点60円引き）")
            df_youpack = pd.DataFrame([
                {"地域": "兵庫県内", "都道府県": "兵庫", "60サイズ(正規)": "820円", "60サイズ(契約)": "499円", "80サイズ(正規)": "1,130円", "80サイズ(契約)": "688円"},
                {"地域": "近畿・中国・四国・東海・北陸", "都道府県": "大阪 京都 奈良 滋賀 和歌山 / 岡山 広島 鳥取 島根 山口 / 徳島 香川 愛媛 高知 / 静岡 愛知 岐阜 三重 / 富山 石川 福井", "60サイズ(正規)": "880円", "60サイズ(契約)": "536円", "80サイズ(正規)": "1,200円", "80サイズ(契約)": "731円"},
                {"地域": "関東・信越・九州", "都道府県": "東京 神奈川 埼玉 千葉 茨城 栃木 群馬 山梨 / 新潟 長野 / 福岡 佐賀 長崎 熊本 大分 宮崎 鹿児島", "60サイズ(正規)": "990円", "60サイズ(契約)": "603円", "80サイズ(正規)": "1,310円", "80サイズ(契約)": "798円"},
                {"地域": "東北", "都道府県": "青森 岩手 宮城 秋田 山形 福島", "60サイズ(正規)": "1,150円", "60サイズ(契約)": "700円", "80サイズ(正規)": "1,440円", "80サイズ(契約)": "877円"},
                {"地域": "沖縄", "都道府県": "沖縄", "60サイズ(正規)": "1,450円", "60サイズ(契約)": "883円", "80サイズ(正規)": "1,810円", "80サイズ(契約)": "1,236円"},
                {"地域": "北海道", "都道府県": "北海道", "60サイズ(正規)": "1,740円", "60サイズ(契約)": "1,244円", "80サイズ(正規)": "2,040円", "80サイズ(契約)": "1,466円"},
            ])
            st.dataframe(df_youpack, use_container_width=True, hide_index=True)

        with col_other:
            st.markdown("#### 🕒 集荷受付時間")
            st.info("""
            * **前日 18:00まで** ➔ 翌日 **10:00 〜 13:00**
            * **当日 12:00まで** ➔ 当日 **13:00 〜 15:00**
            * **当日 15:00まで** ➔ 当日 **15:00 〜 18:00**
            """)

            st.markdown("#### ✉️ レターパック")
            st.markdown("""
            | 種別 | 料金 |
            | :--- | :---: |
            | レターパックライト | **430 円** |
            | レターパックプラス | **600 円** |
            """)

            st.markdown("#### 📏 ゆうパケット規格")
            st.caption("3辺合計 60cm以内 ／ 長辺 34cm以内 ／ 厚さ 3cm以内 ／ 重量 1kgまで")
