import streamlit as st
import pandas as pd
import io
import re
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
    st.caption("塊ごとの個数を「12+9+14」や「12 9 14」のように入力すると自動合算して計算します。")

    def parse_count_input(text_val):
        """12+9+14 や 12, 9, 14 などの文字列から数値を抽出して合算"""
        if not text_val or not text_val.strip():
            return 0
        numbers = re.findall(r"\d+", text_val)
        return sum(int(n) for n in numbers) if numbers else 0

    col1, col2, col3 = st.columns(3)
    with col1:
        raw_1cm = st.text_input("1cm 個数加算（例: 12+9+14）", value="", placeholder="12+9+14", key="p1")
        p_cnt_1cm = parse_count_input(raw_1cm)
        st.caption(f"➔ 合計: **{p_cnt_1cm}** 通 (後納173円 / 正規250円)")

    with col2:
        raw_2cm = st.text_input("2cm 個数加算（例: 5+4）", value="", placeholder="5+4", key="p2")
        p_cnt_2cm = parse_count_input(raw_2cm)
        st.caption(f"➔ 合計: **{p_cnt_2cm}** 通 (後納204円 / 正規310円)")

    with col3:
        raw_3cm = st.text_input("3cm 個数加算（例: 2+1）", value="", placeholder="2+1", key="p3")
        p_cnt_3cm = parse_count_input(raw_3cm)
        st.caption(f"➔ 合計: **{p_cnt_3cm}** 通 (後納280円 / 正規360円)")

    # 金額計算
    tot_1cm = p_cnt_1cm * 173
    tot_2cm = p_cnt_2cm * 204
    tot_3cm = p_cnt_3cm * 280

    all_cnt = p_cnt_1cm + p_cnt_2cm + p_cnt_3cm
    all_tot = tot_1cm + tot_2cm + tot_3cm

    # サマリー表示
    st.markdown("---")
    m1, m2 = st.columns(2)
    m1.metric("総個数", f"{all_cnt} 通")
    m2.metric("後納運賃 合計", f"{all_tot:,} 円")

    # 一覧表（正規合計は除外、正規単価のみ表示）
    df_packet = pd.DataFrame([
        {"区分": "1cm", "運賃(後納)": "173 円", "正規料金": "250 円", "個数": f"{p_cnt_1cm} 通", "合計(後納)": f"{tot_1cm:,} 円"},
        {"区分": "2cm", "運賃(後納)": "204 円", "正規料金": "310 円", "個数": f"{p_cnt_2cm} 通", "合計(後納)": f"{tot_2cm:,} 円"},
        {"区分": "3cm", "運賃(後納)": "280 円", "正規料金": "360 円", "個数": f"{p_cnt_3cm} 通", "合計(後納)": f"{tot_3cm:,} 円"},
        {"区分": "【合計】", "運賃(後納)": "-", "正規料金": "-", "個数": f"{all_cnt} 通", "合計(後納)": f"{all_tot:,} 円"},
    ])
    st.table(df_packet)

    # ==========================================
    # 送料早見表セクション（ウィンドウ100%にしなくても全文が見える全幅テーブル）
    # ==========================================
    st.markdown("---")
    st.subheader("📋 送料早見表")

    # ゆうパック運賃表（全幅HTMLレスポンシブデザイン）
    youpack_html = """
    <div style="width:100%; overflow-x:auto; margin-bottom:20px;">
        <table style="width:100%; border-collapse:collapse; font-size:13px; text-align:center; background:#fff;">
            <thead>
                <tr style="background:#f1f3f5; border-bottom:2px solid #ccc;">
                    <th style="padding:8px 6px; border:1px solid #ddd; width:16%;">地域</th>
                    <th style="padding:8px 6px; border:1px solid #ddd; width:44%;">対象都道府県</th>
                    <th style="padding:8px 6px; border:1px solid #ddd; width:10%;">60(正規)</th>
                    <th style="padding:8px 6px; border:1px solid #ddd; width:10%; background:#e8f4fd;">60(契約)</th>
                    <th style="padding:8px 6px; border:1px solid #ddd; width:10%;">80(正規)</th>
                    <th style="padding:8px 6px; border:1px solid #ddd; width:10%; background:#e8f4fd;">80(契約)</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold;">兵庫県内</td>
                    <td style="padding:6px; border:1px solid #ddd; text-align:left;">兵庫</td>
                    <td style="padding:6px; border:1px solid #ddd;">820円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">499円</td>
                    <td style="padding:6px; border:1px solid #ddd;">1,130円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">688円</td>
                </tr>
                <tr>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold;">近畿・中国・四国<br>東海・北陸</td>
                    <td style="padding:6px; border:1px solid #ddd; text-align:left; font-size:12px;">大阪 京都 奈良 滋賀 和歌山 / 岡山 広島 鳥取 島根 山口<br>徳島 香川 愛媛 高知 / 静岡 愛知 岐阜 三重 / 富山 石川 福井</td>
                    <td style="padding:6px; border:1px solid #ddd;">880円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">536円</td>
                    <td style="padding:6px; border:1px solid #ddd;">1,200円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">731円</td>
                </tr>
                <tr>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold;">関東・信越・九州</td>
                    <td style="padding:6px; border:1px solid #ddd; text-align:left; font-size:12px;">東京 神奈川 埼玉 千葉 茨城 栃木 群馬 山梨 / 新潟 長野<br>福岡 佐賀 長崎 熊本 大分 宮崎 鹿児島</td>
                    <td style="padding:6px; border:1px solid #ddd;">990円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">603円</td>
                    <td style="padding:6px; border:1px solid #ddd;">1,310円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">798円</td>
                </tr>
                <tr>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold;">東北</td>
                    <td style="padding:6px; border:1px solid #ddd; text-align:left;">青森 岩手 宮城 秋田 山形 福島</td>
                    <td style="padding:6px; border:1px solid #ddd;">1,150円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">700円</td>
                    <td style="padding:6px; border:1px solid #ddd;">1,440円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">877円</td>
                </tr>
                <tr>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold;">沖縄</td>
                    <td style="padding:6px; border:1px solid #ddd; text-align:left;">沖縄</td>
                    <td style="padding:6px; border:1px solid #ddd;">1,450円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">883円</td>
                    <td style="padding:6px; border:1px solid #ddd;">1,810円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">1,236円</td>
                </tr>
                <tr>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold;">北海道</td>
                    <td style="padding:6px; border:1px solid #ddd; text-align:left;">北海道</td>
                    <td style="padding:6px; border:1px solid #ddd;">1,740円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">1,244円</td>
                    <td style="padding:6px; border:1px solid #ddd;">2,040円</td>
                    <td style="padding:6px; border:1px solid #ddd; font-weight:bold; color:#0056b3; background:#f8fbfe;">1,466円</td>
                </tr>
            </tbody>
        </table>
        <div style="font-size:12px; color:#666; margin-top:4px;">※ゆうパック：兵庫発（契約運賃は1点あたり60円引き適用済み）</div>
    </div>
    """
    st.markdown(youpack_html, unsafe_allow_html=True)

    # その他規格
    c_lp, c_std = st.columns([1, 1])
    with c_lp:
        st.markdown("#### ✉️ レターパック")
        st.markdown("""
        | 種別 | 料金 |
        | :--- | :---: |
        | レターパックライト | **430 円** |
        | レターパックプラス | **600 円** |
        """)
    with c_std:
        st.markdown("#### 📏 ゆうパケット規格")
        st.info("3辺合計 60cm以内 ／ 長辺 34cm以内 ／ 厚さ 3cm以内 ／ 重量 1kgまで")
