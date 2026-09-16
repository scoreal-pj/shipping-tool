import streamlit as st
import pandas as pd
import io
import re
import openpyxl
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="業務支援システム", layout="wide")

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

    def parse_textarea_numbers(text_val):
        if not text_val:
            return 0, []
        nums = [int(n) for n in re.findall(r"\d+", text_val)]
        return sum(nums), nums

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("### 1cm")
        txt_1cm = st.text_area("個数入力", height=120, placeholder="例:\n12\n9\n14", key="area_1cm", label_visibility="collapsed")
        sum_1cm, _ = parse_textarea_numbers(txt_1cm)
        st.markdown(f"<div style='font-size:18px; font-weight:bold; color:#0056b3;'>合計: {sum_1cm} 通</div>", unsafe_allow_html=True)

    with col2:
        st.markdown("### 2cm")
        txt_2cm = st.text_area("個数入力", height=120, placeholder="例:\n5\n4", key="area_2cm", label_visibility="collapsed")
        sum_2cm, _ = parse_textarea_numbers(txt_2cm)
        st.markdown(f"<div style='font-size:18px; font-weight:bold; color:#0056b3;'>合計: {sum_2cm} 通</div>", unsafe_allow_html=True)

    with col3:
        st.markdown("### 3cm")
        txt_3cm = st.text_area("個数入力", height=120, placeholder="例:\n2\n1", key="area_3cm", label_visibility="collapsed")
        sum_3cm, _ = parse_textarea_numbers(txt_3cm)
        st.markdown(f"<div style='font-size:18px; font-weight:bold; color:#0056b3;'>合計: {sum_3cm} 通</div>", unsafe_allow_html=True)

    # 金額計算
    tot_1cm = sum_1cm * 173
    tot_2cm = sum_2cm * 204
    tot_3cm = sum_3cm * 280

    all_cnt = sum_1cm + sum_2cm + sum_3cm
    all_tot = tot_1cm + tot_2cm + tot_3cm

    # ----------------------------------------------------
    # ゆうパケット一覧表（最下行の合計で集計確認）
    # ----------------------------------------------------
    st.markdown("---")
    table_packet_html = f"""
    <div style="width:100%; margin-bottom:25px;">
        <table style="width:100%; border-collapse:collapse; font-size:14px; text-align:center; background:#fff;">
            <thead>
                <tr style="background:#f4f5f7; border-bottom:2px solid #ccc;">
                    <th style="padding:10px 8px; border:1px solid #ddd; width:15%;">区分</th>
                    <th style="padding:10px 8px; border:1px solid #ddd; width:15%; color:#666; font-size:12.5px;">運賃(後納)</th>
                    <th style="padding:10px 8px; border:1px solid #0066cc; width:25%; background:#e8f4fd; color:#004085; font-size:16px; font-weight:bold;">個数</th>
                    <th style="padding:10px 8px; border:1px solid #e67e22; width:30%; background:#fef5ea; color:#b94a00; font-size:16px; font-weight:bold;">合計（後納）</th>
                    <th style="padding:10px 8px; border:1px solid #ddd; width:15%; color:#666; font-size:12.5px;">正規料金</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td style="padding:10px 8px; border:1px solid #ddd; font-weight:bold; font-size:16px;">1cm</td>
                    <td style="padding:10px 8px; border:1px solid #ddd; color:#555;">173 円</td>
                    <td style="padding:10px 8px; border:1px solid #0066cc; background:#f4f9fe; font-size:19px; font-weight:900; color:#0056b3;">{sum_1cm:,} 通</td>
                    <td style="padding:10px 8px; border:1px solid #e67e22; background:#fffcf6; font-size:19px; font-weight:900; color:#c0392b;">{tot_1cm:,} 円</td>
                    <td style="padding:10px 8px; border:1px solid #ddd; color:#777;">250 円</td>
                </tr>
                <tr>
                    <td style="padding:10px 8px; border:1px solid #ddd; font-weight:bold; font-size:16px;">2cm</td>
                    <td style="padding:10px 8px; border:1px solid #ddd; color:#555;">204 円</td>
                    <td style="padding:10px 8px; border:1px solid #0066cc; background:#f4f9fe; font-size:19px; font-weight:900; color:#0056b3;">{sum_2cm:,} 通</td>
                    <td style="padding:10px 8px; border:1px solid #e67e22; background:#fffcf6; font-size:19px; font-weight:900; color:#c0392b;">{tot_2cm:,} 円</td>
                    <td style="padding:10px 8px; border:1px solid #ddd; color:#777;">310 円</td>
                </tr>
                <tr>
                    <td style="padding:10px 8px; border:1px solid #ddd; font-weight:bold; font-size:16px;">3cm</td>
                    <td style="padding:10px 8px; border:1px solid #ddd; color:#555;">280 円</td>
                    <td style="padding:10px 8px; border:1px solid #0066cc; background:#f4f9fe; font-size:19px; font-weight:900; color:#0056b3;">{sum_3cm:,} 通</td>
                    <td style="padding:10px 8px; border:1px solid #e67e22; background:#fffcf6; font-size:19px; font-weight:900; color:#c0392b;">{tot_3cm:,} 円</td>
                    <td style="padding:10px 8px; border:1px solid #ddd; color:#777;">360 円</td>
                </tr>
                <tr style="background:#f9fafb; border-top:3px solid #666;">
                    <td style="padding:12px 8px; border:1px solid #ccc; font-weight:900; font-size:17px;">【合計】</td>
                    <td style="padding:12px 8px; border:1px solid #ccc; color:#888;">-</td>
                    <td style="padding:12px 8px; border:2px solid #0066cc; background:#dbeafe; font-size:22px; font-weight:900; color:#003366;">{all_cnt:,} 通</td>
                    <td style="padding:12px 8px; border:2px solid #e67e22; background:#fde8d0; font-size:22px; font-weight:900; color:#962d00;">{all_tot:,} 円</td>
                    <td style="padding:12px 8px; border:1px solid #ccc; color:#888;">-</td>
                </tr>
            </tbody>
        </table>
    </div>
    """
    st.markdown(table_packet_html, unsafe_allow_html=True)

    # ==========================================
    # ゆうパック送料早見表（正規料金を最強調配置）
    # ==========================================
    st.markdown("---")
    st.subheader("📋 ゆうパック 送料早見表（兵庫発）")
    st.caption("客対応時は **太字の「正規料金」** を案内。右側の契約・差額は自社確認用です。")

    youpack_html = """
    <div style="width:100%; overflow-x:auto; margin-bottom:20px;">
        <table style="width:100%; border-collapse:collapse; font-size:13px; text-align:center; background:#fff;">
            <thead>
                <tr style="background:#2c3e50; color:#fff;">
                    <th rowspan="2" style="padding:8px 4px; border:1px solid #455a64; width:13%;">地域</th>
                    <th rowspan="2" style="padding:8px 6px; border:1px solid #455a64; width:33%;">対象都道府県</th>
                    <th colspan="4" style="padding:6px; border:1px solid #455a64; background:#0056b3; font-size:14px;">【お客様提示用】 正規料金</th>
                    <th colspan="2" style="padding:6px; border:1px solid #455a64; background:#d35400; font-size:12.5px;">60サイズ自社差額</th>
                    <th colspan="2" style="padding:6px; border:1px solid #455a64; background:#d35400; font-size:12.5px;">80サイズ自社差額</th>
                </tr>
                <tr style="background:#e8f4fd; color:#003366; font-weight:bold; border-bottom:2px solid #999;">
                    <th style="padding:8px 4px; border:1px solid #bcd8f5; width:8%; font-size:13px;">60サイズ</th>
                    <th style="padding:8px 4px; border:1px solid #bcd8f5; width:8%; font-size:13px;">80サイズ</th>
                    <th style="padding:8px 4px; border:1px solid #bcd8f5; width:9%; font-size:13px;">100サイズ</th>
                    <th style="padding:8px 4px; border:1px solid #bcd8f5; width:9%; font-size:13px;">120サイズ</th>
                    <th style="padding:6px 2px; border:1px solid #ddd; width:6.5%; background:#fff3e0; color:#b94a00; font-size:11.5px;">契約運賃</th>
                    <th style="padding:6px 2px; border:1px solid #ddd; width:6.5%; background:#ffe0b2; color:#d35400; font-size:11.5px;">差額</th>
                    <th style="padding:6px 2px; border:1px solid #ddd; width:6.5%; background:#fff3e0; color:#b94a00; font-size:11.5px;">契約運賃</th>
                    <th style="padding:6px 2px; border:1px solid #ddd; width:6.5%; background:#ffe0b2; color:#d35400; font-size:11.5px;">差額</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td style="padding:8px 4px; border:1px solid #ddd; font-weight:bold; font-size:14px; background:#f9f9f9;">兵庫県内</td>
                    <td style="padding:8px 6px; border:1px solid #ddd; text-align:left;">兵庫</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">820円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,130円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,450円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,770円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">499円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+321円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">688円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+442円</td>
                </tr>
                <tr>
                    <td style="padding:8px 4px; border:1px solid #ddd; font-weight:bold; font-size:13.5px; background:#f9f9f9;">北陸・東海<br>近畿・中国・四国</td>
                    <td style="padding:8px 6px; border:1px solid #ddd; text-align:left; font-size:11.5px; line-height:1.4;">富山 石川 福井 / 静岡 愛知 岐阜 三重<br>大阪 京都 奈良 滋賀 和歌山 / 岡山 広島 鳥取 島根 山口<br>徳島 香川 愛媛 高知</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">880円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,200円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,500円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,830円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">536円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+344円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">731円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+469円</td>
                </tr>
                <tr>
                    <td style="padding:8px 4px; border:1px solid #ddd; font-weight:bold; font-size:14px; background:#f9f9f9;">関東・信越</td>
                    <td style="padding:8px 6px; border:1px solid #ddd; text-align:left; font-size:11.5px;">茨城 栃木 群馬 埼玉 千葉 東京 神奈川 山梨 / 新潟 長野</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">990円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,310円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,620円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,940円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">603円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+387円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">798円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+512円</td>
                </tr>
                <tr>
                    <td style="padding:8px 4px; border:1px solid #ddd; font-weight:bold; font-size:14px; background:#f9f9f9;">九州</td>
                    <td style="padding:8px 6px; border:1px solid #ddd; text-align:left; font-size:11.5px;">福岡 佐賀 長崎 熊本 大分 宮崎 鹿児島</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">990円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,310円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,620円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,940円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">603円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+387円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">798円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+512円</td>
                </tr>
                <tr>
                    <td style="padding:8px 4px; border:1px solid #ddd; font-weight:bold; font-size:14px; background:#f9f9f9;">東北</td>
                    <td style="padding:8px 6px; border:1px solid #ddd; text-align:left; font-size:11.5px;">青森 岩手 宮城 秋田 山形 福島</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,150円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,440円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,780円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">2,080円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">700円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+450円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">877円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+563円</td>
                </tr>
                <tr>
                    <td style="padding:8px 4px; border:1px solid #ddd; font-weight:bold; font-size:14px; background:#f9f9f9;">沖縄</td>
                    <td style="padding:8px 6px; border:1px solid #ddd; text-align:left; font-size:11.5px;">沖縄</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,450円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,810円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">2,160円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">2,490円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">883円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+567円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">1,236円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+574円</td>
                </tr>
                <tr>
                    <td style="padding:8px 4px; border:1px solid #ddd; font-weight:bold; font-size:14px; background:#f9f9f9;">北海道</td>
                    <td style="padding:8px 6px; border:1px solid #ddd; text-align:left; font-size:11.5px;">北海道</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">1,740円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:16px; font-weight:900; color:#0056b3; background:#f4f9fe;">2,040円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">2,350円</td>
                    <td style="padding:8px 4px; border:1px solid #bcd8f5; font-size:15px; font-weight:900; color:#0056b3; background:#f4f9fe;">2,650円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">1,244円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+496円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; color:#555; background:#fffcf6;">1,466円</td>
                    <td style="padding:8px 2px; border:1px solid #ddd; font-weight:bold; color:#27ae60; background:#f6fbf7;">+574円</td>
                </tr>
            </tbody>
        </table>
    </div>
    """
    st.markdown(youpack_html, unsafe_allow_html=True)

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
