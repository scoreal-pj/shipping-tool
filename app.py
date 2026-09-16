import pandas as pd
import streamlit as st
import io
import re

# ページの基本設定
st.set_page_config(
    page_title="出荷ファイル自動整形Webシステム",
    layout="centered",
    initial_sidebar_state="expanded"
)

st.title("📦 出荷ファイル自動整形Webシステム")
st.markdown("ブラウザ上で元データを読み込み、チェックボックスで担当者を選ぶだけで、不要な行の非表示・整形を自動で行います。")

# 1. ファイルのアップロード
st.header("1. 元データのアップロード")
uploaded_file = st.file_uploader("メルカリ等の売却済みデータ CSV (例: orders_...csv)", type=["csv"])

if uploaded_file is not None:
    # 文字コードの自動判定（Shift_JISまたはUTF-8）
    try:
        df = pd.read_csv(uploaded_file, encoding="cp932")
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, encoding="utf-8")

    st.success(f"ファイルを正常に読み込みました（全 {len(df)} 件）")

    # 管理番号やSKU、担当者コードが入っている列を自動検出
    target_col = None
    for col in ["original_product_id", "品名２", "品名１", "お客様管理番号"]:
        if col in df.columns:
            target_col = col
            break

    if target_col:
        st.info(f"データ内の識別カラム： **{target_col}**")

        # データ内から担当者コード（_sy, _mb, _ami, _kj など）を自動で洗い出す
        all_texts = df[target_col].dropna().astype(str).tolist()
        found_managers = set()
        for t in all_texts:
            match = re.search(r'_([a-zA-Z]+)\d*$', t)
            if match:
                found_managers.add(match.group(1))
        
        if not found_managers:
            found_managers = {"sy", "mb", "ami", "kj"}

        # 2. 担当者の選択（チェックボックス）
        st.header("2. 残す担当者の選択")
        st.markdown("チェックを入れた担当者のデータのみを残し、他を自動で非表示（除外）します。")
        
        selected_managers = []
        cols = st.columns(len(found_managers) if len(found_managers) > 0 else 4)
        for i, mgr in enumerate(sorted(list(found_managers))):
            # デフォルトで 'sy' にチェックを入れる
            default_val = True if mgr.lower() == "sy" else False
            with cols[i % len(cols)]:
                if st.checkbox(f"担当: {mgr}", value=default_val, key=f"mgr_{mgr}"):
                    selected_managers.append(mgr)

        # 3. 実行ボタン
        st.header("3. 処理の実行")
        if st.button("✨ 選択した担当者以外を非表示・整形する", type="primary"):
            if not selected_managers:
                st.warning("少なくとも1つの担当者を選択してください。")
            else:
                # 選択された担当者にマッチする行を抽出
                pattern = '|'.join([f"_{mgr}" for mgr in selected_managers])
                mask = df[target_col].astype(str).str.contains(pattern, case=False, na=False)
                
                df_filtered = df[mask].copy()
                excluded_count = len(df) - len(df_filtered)

                st.write(f"📊 処理結果: 残ったデータ **{len(df_filtered)} 件** （非表示・除外したデータ {excluded_count} 件）")

                # プレビュー表示
                st.subheader("📋 処理後データのプレビュー（一部）")
                st.dataframe(df_filtered.head(10))

                # ダウンロード用ファイルの生成（Excelで開けるShift_JIS形式）
                csv_buffer = io.BytesIO()
                df_filtered.to_csv(csv_buffer, index=False, encoding="cp932", errors="replace")
                csv_bytes = csv_buffer.getvalue()

                st.download_button(
                    label="📥 整形済みCSVファイルをダウンロード",
                    data=csv_bytes,
                    file_name="shipped_formatted.csv",
                    mime="text/csv",
                )
    else:
        st.error("ファイル内に管理番号や担当者コードの列が見つかりませんでした。")