import io
import re
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="股票動態篩選器", layout="wide")
st.title("📈 跨檔案股票動態篩選系統")

# 1. 檔案上傳元件
uploaded_files = st.file_uploader(
    "請上傳篩選所需的 Excel 報表 (可一次多選 5 個檔案)",
    type=["xlsx"],
    accept_multiple_files=True,
)


def load_excel_smart(file_bytes, sheet_name=0):
    """清理 CMoney autofilter 錯誤並自動動態定位『股票代號』所在的標題列"""
    buffer = io.BytesIO()
    with zipfile.ZipFile(file_bytes, "r") as zin:
        with zipfile.ZipFile(buffer, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("xl/worksheets/sheet"):
                    data = re.sub(rb"<autoFilter[^>]*/>", b"", data)
                    data = re.sub(
                        rb"<autoFilter[^>]*>.*?</autoFilter>", b"", data
                    )
                zout.writestr(item, data)
    buffer.seek(0)

    # 讀取完整 raw data，不指定 header
    raw_df = pd.read_excel(buffer, sheet_name=sheet_name, header=None)

    # 動態尋找包含「股票代號」、「股票代碼」或「代號」所在的列作為標題列
    header_idx = 0
    for i, row in raw_df.iloc[:10].iterrows():
        row_str = row.astype(str).values
        if any(
            k in "".join(row_str) for k in ["股票代號", "股票代碼", "代號"]
        ):
            header_idx = i
            break

    # 重新解析資料
    buffer.seek(0)
    df = pd.read_excel(buffer, sheet_name=sheet_name, header=header_idx)

    # 欄位名稱清理
    df.columns = [
        str(c).strip() if pd.notna(c) else f"Unnamed_{i}"
        for i, c in enumerate(df.columns)
    ]

    # 統一將代號欄位轉為 string 並清洗格式
    code_col = next(
        (c for c in df.columns if "股票代號" in c or "股票代碼" in c), None
    )
    if code_col:
        df["標準股票代碼"] = (
            df[code_col]
            .astype(str)
            .str.replace(".0", "", regex=False)
            .str.strip()
        )
        # 過濾掉非股票代碼的雜訊列
        df = df[df["標準股票代碼"].str.contains(r"^\d{4,6}$", na=False)]

    return df


# 2. 側邊欄條件勾選
st.sidebar.header("🎯 篩選條件設定")
use_fundamental = st.sidebar.checkbox("1. 基本面 (營收/毛利率/EPS)", value=True)
use_ma24_safe = st.sidebar.checkbox(
    "2. 安全位階 (收盤價 > MA24月線/24日線)", value=True
)
use_rev_grade = st.sidebar.checkbox("3. 營收分級成長性", value=True)
use_month_high = st.sidebar.checkbox("4. 檢查與上月高點距離", value=True)
use_main_buy = st.sidebar.checkbox("5. 主力買超天數 > 0", value=True)

# 3. 執行按鈕與核心邏輯
if st.button("🚀 開始執行篩選", type="primary"):
    if not uploaded_files:
        st.warning("⚠️ 請先上傳 Excel 檔案！")
    else:
        file_dict = {f.name: f for f in uploaded_files}

        try:
            with st.spinner("正在自動定位標題列與跨檔案資料整合中..."):
                # --- A. 解析『籌碼集中度選股(pressplay).xlsx』---
                f_chip_center = next(
                    (
                        v
                        for k, v in file_dict.items()
                        if "籌碼集中度" in k
                    ),
                    None,
                )

                if not f_chip_center:
                    st.error("❌ 找不到包含《籌碼集中度》字樣的 Excel 檔案！")
                    st.stop()

                df_base = load_excel_smart(
                    f_chip_center, sheet_name="個股圖表資料"
                )

                # 讀取 24MA 均線頁籤
                df_ma = load_excel_smart(f_chip_center, sheet_name="均線")

                # --- B. 解析『型態及籌碼細部篩選(pressplay).xlsx』---
                f_chip_detail = next(
                    (
                        v
                        for k, v in file_dict.items()
                        if "型態及籌碼" in k
                    ),
                    None,
                )
                df_detail = (
                    load_excel_smart(f_chip_detail, sheet_name="近5日")
                    if f_chip_detail
                    else None
                )

                # --- C. 解析『收盤價大於上個月高點.xlsx』---
                f_high = next(
                    (
                        v
                        for k, v in file_dict.items()
                        if "上個月高點" in k
                    ),
                    None,
                )
                df_high = (
                    load_excel_smart(
                        f_high, sheet_name="收盤價大於上個月高點"
                    )
                    if f_high
                    else None
                )

                # --- D. 跨檔案整合 (Merge) ---
                merged_df = df_base.copy()

                # 合併 24日均線
                if df_ma is not None:
                    ma_cols = [
                        c
                        for c in df_ma.columns
                        if "24" in c or "均線" in c or "平均" in c
                    ]
                    if ma_cols:
                        df_ma_sub = df_ma[["標準股票代碼", ma_cols[-1]]].rename(
                            columns={ma_cols[-1]: "MA24均線"}
                        )
                        merged_df = merged_df.merge(
                            df_ma_sub, on="標準股票代碼", how="left"
                        )

                # 合併主力買超天數
                if df_detail is not None:
                    buy_cols = [
                        c for c in df_detail.columns if "主力買超" in c
                    ]
                    if buy_cols:
                        df_detail_sub = df_detail[
                            ["標準股票代碼", buy_cols[0]]
                        ].rename(columns={buy_cols[0]: "主力買超天數"})
                        merged_df = merged_df.merge(
                            df_detail_sub, on="標準股票代碼", how="left"
                        )

                # 合併上月高點距離
                if df_high is not None:
                    dist_cols = [
                        c
                        for c in df_high.columns
                        if "距離" in c or "高點" in c or "收盤價" in c
                    ]
                    if dist_cols:
                        select_cols = ["標準股票代碼"] + dist_cols[:3]
                        merged_df = merged_df.merge(
                            df_high[select_cols], on="標準股票代碼", how="left"
                        )

                # --- E. 條件過濾 ---
                filtered_df = merged_df.copy()

                # 條件 2: 安全位階 (股價 >= 24MA)
                if use_ma24_safe and "MA24均線" in filtered_df.columns:
                    close_col = next(
                        (
                            c
                            for c in filtered_df.columns
                            if "收盤價" in c
                        ),
                        None,
                    )
                    if close_col:
                        filtered_df = filtered_df[
                            filtered_df[close_col] >= filtered_df["MA24均線"]
                        ]

                # 條件 4: 與上月高點距離 (距離 >= -5%)
                if use_month_high:
                    dist_col = next(
                        (
                            c
                            for c in filtered_df.columns
                            if "距離" in c
                        ),
                        None,
                    )
                    if dist_col:
                        filtered_df = filtered_df[
                            filtered_df[dist_col] >= -0.05
                        ]

                # 條件 5: 主力買超天數 > 0
                if use_main_buy and "主力買超天數" in filtered_df.columns:
                    filtered_df = filtered_df[
                        pd.to_numeric(
                            filtered_df["主力買超天數"], errors="coerce"
                        ).fillna(0)
                        > 0
                    ]

                # --- F. 呈現結果與下載 ---
                st.success(
                    f"🎉 篩選完成！共找到 {len(filtered_df)} 檔符合條件的股票"
                )

                # 整理顯示欄位 (保留常用重要欄位)
                display_cols = [
                    c
                    for c in filtered_df.columns
                    if not c.startswith("Unnamed")
                ]
                result_df = filtered_df[display_cols]

                # 網頁表格呈現
                st.dataframe(result_df, use_container_width=True)

                # Excel 下載
                buffer_out = io.BytesIO()
                with pd.ExcelWriter(buffer_out, engine="openpyxl") as writer:
                    result_df.to_excel(writer, index=False, sheet_name="篩選結果")

                st.download_button(
                    label="📥 下載篩選結果 Excel 檔",
                    data=buffer_out.getvalue(),
                    file_name="股票動態篩選結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error(f"❌ 執行過程中發生錯誤：{str(e)}")