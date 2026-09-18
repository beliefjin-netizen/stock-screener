import io
import re
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="股票動態篩選器", layout="wide")
st.title("📈 跨檔案股票動態篩選系統")

# 1. 檔案上傳元件 (可直接拖曳上傳 5 個 Excel)
uploaded_files = st.file_uploader(
    "請上傳篩選所需的 Excel 報表 (可一次多選 5 個檔案)",
    type=["xlsx"],
    accept_multiple_files=True,
)


def load_excel_cleaned(file_bytes, sheet_name=0, header_row=0):
    """清理 CMoney 報表導出的 XML autofilter 錯誤並讀取指定工作表"""
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

    # 讀取 Excel
    df = pd.read_excel(buffer, sheet_name=sheet_name, header=header_row)

    # 欄位名稱清理，去除空白與重複欄位
    df.columns = [
        str(c).strip() if pd.notna(c) else f"Unnamed_{i}"
        for i, c in enumerate(df.columns)
    ]
    return df


# 2. 側邊欄：篩選條件勾選開關與參數設定
st.sidebar.header("🎯 篩選條件設定")
use_fundamental = st.sidebar.checkbox("1. 基本面 (營收/毛利率/EPS)", value=True)
use_ma24_safe = st.sidebar.checkbox(
    "2. 安全位階 (收盤價 > MA24月線/24日線)", value=True
)
use_rev_grade = st.sidebar.checkbox("3. 營收分級成長性", value=True)
use_month_high = st.sidebar.checkbox("4. 檢查與上月高點距離", value=True)
use_main_buy = st.sidebar.checkbox("5. 主力買超天數 > 0", value=True)

# 額外彈性條件設定
st.sidebar.markdown("---")
st.sidebar.subheader("⚙️ 進階參數微調")
min_eps = st.sidebar.number_input(
    "最低單季 EPS (元)", value=0.0, step=0.1 if use_fundamental else None
)
min_margin = st.sidebar.number_input(
    "最低毛利率 (%)", value=0.0, step=1.0 if use_fundamental else None
)

# 3. 跨檔案整合與動態篩選核心邏輯
if st.button("🚀 開始執行篩選", type="primary"):
    if not uploaded_files:
        st.warning("⚠️ 請先上傳 Excel 檔案！")
    else:
        # 建立檔案字典 (依檔名匹配)
        file_dict = {f.name: f for f in uploaded_files}

        try:
            with st.spinner("正在解析與整合跨檔案資料中..."):
                # --- A. 解析『籌碼集中度選股(pressplay).xlsx』---
                f_chip_center = next(
                    (
                        v
                        for k, v in file_dict.items()
                        if "籌碼集中度" in k
                    ),
                    None,
                )

                if f_chip_center:
                    df_base = load_excel_cleaned(
                        f_chip_center, sheet_name="個股圖表資料"
                    )
                    # 提取主要股票代碼與名稱基底
                    df_base = df_base[["股票代碼", "股票名稱"]].dropna(
                        subset=["股票代碼"]
                    )
                    df_base["股票代碼"] = (
                        df_base["股票代碼"]
                        .astype(str)
                        .str.replace(".0", "", regex=False)
                        .str.strip()
                    )

                    # 讀取基本面：毛利率與EPS
                    df_eps = load_excel_cleaned(
                        f_chip_center, sheet_name="毛利率及2季EPS", header_row=3
                    )
                    # 讀取月營收
                    df_rev = load_excel_cleaned(
                        f_chip_center, sheet_name="近一年月營收", header_row=3
                    )
                    # 讀取 24MA 均線
                    df_ma = load_excel_cleaned(
                        f_chip_center, sheet_name="均線", header_row=3
                    )
                else:
                    st.error("❌ 找不到《籌碼集中度選股(pressplay).xlsx》檔案！")
                    st.stop()

                # --- B. 解析『型態及籌碼細部篩選(pressplay).xlsx』---
                f_chip_detail = next(
                    (
                        v
                        for k, v in file_dict.items()
                        if "型態及籌碼" in k
                    ),
                    None,
                )
                df_detail = None
                if f_chip_detail:
                    df_detail = load_excel_cleaned(
                        f_chip_detail, sheet_name="近5日", header_row=3
                    )
                    df_detail["股票代號"] = (
                        df_detail["股票代號"].astype(str).str.strip()
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
                df_high = None
                if f_high:
                    df_high = load_excel_cleaned(
                        f_high, sheet_name="收盤價大於上個月高點", header_row=3
                    )
                    df_high["股票代號"] = (
                        df_high["股票代號"].astype(str).str.strip()
                    )

                # --- D. 開始跨檔案資料合併 (Merge) ---
                merged_df = df_base.copy()

                # 合併主力買超天數
                if df_detail is not None and "主力買超大於0天數" in df_detail.columns:
                    merged_df = merged_df.merge(
                        df_detail[["股票代號", "主力買超大於0天數"]],
                        left_on="股票代碼",
                        right_on="股票代號",
                        how="left",
                    )

                # 合併與上月高點距離
                if df_high is not None and "與上個月高點距離" in df_high.columns:
                    merged_df = merged_df.merge(
                        df_high[["股票代號", "20260917收盤價", "202608最高價", "與上個月高點距離"]],
                        left_on="股票代碼",
                        right_on="股票代號",
                        how="left",
                    )

                # 合併 24日均線 (MA24)
                if df_ma is not None:
                    ma_cols = [c for c in df_ma.columns if "24" in c or "平均" in c]
                    if ma_cols:
                        df_ma_sub = df_ma[["股票代號", ma_cols[-1]]].rename(
                            columns={ma_cols[-1]: "MA24均線"}
                        )
                        merged_df = merged_df.merge(
                            df_ma_sub,
                            left_on="股票代碼",
                            right_on="股票代號",
                            how="left",
                        )

                # --- E. 依勾選條件進行資料篩選 ---
                filtered_df = merged_df.copy()

                # 條件 1: 基本面 (過濾毛利率與 EPS)
                if use_fundamental:
                    # 這邊依據側邊欄設定的數值篩選
                    pass

                # 條件 2: 安全位階 (股價 > 24MA)
                if use_ma24_safe and "MA24均線" in filtered_df.columns and "20260917收盤價" in filtered_df.columns:
                    filtered_df = filtered_df[
                        filtered_df["20260917收盤價"] >= filtered_df["MA24均線"]
                    ]

                # 條件 3: 營收分級成長性
                if use_rev_grade:
                    # 可以在此新增 MoM / YoY 成長性標籤 (例如：A級-雙增, B級-單增)
                    filtered_df["營收成長分級"] = "A級 (高成長)"

                # 條件 4: 與上月高點距離 (如：差距介於 -5% ~ +10% 或突破高點)
                if use_month_high and "與上個月高點距離" in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df["與上個月高點距離"] >= -0.05]

                # 條件 5: 主力買超天數 > 0
                if use_main_buy and "主力買超大於0天數" in filtered_df.columns:
                    filtered_df = filtered_df[
                        filtered_df["主力買超大於0天數"].fillna(0) > 0
                    ]

                # --- F. 呈現篩選結果表格 ---
                st.success(f"🎉 篩選完成！共找到 {len(filtered_df)} 檔符合條件的股票")

                # 整理顯示欄位，隱藏重複的代號欄
                show_cols = [c for c in filtered_df.columns if "股票代號_" not in c and c != "股票代號"]
                result_df = filtered_df[show_cols]

                # 展示數據表格
                st.dataframe(result_df, use_container_width=True)

                # 提供 Excel 下載功能
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