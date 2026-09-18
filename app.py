import io
import re
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="股票策略動態篩選器", layout="wide")
st.title("📈 策略選股：基本面 + 均線位階 + 臨界點突破 + 主力籌碼")

# 1. 檔案上傳元件
uploaded_files = st.file_uploader(
    "請上傳篩選所需的 4 份 Excel 報表 (支援多選)",
    type=["xlsx"],
    accept_multiple_files=True,
    help="請確保包含：籌碼集中度、量價均線法人、收盤價大於上月高點、型態及籌碼細部 等檔案"
)

def clean_code_series(series):
    """終極股票代碼清洗：解決 0050 變 50、帶小數點、前後空白等問題"""
    s = series.astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    s = s.apply(lambda x: x.zfill(4) if x.isdigit() and len(x) < 4 else x)
    return s.str.extract(r"(\d{4,6})")[0]

def clean_number_series(series):
    """將包含 %, 逗號, -- 的字串安全轉換為純浮點數"""
    if series is None or series.empty:
        return pd.Series(dtype=float)
    s_clean = (
        series.astype(str)
        .str.replace(r"[%,\s]", "", regex=True)
        .str.replace("--", "NaN", regex=False)
        .str.replace("NA", "NaN", regex=False)
    )
    return pd.to_numeric(s_clean, errors="coerce")

def load_excel_smart(file_bytes, kw_sheet_list=None):
    """智慧讀取：透過評分機制精準定位真實資料表頭"""
    if not file_bytes:
        return pd.DataFrame()
        
    buffer = io.BytesIO()
    with zipfile.ZipFile(file_bytes, "r") as zin:
        with zipfile.ZipFile(buffer, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("xl/worksheets/sheet"):
                    data = re.sub(rb"<autoFilter[^>]*/>", b"", data)
                    data = re.sub(rb"<autoFilter[^>]*>.*?</autoFilter>", b"", data)
                zout.writestr(item, data)
    
    buffer.seek(0)
    xls = pd.ExcelFile(buffer)
    
    # 尋找最符合關鍵字的頁籤
    target_sheet = xls.sheet_names[0]
    if kw_sheet_list:
        for kw in kw_sheet_list:
            for s in xls.sheet_names:
                if kw in s:
                    target_sheet = s
                    break

    buffer.seek(0)
    raw_df = pd.read_excel(buffer, sheet_name=target_sheet, header=None)

    # 尋找真實標題列
    best_idx = 0
    max_score = 0
    for i, row in raw_df.iloc[:50].iterrows():
        cells = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != ""]
        row_str = "".join(cells)
        
        if any(kw in row_str for kw in ["報表名稱", "篩選條件", "產出時間", "條件描述"]):
            continue
            
        score = sum(1 for kw in ["股票", "代碼", "代號", "名稱", "收盤", "均線", "營收", "距離", "主力"] if any(kw in c for c in cells))
        if score > max_score and len(cells) >= 3:
            max_score = score
            best_idx = i

    buffer.seek(0)
    df = pd.read_excel(buffer, sheet_name=target_sheet, header=best_idx)
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() if pd.notna(c) else f"Unnamed_{i}" for i, c in enumerate(df.columns)]

    # 標準化股票代碼
    code_col = next((c for c in df.columns if "股票代號" in c or "股票代碼" in c or c == "代號"), None)
    if code_col:
        df["標準股票代碼"] = clean_code_series(df[code_col])
        df = df[df["標準股票代碼"].notna()]
    else:
        df["標準股票代碼"] = pd.Series(dtype=str)

    return df

# 2. 側邊欄條件設定
st.sidebar.header("🎯 篩選條件設定")
use_fundamental = st.sidebar.checkbox("1. 基本面 (營收YoY>0%, 毛利率季增, EPS>0)", value=True)
use_ma24_bias = st.sidebar.checkbox("2. 安全位階 (收盤價在 24日均線 的 -5% ~ +15% 之間)", value=True)
use_month_high_dist = st.sidebar.checkbox("3. 關鍵突破位階 (與上個月高點距離 -3% ~ +3%)", value=True)
use_main_buy = st.sidebar.checkbox("4. 籌碼面 (主力買超大於0天數 > 0)", value=True)

# 3. 執行邏輯
if st.button("🚀 開始執行策略篩選", type="primary"):
    if not uploaded_files:
        st.warning("⚠️ 請先上傳 Excel 報表！")
    else:
        file_dict = {f.name: f for f in uploaded_files}

        try:
            with st.spinner("依據指定欄位進行資料讀取與整合中..."):
                
                # 依據您提供的檔名特徵，精準分配檔案
                f_fund, f_tech, f_high, f_chip = None, None, None, None
                for name, f in file_dict.items():
                    if "籌碼集中度" in name: f_fund = f
                    elif "量價均線" in name: f_tech = f
                    elif "高點" in name or "上月" in name: f_high = f
                    elif "型態" in name or "細部" in name: f_chip = f
                
                # --- A. 載入各表 (依照您指定的頁籤與檔案) ---
                # 1. 基本面：籌碼集中度選股 -> 頁籤「營收年成長.毛利率及2季EPS」
                df_fund = load_excel_smart(f_fund, ["營收年成長", "毛利率及", "EPS"]) if f_fund else pd.DataFrame()
                
                # 2. 安全位階：量價均線法人選股
                df_tech = load_excel_smart(f_tech) if f_tech else pd.DataFrame()
                
                # 3. 關鍵突破位階：收盤價大於上個月高點
                df_high = load_excel_smart(f_high) if f_high else pd.DataFrame()
                
                # 4. 籌碼面：型態及籌碼細部篩選
                df_chip = load_excel_smart(f_chip) if f_chip else pd.DataFrame()

                # --- B. 進行跨表合併 (Left Join，以技術面檔案為基底) ---
                # 因為安全位階有收盤價，我們以 df_tech 為基底較為保險
                if df_tech.empty and not df_fund.empty:
                    merged_df = df_fund.copy()
                else:
                    merged_df = df_tech.copy()

                # 合併基本面
                if not df_fund.empty and "標準股票代碼" in df_fund.columns and f_tech:
                    df_fund_sub = df_fund.drop(columns=[c for c in df_fund.columns if c in merged_df.columns and c != "標準股票代碼"])
                    merged_df = merged_df.merge(df_fund_sub, on="標準股票代碼", how="left")

                # 合併上月高點距離 (指定擷取「與上個月高點距離」)
                if not df_high.empty and "標準股票代碼" in df_high.columns:
                    dist_cols = [c for c in df_high.columns if "高點距離" in c or "與上個月高點距離" in c]
                    if dist_cols:
                        df_high_sub = df_high[["標準股票代碼", dist_cols[0]]].drop_duplicates(subset=["標準股票代碼"]).rename(columns={dist_cols[0]: "與上個月高點距離"})
                        merged_df = merged_df.merge(df_high_sub, on="標準股票代碼", how="left")

                # 合併主力買超 (指定擷取「主力買超大於0天數」)
                if not df_chip.empty and "標準股票代碼" in df_chip.columns:
                    buy_cols = [c for c in df_chip.columns if "主力買超" in c or "大於0天數" in c]
                    if buy_cols:
                        df_chip_sub = df_chip[["標準股票代碼", buy_cols[0]]].drop_duplicates(subset=["標準股票代碼"]).rename(columns={buy_cols[0]: "主力買超大於0天數"})
                        merged_df = merged_df.merge(df_chip_sub, on="標準股票代碼", how="left")

                # 監控面板
                with st.expander("📊 各檔案載入與對應狀態 (除錯用)", expanded=True):
                    st.markdown(f"""
                    - **基本面檔 (籌碼集中度)**: 載入 `{len(df_fund)}` 筆
                    - **技術面檔 (量價均線)**: 載入 `{len(df_tech)}` 筆 
                    - **高點突破檔 (上個月高點)**: 載入 `{len(df_high)}` 筆
                    - **籌碼面檔 (型態及籌碼)**: 載入 `{len(df_chip)}` 筆
                    - **資料合併後總庫**: 基準庫維持 `{len(merged_df)}` 筆
                    """)

                # --- C. 條件篩選 ---
                filtered_df = merged_df.copy()

                # 1. 基本面
                if use_fundamental:
                    rev_col = next((c for c in filtered_df.columns if "營收年成長" in c or "營收年增" in c), None)
                    eps_col = next((c for c in filtered_df.columns if "EPS" in c or "每股盈餘" in c), None)
                    gm_curr_col = next((c for c in filtered_df.columns if "毛利率" in c and "前" not in c), None)
                    gm_prev_col = next((c for c in filtered_df.columns if "前一季毛利率" in c or "前季毛利率" in c), None)

                    if rev_col: filtered_df = filtered_df[clean_number_series(filtered_df[rev_col]) > 0]
                    if eps_col: filtered_df = filtered_df[clean_number_series(filtered_df[eps_col]) > 0]
                    if gm_curr_col and gm_prev_col:
                        filtered_df = filtered_df[clean_number_series(filtered_df[gm_curr_col]) > clean_number_series(filtered_df[gm_prev_col])]

                # 2. 安全位階 (收盤價、24日均線)
                if use_ma24_bias:
                    close_col = next((c for c in filtered_df.columns if "收盤價" in c), None)
                    ma24_col = next((c for c in filtered_df.columns if "24日均線" in c or "MA24" in c), None)
                    
                    if close_col and ma24_col:
                        ma_num = clean_number_series(filtered_df[ma24_col])
                        close_num = clean_number_series(filtered_df[close_col])
                        valid_mask = ma_num.notna() & close_num.notna() & (close_num >= ma_num * 0.95) & (close_num <= ma_num * 1.15)
                        filtered_df = filtered_df[valid_mask]

                # 3. 關鍵突破位階 (與上個月高點距離)
                if use_month_high_dist and "與上個月高點距離" in filtered_df.columns:
                    dist_num = clean_number_series(filtered_df["與上個月高點距離"])
                    # 處理百分比 (如果是 3% 可能是 3.0 或 0.03)
                    if dist_num.abs().max() > 1:
                        filtered_df = filtered_df[(dist_num >= -3.0) & (dist_num <= 3.0)]
                    else:
                        filtered_df = filtered_df[(dist_num >= -0.03) & (dist_num <= 0.03)]

                # 4. 籌碼面 (主力買超大於0天數)
                if use_main_buy and "主力買超大於0天數" in filtered_df.columns:
                    buy_num = clean_number_series(filtered_df["主力買超大於0天數"]).fillna(0)
                    filtered_df = filtered_df[buy_num > 0]

                # --- E. 呈現結果與下載 ---
                st.success(f"🎉 篩選完成！從 {len(merged_df)} 檔基準股票中，共篩出 **{len(filtered_df)}** 檔")

                front_cols = ["標準股票代碼", "股票名稱", "收盤價", "24日均線", "與上個月高點距離", "主力買超大於0天數"]
                existing_front = [c for c in front_cols if c in filtered_df.columns]
                other_cols = [c for c in filtered_df.columns if c not in existing_front and not c.startswith("Unnamed")]

                result_df = filtered_df[existing_front + other_cols]
                st.dataframe(result_df, use_container_width=True)

                buffer_out = io.BytesIO()
                with pd.ExcelWriter(buffer_out, engine="openpyxl") as writer:
                    result_df.to_excel(writer, index=False, sheet_name="選股結果")

                st.download_button(
                    label="📥 下載最終選股結果",
                    data=buffer_out.getvalue(),
                    file_name="對位精準篩選結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error(f"❌ 發生錯誤：{str(e)}")