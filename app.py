import io
import re
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="股票策略動態篩選器", layout="wide")
st.title("📈 策略選股：基本面 + 均線位階 + 臨界點突破 + 主力籌碼")

# 1. 檔案上傳元件
uploaded_files = st.file_uploader(
    "請上傳篩選所需的 Excel 報表 (支援多選，系統會自動辨識檔名與頁籤)",
    type=["xlsx"],
    accept_multiple_files=True,
)


def clean_code_series(series):
    """終極股票代碼清洗：解決 0050 變 50、帶小數點、前後空白等所有問題"""
    s = series.astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    # 將純數字但不足 4 碼的代碼補零 (解決 0050 問題)
    s = s.apply(lambda x: x.zfill(4) if x.isdigit() and len(x) < 4 else x)
    # 強制抽出 4~6 碼數字
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


def load_excel_smart(file_bytes, kw_sheet_list):
    """智慧讀取：透過評分機制精準定位真實資料表頭，無視任何干擾備註"""
    buffer = io.BytesIO()
    with zipfile.ZipFile(file_bytes, "r") as zin:
        with zipfile.ZipFile(buffer, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("xl/worksheets/sheet"):
                    # 強制移除擾人的 AutoFilter
                    data = re.sub(rb"<autoFilter[^>]*/>", b"", data)
                    data = re.sub(rb"<autoFilter[^>]*>.*?</autoFilter>", b"", data)
                zout.writestr(item, data)
    
    buffer.seek(0)
    xls = pd.ExcelFile(buffer)
    
    # 尋找最符合關鍵字的頁籤
    target_sheet = xls.sheet_names[0]
    for kw in kw_sheet_list:
        for s in xls.sheet_names:
            if kw in s:
                target_sheet = s
                break

    buffer.seek(0)
    raw_df = pd.read_excel(buffer, sheet_name=target_sheet, header=None)

    # 【核心修復】評分制尋找真實標題列
    best_idx = 0
    max_score = 0
    for i, row in raw_df.iloc[:50].iterrows():
        cells = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != ""]
        row_str = "".join(cells)
        
        # 遇到文件資訊列直接跳過扣分
        if any(kw in row_str for kw in ["報表名稱", "篩選條件", "產出時間", "條件描述"]):
            continue
            
        score = 0
        for kw in ["股票", "代碼", "代號", "名稱", "收盤", "成交", "均線", "高點", "營收", "毛利", "主力"]:
            if any(kw in c for c in cells):
                score += 1
                
        # 標題列通常具備多個財經欄位 (大於 2 分)，且長度夠長
        if score > max_score and len(cells) >= 3:
            max_score = score
            best_idx = i

    # 正式讀取
    buffer.seek(0)
    df = pd.read_excel(buffer, sheet_name=target_sheet, header=best_idx)
    df = df.dropna(how="all").reset_index(drop=True)
    
    df.columns = [str(c).strip() if pd.notna(c) else f"Unnamed_{i}" for i, c in enumerate(df.columns)]

    # 尋找股票代碼欄位並轉換
    code_col = next((c for c in df.columns if "股票代號" in c or "股票代碼" in c or c == "代號"), None)
    if code_col:
        df["標準股票代碼"] = clean_code_series(df[code_col])
        df = df[df["標準股票代碼"].notna()]
    else:
        df["標準股票代碼"] = pd.Series(dtype=str)

    return df


def classify_rev_growth(yoy):
    """營收年增率 (YoY) 分級標籤"""
    if pd.isna(yoy): return "資料缺失", "無營收數據"
    elif yoy > 50: return "強勢爆發", "優先關注，通常為飆股預備隊"
    elif 30 <= yoy <= 50: return "高成長", "具備強勁基本面動能"
    elif 10 <= yoy < 30: return "穩健", "適合中長線分批佈局"
    elif 0 <= yoy < 10: return "平庸", "僅達基本門檻，觀察籌碼配合度"
    else: return "未達標", "直接剔除"


# 2. 側邊欄條件設定
st.sidebar.header("🎯 篩選條件設定")
use_fundamental = st.sidebar.checkbox("1. 基本面 (營收YoY>0%, 毛利率季增, EPS>0)", value=True)
use_ma24_bias = st.sidebar.checkbox("2. 安全位階 (收盤價在 MA24 的 -5% ~ +15% 之間)", value=True)
use_month_high_dist = st.sidebar.checkbox("3. 關鍵突破位階 (距離上月高點 -3% ~ +3%)", value=True)
use_main_buy = st.sidebar.checkbox("4. 籌碼面 (主力買超天數 > 0)", value=True)

# 3. 執行邏輯
if st.button("🚀 開始執行策略篩選", type="primary"):
    if not uploaded_files:
        st.warning("⚠️ 請上傳相關的 Excel 報表！")
    else:
        file_dict = {f.name: f for f in uploaded_files}

        try:
            with st.spinner("資料讀取與精準整合中..."):
                
                # 智慧檔案分配器：防止檔案被錯誤重複覆蓋
                f_chip_center, f_chip_detail, f_high = None, None, None
                for name, f in file_dict.items():
                    if "高點" in name or "上月" in name:
                        f_high = f
                    elif "型態" in name or "細部" in name or "近5日" in name:
                        f_chip_detail = f
                    else:
                        # 預設把剩餘的主力報表丟給主表
                        f_chip_center = f_chip_center or f
                
                if not f_chip_center and uploaded_files:
                    f_chip_center = uploaded_files[0]

                # --- A. 載入各表 ---
                df_base = load_excel_smart(f_chip_center, ["圖表資料", "基本面", "個股"])
                df_ma = load_excel_smart(f_chip_center, ["均線", "MA"])
                df_detail = load_excel_smart(f_chip_detail, ["近5日", "籌碼"]) if f_chip_detail else pd.DataFrame()
                df_high = load_excel_smart(f_high, ["高點", "收盤價大於"]) if f_high else pd.DataFrame()

                # --- B. 進行跨表合併 (Left Join) ---
                merged_df = df_base.copy()

                if not df_ma.empty and "標準股票代碼" in df_ma.columns:
                    ma_cols = [c for c in df_ma.columns if "24" in str(c) or "月線" in str(c) or "MA" in str(c).upper()]
                    if ma_cols:
                        df_ma_sub = df_ma[["標準股票代碼", ma_cols[-1]]].drop_duplicates(subset=["標準股票代碼"]).rename(columns={ma_cols[-1]: "MA24均線"})
                        merged_df = merged_df.merge(df_ma_sub, on="標準股票代碼", how="left")

                if not df_detail.empty and "標準股票代碼" in df_detail.columns:
                    buy_cols = [c for c in df_detail.columns if "主力買超" in c or "買超天數" in c]
                    if buy_cols:
                        df_detail_sub = df_detail[["標準股票代碼", buy_cols[0]]].drop_duplicates(subset=["標準股票代碼"]).rename(columns={buy_cols[0]: "主力買超天數"})
                        merged_df = merged_df.merge(df_detail_sub, on="標準股票代碼", how="left")

                if not df_high.empty and "標準股票代碼" in df_high.columns:
                    dist_cols = [c for c in df_high.columns if "距離" in c or "高點" in c]
                    if dist_cols:
                        df_high_sub = df_high[["標準股票代碼", dist_cols[0]]].drop_duplicates(subset=["標準股票代碼"]).rename(columns={dist_cols[0]: "距離上月高點幅度"})
                        merged_df = merged_df.merge(df_high_sub, on="標準股票代碼", how="left")

                # 【監控面板】在網頁上印出每張表的讀取狀態，幫助追蹤資料去向
                with st.expander("📊 資料處理與整合監控狀態 (若篩選結果異常請點此查看)", expanded=True):
                    st.markdown(f"""
                    - **主表 (基準資料)**: 成功載入 `{len(df_base)}` 筆
                    - **副表 A (均線資料)**: 成功載入 `{len(df_ma)}` 筆
                    - **副表 B (籌碼資料)**: 成功載入 `{len(df_detail)}` 筆
                    - **副表 C (高點資料)**: 成功載入 `{len(df_high)}` 筆
                    - **跨表合併完成後**: 基準庫維持 `{len(merged_df)}` 筆等待篩選
                    """)

                # --- C. 條件篩選 ---
                filtered_df = merged_df.copy()

                if use_fundamental:
                    rev_col = next((c for c in filtered_df.columns if "營收年增" in c or "YoY" in c), None)
                    eps_col = next((c for c in filtered_df.columns if "EPS" in c or "每股盈餘" in c), None)
                    gm_curr_col = next((c for c in filtered_df.columns if "毛利率" in c and "前" not in c), None)
                    gm_prev_col = next((c for c in filtered_df.columns if "前一季毛利率" in c or "前季毛利率" in c), None)

                    if rev_col:
                        filtered_df = filtered_df[clean_number_series(filtered_df[rev_col]) > 0]
                    if eps_col:
                        filtered_df = filtered_df[clean_number_series(filtered_df[eps_col]) > 0]
                    if gm_curr_col and gm_prev_col:
                        filtered_df = filtered_df[clean_number_series(filtered_df[gm_curr_col]) > clean_number_series(filtered_df[gm_prev_col])]

                if use_ma24_bias and "MA24均線" in filtered_df.columns:
                    close_col = next((c for c in filtered_df.columns if c.strip() == "收盤價" or "收盤" in c), None)
                    if close_col:
                        ma_num = clean_number_series(filtered_df["MA24均線"])
                        close_num = clean_number_series(filtered_df[close_col])
                        valid_mask = ma_num.notna() & close_num.notna() & (close_num >= ma_num * 0.95) & (close_num <= ma_num * 1.15)
                        filtered_df = filtered_df[valid_mask]

                if use_month_high_dist and "距離上月高點幅度" in filtered_df.columns:
                    dist_num = clean_number_series(filtered_df["距離上月高點幅度"])
                    # 相容 % 寫法與小數寫法
                    if dist_num.abs().max() > 1:
                        filtered_df = filtered_df[(dist_num >= -3.0) & (dist_num <= 3.0)]
                    else:
                        filtered_df = filtered_df[(dist_num >= -0.03) & (dist_num <= 0.03)]

                if use_main_buy and "主力買超天數" in filtered_df.columns:
                    buy_num = clean_number_series(filtered_df["主力買超天數"]).fillna(0)
                    filtered_df = filtered_df[buy_num > 0]

                # --- D. 營收分級標記 ---
                rev_col = next((c for c in filtered_df.columns if "營收年增" in c or "YoY" in c), None)
                if rev_col:
                    grade_res = [classify_rev_growth(v) for v in clean_number_series(filtered_df[rev_col])]
                    filtered_df["營收成長等級"] = [r[0] for r in grade_res]
                    filtered_df["策略評語"] = [r[1] for r in grade_res]
                    # 只過濾掉營收衰退，保留資料缺失者以免誤殺
                    filtered_df = filtered_df[filtered_df["營收成長等級"] != "未達標"]

                # --- E. 呈現結果與下載 ---
                st.success(f"🎉 篩選完成！從 {len(merged_df)} 檔基準股票中，共篩出 **{len(filtered_df)}** 檔潛力個股")

                # 美化呈現順序
                front_cols = ["標準股票代碼", "股票名稱", "營收成長等級", "策略評語", "收盤價"]
                existing_front = [c for c in front_cols if c in filtered_df.columns]
                other_cols = [c for c in filtered_df.columns if c not in existing_front and not c.startswith("Unnamed")]

                result_df = filtered_df[existing_front + other_cols]
                st.dataframe(result_df, use_container_width=True)

                buffer_out = io.BytesIO()
                with pd.ExcelWriter(buffer_out, engine="openpyxl") as writer:
                    result_df.to_excel(writer, index=False, sheet_name="選股結果")

                st.download_button(
                    label="📥 下載策略選股結果 Excel 檔",
                    data=buffer_out.getvalue(),
                    file_name="動態策略選股結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error(f"❌ 執行過程中發生錯誤：{str(e)}")
            st.warning("提示：請確認上傳的報表是否為 CMoney 或 PressPlay 匯出之標準格式。")