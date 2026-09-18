import io
import re
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="股票策略動態篩選器", layout="wide")
st.title("📈 策略選股：基本面 + 均線位階 + 臨界點突破 + 主力籌碼")

# 1. 檔案上傳元件
uploaded_files = st.file_uploader(
    "請上傳篩選所需的 Excel 報表 (支援多選)",
    type=["xlsx"],
    accept_multiple_files=True,
)


def clean_code_series(series):
    """強效股票代碼清洗：去除 .0、空白、非數字字元並統一補零至 4 位以上"""
    return (
        series.astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.extract(r"(\d{4,6})")[0]
        .str.strip()
    )


def clean_number_series(series):
    """將包含 %, 逗號, -- 的字串安全轉換為純浮點數"""
    if series is None:
        return pd.Series(dtype=float)
    s_clean = (
        series.astype(str)
        .str.replace("%", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.replace("--", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(s_clean, errors="coerce")


def load_excel_smart(file_bytes, kw_sheet=""):
    """修復版：精準跳過備註列、定位正確標題列並抽取標準股票代碼"""
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

    # 尋找匹配的工作表 (Sheet)
    xls = pd.ExcelFile(buffer)
    target_sheet = xls.sheet_names[0]
    if kw_sheet:
        for s in xls.sheet_names:
            if kw_sheet in s:
                target_sheet = s
                break

    buffer.seek(0)
    raw_df = pd.read_excel(buffer, sheet_name=target_sheet, header=None)

    # 精準尋找真正的標題列：必須包含多個欄位，且不可為條件描述列
    header_idx = 0
    for i, row in raw_df.iloc[:25].iterrows():
        row_vals = [str(v).strip() for v in row.values if pd.notna(v)]
        row_str = " ".join(row_vals)

        # 排除說明備註列
        if "篩選條件" in row_str or "報表名稱" in row_str or "產出時間" in row_str:
            continue

        # 必須包含股票代號欄位，且該列非空元素必須大於 3 個
        if (
            any(
                k in row_str
                for k in ["股票代號", "股票代碼", "代號", "股票名稱"]
            )
            and len(row_vals) >= 3
        ):
            header_idx = i
            break

    buffer.seek(0)
    df = pd.read_excel(buffer, sheet_name=target_sheet, header=header_idx)
    df = df.dropna(how="all").reset_index(drop=True)

    # 清理欄位名稱
    df.columns = [
        str(c).strip() if pd.notna(c) else f"Unnamed_{i}"
        for i, c in enumerate(df.columns)
    ]

    # 定位股票代碼欄位並清洗
    code_col = next(
        (
            c
            for c in df.columns
            if "股票代號" in c or "股票代碼" in c or c == "代號"
        ),
        None,
    )

    if code_col:
        df["標準股票代碼"] = clean_code_series(df[code_col])
        df = df[df["標準股票代碼"].notna()]
    else:
        df["標準股票代碼"] = pd.Series(dtype=str)

    return df


def classify_rev_growth(yoy):
    """營收年增率 (YoY) 分級與策略評語"""
    if pd.isna(yoy):
        return "資料缺失", "無營收數據"
    elif yoy > 50:
        return "強勢爆發", "優先關注，通常為飆股預備隊"
    elif 30 <= yoy <= 50:
        return "高成長", "具備強勁基本面動能"
    elif 10 <= yoy < 30:
        return "穩健", "適合中長線分批佈局"
    elif 0 <= yoy < 10:
        return "平庸", "僅達基本門檻，觀察籌碼配合度"
    else:
        return "未達標", "直接剔除"


# 2. 側邊欄條件設定
st.sidebar.header("🎯 篩選條件設定")
use_fundamental = st.sidebar.checkbox(
    "1. 基本面 (營收YoY>0%, 毛利率季增, EPS>0)", value=True
)
use_ma24_bias = st.sidebar.checkbox(
    "2. 安全位階 (收盤價在 MA24 的 -5% ~ +15% 之間)", value=True
)
use_month_high_dist = st.sidebar.checkbox(
    "3. 關鍵突破位階 (距離上月高點 -3% ~ +3%)", value=True
)
use_main_buy = st.sidebar.checkbox("4. 籌碼面 (主力買超天數 > 0)", value=True)

# 3. 執行篩選邏輯
if st.button("🚀 開始執行策略篩選", type="primary"):
    if not uploaded_files:
        st.warning("⚠️ 請上傳相關的 Excel 報表！")
    else:
        file_dict = {f.name: f for f in uploaded_files}

        try:
            with st.spinner("資料讀取與精準整合中..."):
                # 精確區分檔名
                f_chip_center = next(
                    (v for k, v in file_dict.items() if "籌碼集中度" in k), None
                )
                if not f_chip_center:
                    f_chip_center = uploaded_files[0]

                f_chip_detail = next(
                    (
                        v
                        for k, v in file_dict.items()
                        if "型態" in k or "細部" in k
                    ),
                    None,
                )
                f_high = next(
                    (
                        v
                        for k, v in file_dict.items()
                        if "高點" in k or "上月" in k
                    ),
                    None,
                )

                # --- A. 載入主表與各附表 ---
                df_base = load_excel_smart(
                    f_chip_center, kw_sheet="個股圖表資料"
                )
                df_ma = load_excel_smart(f_chip_center, kw_sheet="均線")
                df_detail = (
                    load_excel_smart(f_chip_detail, kw_sheet="近5日")
                    if f_chip_detail
                    else pd.DataFrame()
                )
                df_high = (
                    load_excel_smart(f_high, kw_sheet="高點")
                    if f_high
                    else pd.DataFrame()
                )

                # --- B. 進行跨表合併 (Left Join) ---
                merged_df = df_base.copy()

                # 合併 24MA 均線
                if not df_ma.empty and "標準股票代碼" in df_ma.columns:
                    ma_cols = [
                        c
                        for c in df_ma.columns
                        if "24" in c or "均線" in c or "MA" in c
                    ]
                    if ma_cols:
                        df_ma_sub = (
                            df_ma[["標準股票代碼", ma_cols[-1]]]
                            .drop_duplicates(subset=["標準股票代碼"])
                            .rename(columns={ma_cols[-1]: "MA24均線"})
                        )
                        merged_df = merged_df.merge(
                            df_ma_sub, on="標準股票代碼", how="left"
                        )

                # 合併主力買超天數
                if not df_detail.empty and "標準股票代碼" in df_detail.columns:
                    buy_cols = [
                        c for c in df_detail.columns if "主力買超" in c
                    ]
                    if buy_cols:
                        df_detail_sub = (
                            df_detail[["標準股票代碼", buy_cols[0]]]
                            .drop_duplicates(subset=["標準股票代碼"])
                            .rename(columns={buy_cols[0]: "主力買超天數"})
                        )
                        merged_df = merged_df.merge(
                            df_detail_sub, on="標準股票代碼", how="left"
                        )

                # 合併與上月高點距離
                if not df_high.empty and "標準股票代碼" in df_high.columns:
                    dist_cols = [
                        c
                        for c in df_high.columns
                        if "距離" in c or "高點" in c
                    ]
                    if dist_cols:
                        df_high_sub = (
                            df_high[["標準股票代碼", dist_cols[0]]]
                            .drop_duplicates(subset=["標準股票代碼"])
                            .rename(columns={dist_cols[0]: "距離上月高點幅度"})
                        )
                        merged_df = merged_df.merge(
                            df_high_sub, on="標準股票代碼", how="left"
                        )

                # 除錯看板：即時查看資料讀取與合併結果
                st.subheader("🔍 整合資料狀態監控")
                col1, col2 = st.columns(2)
                col1.metric("主表成功載入筆數", f"{len(df_base)} 筆")
                col2.metric("跨表整合完成筆數", f"{len(merged_df)} 筆")

                # --- C. 條件篩選 ---
                filtered_df = merged_df.copy()

                # 1. 基本面篩選 (YoY > 0, 毛利率季增, EPS > 0)
                if use_fundamental:
                    rev_col = next(
                        (
                            c
                            for c in filtered_df.columns
                            if "營收年增" in c or "YoY" in c
                        ),
                        None,
                    )
                    eps_col = next(
                        (
                            c
                            for c in filtered_df.columns
                            if "EPS" in c or "每股盈餘" in c
                        ),
                        None,
                    )
                    gm_curr_col = next(
                        (
                            c
                            for c in filtered_df.columns
                            if "毛利率" in c and "前" not in c
                        ),
                        None,
                    )
                    gm_prev_col = next(
                        (
                            c
                            for c in filtered_df.columns
                            if "前一季毛利率" in c or "前季毛利率" in c
                        ),
                        None,
                    )

                    if rev_col:
                        filtered_df = filtered_df[
                            clean_number_series(filtered_df[rev_col]) > 0
                        ]
                    if eps_col:
                        filtered_df = filtered_df[
                            clean_number_series(filtered_df[eps_col]) > 0
                        ]
                    if gm_curr_col and gm_prev_col:
                        gm_c = clean_number_series(filtered_df[gm_curr_col])
                        gm_p = clean_number_series(filtered_df[gm_prev_col])
                        filtered_df = filtered_df[gm_c > gm_p]

                # 2. 安全位階：收盤價在 MA24 的 0.95 ~ 1.15 倍之間
                if use_ma24_bias and "MA24均線" in filtered_df.columns:
                    close_col = next(
                        (
                            c
                            for c in filtered_df.columns
                            if "收盤價" in c
                        ),
                        None,
                    )
                    if close_col:
                        ma_num = clean_number_series(filtered_df["MA24均線"])
                        close_num = clean_number_series(filtered_df[close_col])

                        valid_mask = (
                            ma_num.notna()
                            & close_num.notna()
                            & (close_num >= ma_num * 0.95)
                            & (close_num <= ma_num * 1.15)
                        )
                        filtered_df = filtered_df[valid_mask]

                # 3. 臨界點突破：距離上月高點介於 -3% ~ +3%
                if (
                    use_month_high_dist
                    and "距離上月高點幅度" in filtered_df.columns
                ):
                    dist_num = clean_number_series(
                        filtered_df["距離上月高點幅度"]
                    )
                    if dist_num.abs().max() > 1:
                        filtered_df = filtered_df[
                            (dist_num >= -3.0) & (dist_num <= 3.0)
                        ]
                    else:
                        filtered_df = filtered_df[
                            (dist_num >= -0.03) & (dist_num <= 0.03)
                        ]

                # 4. 主力買超天數 > 0
                if use_main_buy and "主力買超天數" in filtered_df.columns:
                    buy_num = clean_number_series(
                        filtered_df["主力買超天數"]
                    ).fillna(0)
                    filtered_df = filtered_df[buy_num > 0]

                # --- D. 營收分級與標記 ---
                rev_col = next(
                    (
                        c
                        for c in filtered_df.columns
                        if "營收年增" in c or "YoY" in c
                    ),
                    None,
                )
                if rev_col:
                    yoy_vals = clean_number_series(filtered_df[rev_col])
                    grade_res = [classify_rev_growth(v) for v in yoy_vals]
                    filtered_df["營收成長等級"] = [
                        res[0] for res in grade_res
                    ]
                    filtered_df["策略評語"] = [
                        res[1] for res in grade_res
                    ]
                    filtered_df = filtered_df[
                        filtered_df["營收成長等級"] != "未達標"
                    ]

                # --- E. 呈現結果與下載 ---
                st.success(
                    f"🎉 篩選完成！共找到 **{len(filtered_df)}** 檔符合策略條件的股票"
                )

                # 欄位順序美化
                front_cols = [
                    "標準股票代碼",
                    "股票名稱",
                    "營收成長等級",
                    "策略評語",
                ]
                existing_front = [
                    c for c in front_cols if c in filtered_df.columns
                ]
                other_cols = [
                    c
                    for c in filtered_df.columns
                    if c not in existing_front
                    and not c.startswith("Unnamed")
                ]

                result_df = filtered_df[existing_front + other_cols]
                st.dataframe(result_df, use_container_width=True)

                # Excel 下載
                buffer_out = io.BytesIO()
                with pd.ExcelWriter(buffer_out, engine="openpyxl") as writer:
                    result_df.to_excel(writer, index=False, sheet_name="選股結果")

                st.download_button(
                    label="📥 下載策略選股結果 Excel 檔",
                    data=buffer_out.getvalue(),
                    file_name="基本面與技術籌碼策略選股結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error(f"❌ 執行過程中發生錯誤：{str(e)}")