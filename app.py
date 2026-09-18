import io
import re
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="股票策略動態篩選器", layout="wide")
st.title("📈 策略選股：基本面 + 均線位階 + 臨界點突破 + 主力籌碼")

# 1. 檔案上傳元件
uploaded_files = st.file_uploader(
    "請上傳篩選所需的 Excel 報表",
    type=["xlsx"],
    accept_multiple_files=True,
)


def load_excel_smart(file_bytes, sheet_name=0):
    """清理 CMoney autofilter 錯誤並自動定位『股票代號』所在的標題列"""
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

    # 讀取完整 raw data
    raw_df = pd.read_excel(buffer, sheet_name=sheet_name, header=None)

    # 定位標題列
    header_idx = 0
    for i, row in raw_df.iloc[:10].iterrows():
        row_str_joined = "".join(
            [str(val) for val in row.values if pd.notna(val)]
        )
        if any(
            k in row_str_joined for k in ["股票代號", "股票代碼", "代號"]
        ):
            header_idx = i
            break

    # 解析標題與清理
    buffer.seek(0)
    df = pd.read_excel(buffer, sheet_name=sheet_name, header=header_idx)
    df.columns = [
        str(c).strip() if pd.notna(c) else f"Unnamed_{i}"
        for i, c in enumerate(df.columns)
    ]

    # 格式化股票代碼
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
        df = df[df["標準股票代碼"].str.contains(r"^\d{4,6}$", na=False)]

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
            with st.spinner("資料整合與策略計算中..."):
                # --- A. 讀取基礎資料 (基本面/圖表資料) ---
                f_chip_center = next(
                    (
                        v
                        for k, v in file_dict.items()
                        if "籌碼集中度" in k or "基本面" in k
                    ),
                    None,
                )
                if not f_chip_center:
                    f_chip_center = list(file_dict.values())[0]

                df_base = load_excel_smart(
                    f_chip_center, sheet_name="個股圖表資料"
                )

                # 讀取 24MA 頁籤
                try:
                    df_ma = load_excel_smart(
                        f_chip_center, sheet_name="均線"
                    )
                except Exception:
                    df_ma = None

                # --- B. 讀取籌碼細部與高點資料 ---
                f_chip_detail = next(
                    (
                        v
                        for k, v in file_dict.items()
                        if "型態" in k or "籌碼" in k
                    ),
                    None,
                )
                df_detail = (
                    load_excel_smart(f_chip_detail, sheet_name="近5日")
                    if f_chip_detail
                    else None
                )

                f_high = next(
                    (
                        v
                        for k, v in file_dict.items()
                        if "高點" in k or "上月" in k
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

                # --- C. 資料表整合 (Merge) ---
                merged_df = df_base.copy()

                # 合併 24MA 均線
                if df_ma is not None and "標準股票代碼" in df_ma.columns:
                    ma_cols = [
                        c
                        for c in df_ma.columns
                        if "24" in c or "均線" in c or "MA" in c
                    ]
                    if ma_cols:
                        merged_df = merged_df.merge(
                            df_ma[["標準股票代碼", ma_cols[-1]]].rename(
                                columns={ma_cols[-1]: "MA24均線"}
                            ),
                            on="標準股票代碼",
                            how="left",
                        )

                # 合併主力買超天數
                if df_detail is not None and "標準股票代碼" in df_detail.columns:
                    buy_cols = [
                        c for c in df_detail.columns if "主力買超" in c
                    ]
                    if buy_cols:
                        merged_df = merged_df.merge(
                            df_detail[
                                ["標準股票代碼", buy_cols[0]]
                            ].rename(columns={buy_cols[0]: "主力買超天數"}),
                            on="標準股票代碼",
                            how="left",
                        )

                # 合併與上月高點距離
                if df_high is not None and "標準股票代碼" in df_high.columns:
                    dist_cols = [
                        c
                        for c in df_high.columns
                        if "距離" in c or "高點" in c
                    ]
                    if dist_cols:
                        merged_df = merged_df.merge(
                            df_high[
                                ["標準股票代碼", dist_cols[0]]
                            ].rename(
                                columns={dist_cols[0]: "距離上月高點幅度"}
                            ),
                            on="標準股票代碼",
                            how="left",
                        )

                # --- D. 條件過濾 ---
                filtered_df = merged_df.copy()

                # 1. 基本面篩選 (YoY > 0, 毛利率季增, EPS > 0)
                if use_fundamental:
                    rev_yoy_col = next(
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

                    if rev_yoy_col:
                        filtered_df = filtered_df[
                            pd.to_numeric(
                                filtered_df[rev_yoy_col], errors="coerce"
                            )
                            > 0
                        ]
                    if eps_col:
                        filtered_df = filtered_df[
                            pd.to_numeric(
                                filtered_df[eps_col], errors="coerce"
                            )
                            > 0
                        ]
                    if gm_curr_col and gm_prev_col:
                        filtered_df = filtered_df[
                            pd.to_numeric(
                                filtered_df[gm_curr_col], errors="coerce"
                            )
                            > pd.to_numeric(
                                filtered_df[gm_prev_col], errors="coerce"
                            )
                        ]

                # 2. 安全位階：收盤價在 MA24 的 0.95 ~ 1.15 倍之間 (乖離率 -5% ~ +15%)
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
                        filtered_df["MA24"] = pd.to_numeric(
                            filtered_df["MA24均線"], errors="coerce"
                        )
                        filtered_df["Close"] = pd.to_numeric(
                            filtered_df[close_col], errors="coerce"
                        )

                        filtered_df = filtered_df[
                            (
                                filtered_df["Close"]
                                >= filtered_df["MA24"] * 0.95
                            )
                            & (
                                filtered_df["Close"]
                                <= filtered_df["MA24"] * 1.15
                            )
                        ]

                # 3. 臨界點突破：距離上月高點介於 -3% ~ +3%
                if (
                    use_month_high_dist
                    and "距離上月高點幅度" in filtered_df.columns
                ):
                    dist_num = pd.to_numeric(
                        filtered_df["距離上月高點幅度"], errors="coerce"
                    )
                    # 自動相容 % 單位 (例如 3% 為 0.03 或 3)
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
                    filtered_df = filtered_df[
                        pd.to_numeric(
                            filtered_df["主力買超天數"], errors="coerce"
                        ).fillna(0)
                        > 0
                    ]

                # --- E. 營收分級與標記 ---
                rev_yoy_col = next(
                    (
                        c
                        for c in filtered_df.columns
                        if "營收年增" in c or "YoY" in c
                    ),
                    None,
                )
                if rev_yoy_col:
                    yoy_vals = pd.to_numeric(
                        filtered_df[rev_yoy_col], errors="coerce"
                    )
                    grade_res = [classify_rev_growth(val) for val in yoy_vals]
                    filtered_df["營收成長等級"] = [
                        res[0] for res in grade_res
                    ]
                    filtered_df["策略評語"] = [
                        res[1] for res in grade_res
                    ]

                    # 剔除未達標 (YoY < 0)
                    filtered_df = filtered_df[
                        filtered_df["營收成長等級"] != "未達標"
                    ]

                # --- F. 結果呈現與匯出 ---
                st.success(
                    f"🎉 篩選完成！共找到 {len(filtered_df)} 檔符合策略的潛力個股"
                )

                # 調整顯示欄位順序
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
                    and c not in ["MA24", "Close"]
                ]

                result_df = filtered_df[existing_front + other_cols]

                # 顯示表格
                st.dataframe(result_df, use_container_width=True)

                # 提供 Excel 下載
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