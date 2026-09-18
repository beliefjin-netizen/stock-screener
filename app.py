import io
import re
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="股票策略動態篩選器", layout="wide")
st.title("📈 策略選股：基本面 + 均線位階 + 臨界點突破 + 主力籌碼")

# 1. 檔案上傳元件
uploaded_files = st.file_uploader(
    "請上傳篩選所需的 Excel 報表 (支援多選，檔名不限)",
    type=["xlsx"],
    accept_multiple_files=True,
    help="請確保上傳的報表中包含：基本面(營收/毛利率/EPS)、均線位階、上月高點距離與主力籌碼資料",
)


def clean_code_series(series):
  """股票代碼清洗：補齊 4 位數，去除 .0 與無效符號"""
  s = series.astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
  s = s.apply(lambda x: x.zfill(4) if x.isdigit() and len(x) < 4 else x)
  return s.str.extract(r"(\d{4,6})")[0]


def clean_number_series(series):
  """將文字轉為純浮點數（自動清除 %、逗號、-- 等文字）"""
  if series is None or series.empty:
    return pd.Series(dtype=float)
  s_clean = (
      series.astype(str)
      .str.replace(r"[%,\s]", "", regex=True)
      .str.replace("--", "NaN", regex=False)
      .str.replace("NA", "NaN", regex=False)
  )
  return pd.to_numeric(s_clean, errors="coerce")


def load_excel_smart(file_bytes):
  """智慧讀取 Excel：自動跳過備註列並提取正確標題列"""
  if not file_bytes:
    return pd.DataFrame(), ""

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

  best_df = pd.DataFrame()
  best_sheet_name = xls.sheet_names[0]
  max_score = -1

  for s in xls.sheet_names:
    buffer.seek(0)
    raw_df = pd.read_excel(buffer, sheet_name=s, header=None)
    for i, row in raw_df.iloc[:50].iterrows():
      cells = [
          str(x).strip()
          for x in row.values
          if pd.notna(x) and str(x).strip() != ""
      ]
      row_str = "".join(cells)

      if any(
          kw in row_str
          for kw in ["報表名稱", "篩選條件", "產出時間", "條件描述"]
      ):
        continue

      score = sum(
          1
          for kw in [
              "股票",
              "代碼",
              "代號",
              "名稱",
              "收盤",
              "均線",
              "營收",
              "毛利",
              "EPS",
              "距離",
              "主力",
          ]
          if any(kw in c for c in cells)
      )
      if score > max_score and len(cells) >= 2:
        max_score = score
        best_sheet_name = s

        buffer.seek(0)
        df = pd.read_excel(buffer, sheet_name=s, header=i)
        df = df.dropna(how="all").reset_index(drop=True)

        # 清洗欄位名稱：剔除換行符號 \n 與空格，確保字串精準比對
        df.columns = [
            str(c)
            .replace("\n", "")
            .replace("\r", "")
            .replace(" ", "")
            .strip()
            if pd.notna(c)
            else f"Unnamed_{i}"
            for i, c in enumerate(df.columns)
        ]

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

        best_df = df

  return best_df, best_sheet_name


# 2. 側邊欄條件設定
st.sidebar.header("🎯 篩選條件設定")
use_fundamental = st.sidebar.checkbox(
    "1. 基本面 (營收年成長>0 AND 毛利率最新>前期 AND EPS>0)", value=True
)
use_ma24_bias = st.sidebar.checkbox(
    "2. 安全位階 (收盤價在 24日均線 的 -5% ~ +15% 之間)", value=True
)
use_month_high_dist = st.sidebar.checkbox(
    "3. 關鍵突破位階 (與上個月高點距離 -3% ~ +3%)", value=True
)
use_main_buy = st.sidebar.checkbox(
    "4. 籌碼面 (主力買超大於0天數 > 0)", value=True
)

# 3. 執行邏輯
if st.button("🚀 開始執行策略篩選", type="primary"):
  if not uploaded_files:
    st.warning("⚠️ 請先上傳 Excel 報表！")
  else:
    try:
      with st.spinner("自動解析上傳檔案與進行跨表篩選中..."):

        # 自動判讀各檔案類型（由資料內容欄位特徵自動分類，不依賴檔名）
        df_fund, df_tech, df_high, df_chip = (
            pd.DataFrame(),
            pd.DataFrame(),
            pd.DataFrame(),
            pd.DataFrame(),
        )
        file_details = {}

        for f in uploaded_files:
          df, sheet_name = load_excel_smart(f)
          cols = list(df.columns)
          file_details[f.name] = {
              "sheet": sheet_name,
              "cols": cols,
              "rows": len(df),
          }

          if any(
              "毛利率" in c or "EPS" in c.upper() or "每股盈餘" in c or "營收" in c
              for c in cols
          ):
            df_fund = df
          elif any(
              "24日均線" in c or "MA24" in c or ("收盤" in c and "高點" not in c)
              for c in cols
          ):
            df_tech = df
          elif any("高點" in c or "距離" in c for c in cols):
            df_high = df
          elif any("主力" in c or "買超" in c for c in cols):
            df_chip = df

        # 診斷面板：顯示載入的檔案與辨識結果
        with st.expander("🔍 檔名與欄位對應自動診斷結果", expanded=False):
          for name, info in file_details.items():
            st.write(
                f"📄 **檔名**: {name} (頁籤: {info['sheet']}, 筆數:"
                f" {info['rows']})"
            )
            st.caption(f"包含欄位: {', '.join(info['cols'])}")

        # A. 跨表合併 (以標籤最齊全的表作為 Base)
        merged_df = (
            df_tech.copy()
            if not df_tech.empty
            else (
                df_fund.copy()
                if not df_fund.empty
                else df_high.copy() if not df_high.empty else df_chip.copy()
            )
        )

        if not df_fund.empty and "標準股票代碼" in df_fund.columns:
          cols_to_use = [
              c
              for c in df_fund.columns
              if c not in merged_df.columns or c == "標準股票代碼"
          ]
          merged_df = merged_df.merge(
              df_fund[cols_to_use].drop_duplicates(subset=["標準股票代碼"]),
              on="標準股票代碼",
              how="left",
          )

        if not df_high.empty and "標準股票代碼" in df_high.columns:
          dist_cols = [
              c for c in df_high.columns if "高點距離" in c or "與上個月高點距離" in c
          ]
          if dist_cols:
            df_high_sub = (
                df_high[["標準股票代碼", dist_cols[0]]]
                .drop_duplicates(subset=["標準股票代碼"])
                .rename(columns={dist_cols[0]: "與上個月高點距離"})
            )
            merged_df = merged_df.merge(
                df_high_sub, on="標準股票代碼", how="left"
            )

        if not df_chip.empty and "標準股票代碼" in df_chip.columns:
          buy_cols = [
              c for c in df_chip.columns if "主力買超" in c or "大於0天數" in c
          ]
          if buy_cols:
            df_chip_sub = (
                df_chip[["標準股票代碼", buy_cols[0]]]
                .drop_duplicates(subset=["標準股票代碼"])
                .rename(columns={buy_cols[0]: "主力買超大於0天數"})
            )
            merged_df = merged_df.merge(
                df_chip_sub, on="標準股票代碼", how="left"
            )

        # B. 執行條件篩選
        filtered_df = merged_df.copy()

        # 1. 基本面篩選 (營收年成長>0 AND 毛利率最新>前期 AND EPS>0)
        if use_fundamental:
          rev_col = next(
              (
                  c
                  for c in filtered_df.columns
                  if "營收年成長" in c or "營收年增" in c or "營收" in c
              ),
              None,
          )
          eps_col = next(
              (
                  c
                  for c in filtered_df.columns
                  if "EPS" in c.upper() or "每股盈餘" in c
              ),
              None,
          )

          gm_cols = [c for c in filtered_df.columns if "毛利率" in c]
          gm_prev_col = next(
              (
                  c
                  for c in gm_cols
                  if any(k in c for k in ["前", "上", "舊", "1期", "前期"])
              ),
              None,
          )
          gm_curr_col = (
              next((c for c in gm_cols if c != gm_prev_col), None)
              if gm_cols
              else None
          )

          missing_cols = []
          if not rev_col:
            missing_cols.append("營收年成長")
          if not eps_col:
            missing_cols.append("EPS")
          if not (gm_curr_col and gm_prev_col):
            missing_cols.append("毛利率(最新/前期)")

          if missing_cols:
            st.error(
                f"❌ 基本面篩選失敗：上傳檔案中缺少以下欄位：{', '.join(missing_cols)}。"
                "請展開上方『檔名與欄位對應自動診斷結果』檢查上傳資料。"
            )
          else:
            cond_rev = clean_number_series(filtered_df[rev_col]) > 0
            cond_eps = clean_number_series(filtered_df[eps_col]) > 0

            gm_curr = clean_number_series(filtered_df[gm_curr_col])
            gm_prev = clean_number_series(filtered_df[gm_prev_col])
            cond_gm = (gm_curr > gm_prev) & gm_curr.notna() & gm_prev.notna()

            fundamental_mask = cond_rev & cond_gm & cond_eps
            filtered_df = filtered_df[fundamental_mask]

        # 2. 安全位階 (-5% ~ +15%)
        if use_ma24_bias:
          close_col = next(
              (c for c in filtered_df.columns if "收盤價" in c), None
          )
          ma24_col = next(
              (c for c in filtered_df.columns if "24日均線" in c or "MA24" in c),
              None,
          )

          if close_col and ma24_col:
            ma_num = clean_number_series(filtered_df[ma24_col])
            close_num = clean_number_series(filtered_df[close_col])
            valid_mask = (
                ma_num.notna()
                & close_num.notna()
                & (close_num >= ma_num * 0.95)
                & (close_num <= ma_num * 1.15)
            )
            filtered_df = filtered_df[valid_mask]

        # 3. 關鍵突破位階 (-3% ~ +3%)
        if use_month_high_dist and "與上個月高點距離" in filtered_df.columns:
          dist_num = clean_number_series(filtered_df["與上個月高點距離"])
          if dist_num.abs().max() > 1:
            filtered_df = filtered_df[(dist_num >= -3.0) & (dist_num <= 3.0)]
          else:
            filtered_df = filtered_df[
                (dist_num >= -0.03) & (dist_num <= 0.03)
            ]

        # 4. 籌碼面
        if use_main_buy and "主力買超大於0天數" in filtered_df.columns:
          buy_num = clean_number_series(filtered_df["主力買超大於0天數"]).fillna(
              0
          )
          filtered_df = filtered_df[buy_num > 0]

        # C. 呈現結果與下載
        st.success(
            f"🎉 篩選完成！從 {len(merged_df)} 檔股票中，共過濾出 **{len(filtered_df)}** 檔精選標的"
        )

        front_cols = [
            "標準股票代碼",
            "股票名稱",
            "收盤價",
            "24日均線",
            "與上個月高點距離",
            "主力買超大於0天數",
        ]
        existing_front = [c for c in front_cols if c in filtered_df.columns]
        other_cols = [
            c
            for c in filtered_df.columns
            if c not in existing_front and not c.startswith("Unnamed")
        ]

        result_df = filtered_df[existing_front + other_cols]
        st.dataframe(result_df, use_container_width=True)

        buffer_out = io.BytesIO()
        with pd.ExcelWriter(buffer_out, engine="openpyxl") as writer:
          result_df.to_excel(writer, index=False, sheet_name="選股結果")

        st.download_button(
            label="📥 下載最終選股結果 Excel",
            data=buffer_out.getvalue(),
            file_name="基本面與技術面綜合篩選結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
      st.error(f"❌ 發生錯誤：{str(e)}")