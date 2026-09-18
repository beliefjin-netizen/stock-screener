import io
import re
import zipfile
import pandas as pd
import streamlit as st

# 頁面標題
st.title("📈 股票動態篩選器")

# 1. 上傳檔案元件 (讓使用者直接在網頁拖曳上傳 5 個 Excel)
uploaded_files = st.file_uploader(
    "請上傳篩選所需的 Excel 報表 (可多選)",
    type=["xlsx"],
    accept_multiple_files=True,
)


def load_excel_cleaned(file_bytes, sheet_name=0):
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
    return pd.read_excel(buffer, sheet_name=sheet_name)


# 2. 側邊欄：篩選條件勾選開關
st.sidebar.header("🎯 篩選條件設定")
use_fundamental = st.sidebar.checkbox("基本面 (營收/毛利率/EPS)", value=True)
use_ma24_safe = st.sidebar.checkbox("安全位階 (24MA)", value=True)
use_rev_grade = st.sidebar.checkbox("營收分級成長性", value=True)
use_month_high = st.sidebar.checkbox("突破/距離上月高點", value=True)
use_main_buy = st.sidebar.checkbox("主力買超天數 > 0", value=True)

# 3. 執行按鈕與邏輯
if st.button("🚀 開始執行篩選", type="primary"):
    if not uploaded_files:
        st.warning("⚠️ 請先上傳 Excel 檔案！")
    else:
        # 將上傳的檔案分類整理並讀取 (簡化範例)
        st.success("分析完成！")
        # 顯示結果表格
        # st.dataframe(result_df)