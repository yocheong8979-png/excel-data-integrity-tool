# ============================================
# Excel 資料完整性分類工具
# Streamlit Web App
# ============================================

import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Excel 資料完整性分類工具",
    page_icon="📊",
    layout="wide",
)

# ============================================
# 頁面標題與說明
# ============================================

st.title("📊 Excel 資料完整性分類工具")
st.markdown(
    """
    上傳 `.xlsx` 或 `.xls` Excel 檔案，系統會自動檢查：

    - 空值、NULL、N/A、NA、None、unknown
    - 日期格式
    - 電話號碼格式
    - 電郵格式
    - 年齡、薪金、消費金額、折扣率等異常值
    - 分類欄位的大小寫混亂

    檢查完成後，可下載包含 **完整資料** 與 **不完整資料** 兩個工作表的 Excel 檔。
    """
)

# ============================================
# 資料檢查函數
# ============================================

MISSING_WORDS = {"", "null", "n/a", "na", "none", "nan", "unknown"}


def is_missing(value):
    """判斷值是否屬於缺失值。"""
    if pd.isna(value):
        return True
    return str(value).strip().lower() in MISSING_WORDS


def has_case_issue(value):
    """檢查文字欄位是否有大小寫混亂。"""
    if is_missing(value):
        return False

    text = str(value).strip()

    if text.replace(".", "", 1).isdigit():
        return False

    letters = [c for c in text if c.isalpha()]
    if not letters:
        return False

    has_upper = any(c.isupper() for c in letters)
    has_lower = any(c.islower() for c in letters)
    is_title_case = text == text.title()

    return has_upper and has_lower and not is_title_case


def is_valid_date(value):
    """檢查日期是否可正確解析。"""
    if is_missing(value):
        return False
    try:
        pd.to_datetime(value, errors="raise")
        return True
    except Exception:
        return False


def is_valid_phone(value):
    """檢查電話號碼格式是否合理。"""
    if is_missing(value):
        return False

    digits = re.sub(r"\D", "", str(value))
    return 8 <= len(digits) <= 15


def is_valid_email(value):
    """檢查電郵格式是否基本正確。"""
    if is_missing(value):
        return False

    pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
    return bool(re.match(pattern, str(value).strip()))


def is_valid_number(value, min_value=None, max_value=None):
    """檢查數值是否有效，並可設定範圍。"""
    if is_missing(value):
        return False

    try:
        number = float(str(value).replace(",", ""))
    except Exception:
        return False

    if min_value is not None and number < min_value:
        return False
    if max_value is not None and number > max_value:
        return False

    return True


# ============================================
# 主要資料清理流程
# ============================================

DATE_COLUMNS = [
    "DOB",
    "Hire_Date",
    "Contract_End_Date",
    "Registration_Date",
    "Last_Login_Timestamp",
    "Last_Purchase_Date",
]

PHONE_COLUMNS = [
    "Phone_Number",
    "Emergency_Contact_Phone",
]

TEXT_COLUMNS = [
    "Gender",
    "Employment_Status",
    "Account_Status",
    "Subscription_Plan",
    "Renewal_Option",
    "Currency",
    "VIP_Tier",
]

NUMERIC_RULES = {
    "Age": (18, 100, "年齡異常或格式錯誤"),
    "Salary_Monthly": (0, 1_000_000, "薪金異常或格式錯誤"),
    "Total_Spend_USD": (0, 100_000, "消費金額異常或格式錯誤"),
    "Discount_Rate": (0, 1, "折扣率異常或格式錯誤"),
}


def check_dataframe(df):
    """檢查 DataFrame，回傳加入問題說明後的資料。"""
    df = df.copy()
    problem_messages = []

    for _, row in df.iterrows():
        row_problems = []

        for column in df.columns:
            if is_missing(row[column]):
                row_problems.append(f"{column} 為空值/NULL/N/A")

        for column in DATE_COLUMNS:
            if column in df.columns and not is_missing(row[column]):
                if not is_valid_date(row[column]):
                    row_problems.append(f"{column} 日期格式不正確")

        for column in PHONE_COLUMNS:
            if column in df.columns and not is_missing(row[column]):
                if not is_valid_phone(row[column]):
                    row_problems.append(f"{column} 電話格式不正確")

        if "Email_Address" in df.columns and not is_missing(row["Email_Address"]):
            if not is_valid_email(row["Email_Address"]):
                row_problems.append("Email_Address 電郵格式不正確")

        for column, (min_value, max_value, message) in NUMERIC_RULES.items():
            if column in df.columns and not is_missing(row[column]):
                if not is_valid_number(row[column], min_value, max_value):
                    row_problems.append(f"{column} {message}")

        for column in TEXT_COLUMNS:
            if column in df.columns and not is_missing(row[column]):
                if has_case_issue(row[column]):
                    row_problems.append(f"{column} 大小寫混亂")

        if row_problems:
            problem_messages.append("；".join(row_problems))
        else:
            problem_messages.append("無問題")

    df["資料問題"] = problem_messages
    df["是否完整"] = np.where(df["資料問題"] == "無問題", "完整", "不完整")

    return df


def create_excel_download(df):
    """建立包含完整資料與不完整資料兩個工作表的 Excel 檔。"""
    clean_df = df[df["是否完整"] == "完整"].copy()
    dirty_df = df[df["是否完整"] == "不完整"].copy()

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        clean_df.to_excel(writer, sheet_name="完整資料", index=False)
        dirty_df.to_excel(writer, sheet_name="不完整資料", index=False)

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]

            for column_cells in worksheet.columns:
                length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                column_letter = get_column_letter(column_cells[0].column)
                worksheet.column_dimensions[column_letter].width = min(length + 2, 50)

    output.seek(0)
    return output, clean_df, dirty_df


# ============================================
# 側邊欄設定
# ============================================

with st.sidebar:
    st.header("⚙️ 上傳與設定")
    uploaded_file = st.file_uploader(
        "請上傳 Excel 檔案",
        type=["xlsx", "xls"],
    )

    sheet_names = []
    excel_file = None

    if uploaded_file is not None:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names

    selected_sheet = None

    if sheet_names:
        selected_sheet = st.selectbox(
            "選擇要檢查的工作表",
            sheet_names,
            index=0,
        )

    st.markdown("---")
    st.caption("支援格式：.xlsx、.xls")

# ============================================
# 主畫面處理
# ============================================

if uploaded_file is None:
    st.info("請先在左側上傳 Excel 檔案。")
    st.stop()

try:
    if selected_sheet:
        df = pd.read_excel(excel_file, sheet_name=selected_sheet)
    else:
        df = pd.read_excel(excel_file)
except Exception as e:
    st.error(f"無法讀取 Excel 檔案：{e}")
    st.stop()

if df.empty:
    st.warning("所選工作表沒有資料，請確認 Excel 內容。")
    st.stop()

st.success(
    f"成功讀取工作表「{selected_sheet}」：共 {df.shape[0]} 列、{df.shape[1]} 欄。"
)

with st.expander("📄 預覽原始資料", expanded=False):
    st.dataframe(df.head(50), use_container_width=True)

if st.button("🧹 開始檢查資料完整性", type="primary"):
    with st.spinner("正在檢查資料，請稍候……"):
        checked_df = check_dataframe(df)

    clean_df = checked_df[checked_df["是否完整"] == "完整"].copy()
    dirty_df = checked_df[checked_df["是否完整"] == "不完整"].copy()

    total_rows = len(checked_df)
    clean_count = len(clean_df)
    dirty_count = len(dirty_df)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("總資料列數", total_rows)

    with col2:
        st.metric("完整資料", clean_count)

    with col3:
        st.metric("不完整資料", dirty_count)

    st.markdown("### 🔍 檢查結果")

    tab1, tab2 = st.tabs(["✅ 完整資料", "⚠️ 不完整資料"])

    with tab1:
        if clean_df.empty:
            st.info("沒有完整資料。")
        else:
            st.dataframe(clean_df.head(100), use_container_width=True)

    with tab2:
        if dirty_df.empty:
            st.info("所有資料均通過檢查。")
        else:
            st.dataframe(dirty_df.head(100), use_container_width=True)

    excel_output, _, _ = create_excel_download(checked_df)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    download_filename = f"整理後資料_{timestamp}.xlsx"

    st.download_button(
        label="⬇️ 下載整理後 Excel 檔",
        data=excel_output,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )