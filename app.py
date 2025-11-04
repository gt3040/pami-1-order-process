import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
import tempfile
import re

st.title("📌 파미-1 주문서 생성기")

# ✅ URL은 secrets.toml 또는 Streamlit Cloud Secrets에서 불러옴
sheet_url = st.secrets["GOOGLE_SHEET_URL"]

@st.cache_data(show_spinner=True)
def process_file(sheet_url):
    df = pd.read_csv(sheet_url, header=None)

    # 1행 삭제 후 2행만 헤더로 유지
    header_row = df.iloc[1:2]
    data_rows = df.iloc[2:].copy()

    # 첫 번째 열 결측이 아닌 행 삭제
    data_rows = data_rows[data_rows[0].isna()]

    # 결측 채우기 → 연월일 + 2자리 순번
    today = datetime.today().strftime("%Y%m%d")
    count = len(data_rows)
    fill_values = [f"{today}{num:02d}" for num in range(1, count + 1)]
    data_rows[0] = fill_values

    # ✅ 전화번호 정규화
    def normalize_phone(phone):
        if pd.isna(phone):
            return ""
        phone = str(phone).replace("-", "").replace(" ", "").replace("+82", "0")
        if phone.startswith("82") and len(phone) >= 11:
            phone = "0" + phone[2:]
        if len(phone) == 10:
            phone = "0" + phone
        if len(phone) == 11:
            return f"{phone[0:3]}-{phone[3:7]}-{phone[7:11]}"
        return phone

    data_rows[5] = data_rows[5].apply(normalize_phone)

    # 다시 합치기
    final_df = pd.concat([header_row, data_rows], ignore_index=True)

    # 임시 엑셀 저장
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    final_df.to_excel(temp_file.name, index=False)

    # openpyxl 스타일 적용 (테두리 + 열 너비 자동조정)
    wb = load_workbook(temp_file.name)
    ws = wb.active
    ws.delete_rows(1)  # 첫줄 삭제

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border

    def visual_len(s: str) -> int:
        if s is None:
            return 0
        return sum(2 if ord(ch) > 255 else 1 for ch in str(s))

    min_width, max_width = 8, 80
    for col_idx in range(1, ws.max_column + 1):
        max_len = max(visual_len(ws.cell(row=row_idx, column=col_idx).value)
                      for row_idx in range(1, ws.max_row + 1))
        ws.column_dimensions[get_column_letter(col_idx)].width = max(min_width, min(max_width, max_len + 2))

    wb.save(temp_file.name)

    return temp_file.name, f"filled_sheet_{today}.xlsx"


# ✅ 다운로드 버튼 1개 → 클릭 즉시 변환 + 다운로드
file_path, file_name = process_file(sheet_url)
with open(file_path, "rb") as f:
    st.download_button(
        label="📥 주문서 다운로드",
        data=f,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.success("✅ 다운로드 버튼을 눌러 저장 후 전송!")

