import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
import tempfile
import re
import io

st.title("📌 Google Sheet 자동정리 → Excel 변환기")

# ✅ 고정된 Google Sheet URL
sheet_url = "https://docs.google.com/spreadsheets/d/1qy0umMpL50qZ_kjSzWbj4iYH-cnm-GBtJ7gYyPAVT_A/export?format=csv"

if st.button("📥 변환 실행"):
    st.write("✅ 구글시트 불러오는 중...")

    df = pd.read_csv(sheet_url, header=None)

    # ===== 2. 1행 삭제 후 2행만 헤더로 유지 =====
    header_row = df.iloc[1:2]
    data_rows = df.iloc[2:].copy()

    # ===== 3. 첫 번째 열 결측이 아닌 행 삭제 =====
    data_rows = data_rows[data_rows[0].isna()]

    # ===== 4. 결측 채우기용 연월일+순번 생성 =====
    today = datetime.today().strftime("%Y%m%d")
    count = len(data_rows)
    fill_values = [f"{today}{num:02d}" for num in range(1, count+1)]
    data_rows[0] = fill_values

    # ===== ✅ 5. F열(6번째 열) 핸드폰번호 정규화 =====
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

    # ===== 6. 다시 합치기 =====
    final_df = pd.concat([header_row, data_rows], ignore_index=True)

    # ===== 7. 엑셀 파일 생성 (임시 저장) =====
    output_filename = f"filled_sheet_{today}.xlsx"
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    final_df.to_excel(temp_file.name, index=False)

    # ===== 8. openpyxl 서식 적용 =====
    wb = load_workbook(temp_file.name)
    ws = wb.active

    # (1) 1번 행 삭제
    ws.delete_rows(1)

    # (2) 테두리 적용
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border

    # (3) 열 너비 자동 조정
    def visual_len(s: str) -> int:
        if s is None:
            return 0
        s = str(s)
        return sum(2 if ord(ch) > 255 else 1 for ch in s)

    min_width, max_width = 8, 80
    for col_idx in range(1, ws.max_column + 1):
        max_len = max(visual_len(ws.cell(row=row_idx, column=col_idx).value)
                      for row_idx in range(1, ws.max_row + 1))
        ws.column_dimensions[get_column_letter(col_idx)].width = max(min_width, min(max_width, max_len + 2))

    wb.save(temp_file.name)

    # ===== 9. 다운로드 버튼 제공 =====
    with open(temp_file.name, "rb") as f:
        st.download_button(
            label="📤 엑셀파일 다운로드",
            data=f,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success("✅ 변환 완료! 아래 버튼을 눌러 다운로드하세요.")
