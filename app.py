import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
import re

# ===== 1. 구글 시트 불러오기 =====
sheet_url = "https://docs.google.com/spreadsheets/d/1qy0umMpL50qZ_kjSzWbj4iYH-cnm-GBtJ7gYyPAVT_A/export?format=csv"
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

# ===== ✅ 5. F열(6번째 열) 핸드폰번호 정규화 함수 =====
def normalize_phone(phone):
    if pd.isna(phone):
        return ""
    phone = str(phone).replace("-", "").replace(" ", "").replace("+82", "0")
    if phone.startswith("82") and len(phone) >= 11:
        phone = "0" + phone[2:]
    if len(phone) == 10:  # 10자리 → 010으로 앞자리 보정
        phone = "0" + phone
    if len(phone) == 11:
        return f"{phone[0:3]}-{phone[3:7]}-{phone[7:11]}"
    return phone  # 규칙 밖의 값은 원본 유지

# ✅ F열(인덱스 5)에 적용
data_rows[5] = data_rows[5].apply(normalize_phone)

# ===== 6. 다시 합치기 =====
final_df = pd.concat([header_row, data_rows], ignore_index=True)

# ===== 7. 엑셀 파일 저장 =====
output_filename = f"filled_sheet_{today}.xlsx"
final_df.to_excel(output_filename, index=False)

# ===== 8. openpyxl 서식 적용 =====
wb = load_workbook(output_filename)
ws = wb.active

# (1) 1번 행 삭제
ws.delete_rows(1)

# (2) 테두리 적용
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                     top=Side(style='thin'), bottom=Side(style='thin'))
for row in ws.iter_rows():
    for cell in row:
        cell.border = thin_border

# (3) 열 너비 자동 맞춤
def visual_len(s: str) -> int:
    if s is None:
        return 0
    s = str(s)
    length = 0
    for ch in s:
        length += 1 if ord(ch) <= 255 else 2
    return length

min_width, max_width = 8, 80
for col_idx in range(1, ws.max_column + 1):
    max_len = 0
    for row_idx in range(1, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=col_idx).value
        max_len = max(max_len, visual_len(val))
    ws.column_dimensions[get_column_letter(col_idx)].width = max(min_width, min(max_width, max_len + 2))

wb.save(output_filename)
print(f"✅ 저장 완료 → {output_filename}")

# ===== 9. (Colab 전용) 자동 다운로드 =====
try:
    from google.colab import files
    files.download(output_filename)
    print("✅ 다운로드 시작됨 (Colab 환경)")
except:
    print("⚠️ 로컬 Jupyter 환경이면 해당 파일이 현재 폴더에 저장되었습니다.")
