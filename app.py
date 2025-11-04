import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import re
import io

# ✅ Streamlit UI
st.title("📌 Google Sheet → 데이터 정리 → Excel 다운로드 자동화 도구")
st.write("A열 결측값 자동 ID 생성 + F열 전화번호 형식 정리 + 결과 엑셀 다운로드")

# ✅ Google Sheet 설정 (고정)
SHEET_URL = "https://docs.google.com/spreadsheets/d/1qy0umMpL50qZ_kjSzWbj4iYH-cnm-GBtJ7gYyPAVT_A/edit?usp=sharing"
SHEET_NAME = "Sheet1"

# ✅ 전화번호 정규화 함수
def format_phone(num):
    if not isinstance(num, str):
        return num
    digits = re.sub(r'[^0-9]', '', num)

    if digits.startswith("01") and len(digits) in (10, 11):
        if len(digits) == 10:
            return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
        if len(digits) == 11:
            return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    return num


# ✅ 버튼 실행
if st.button("✅ 실행하기 (시트 불러와서 처리 & 엑셀 다운로드)"):
    st.write("🔄 Google Sheet 불러오는 중... 잠시만 기다려주세요!")

    # ✅ Google API 인증
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
    client = gspread.authorize(creds)

    # ✅ 시트 불러오기
    sheet = client.open_by_url(SHEET_URL).worksheet(SHEET_NAME)
    data = sheet.get_all_values()
    df = pd.DataFrame(data)
    df.columns = df.iloc[0]
    df = df[1:]
    df.reset_index(drop=True, inplace=True)

    # ✅ A열 결측 행만 남김
    first_col = df.columns[0]
    df_missing = df[df[first_col].isna() | (df[first_col] == "")].copy()

    # ✅ 날짜+2자리 ID 생성
    today_str = datetime.now().strftime("%Y%m%d")
    df_missing[first_col] = [
        f"{today_str}{str(i+1).zfill(2)}" for i in range(len(df_missing))
    ]

    # ✅ 전화번호 정규화 (F열 고정)
    phone_col = "F"
    if phone_col in df_missing.columns:
        df_missing[phone_col] = df_missing[phone_col].apply(format_phone)

    # ✅ 엑셀로 변환 (메모리 버퍼로 저장)
    output = io.BytesIO()
    df_missing.to_excel(output, index=False)
    output.seek(0)

    # ✅ 다운로드 버튼 생성
    st.success("✅ 처리 완료! 아래 버튼을 눌러 엑셀 파일을 다운로드하세요.")
    st.download_button(
        label="📥 엑셀 다운로드",
        data=output,
        file_name=f"processed_{today_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

