import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

# ==========================================
# 0. 기본 설정 및 구글 API 연동
# ==========================================
st.set_page_config(page_title="포인트 및 상품권 지급 관리", layout="wide")

SPREADSHEET_ID = '1N_LkqUzCgB4xrXT4FQNJ02S2NHnNSbvKTw_hyzIk9pQ'

def authenticate_gspread():
    try:
        # 스트림릿 Secrets에서 GCP 서비스 계정 정보 불러오기
        creds_info = st.secrets["gcp_service_account"]
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        doc = client.open_by_key(SPREADSHEET_ID)
        return doc
    except Exception as e:
        st.error(f"구글 스프레드시트 인증 실패: {e}")
        return None

def overwrite_google_sheet(doc, sheet_name, df):
    try:
        worksheet = doc.worksheet(sheet_name)
        worksheet.clear()  
        df = df.fillna("")
        worksheet.update(df.values.tolist())
        return True
    except Exception as e:
        st.error(f"'{sheet_name}' 시트 업데이트 중 오류 발생: {e}")
        return False

def append_to_google_sheet(doc, sheet_name, df):
    try:
        worksheet = doc.worksheet(sheet_name)
        df = df.fillna("")
        worksheet.append_rows(df.values.tolist())
        return True
    except Exception as e:
        st.error(f"'{sheet_name}' 시트 누적 중 오류 발생: {e}")
        return False

def col2idx(col_str):
    expn = 0
    col_num = 0
    for char in reversed(col_str.upper()):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    return col_num - 1

def load_file_generic(file, skip_rows=0):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, header=None, dtype=str, skiprows=skip_rows)
    else:
        df = pd.read_excel(file, header=None, dtype=str, skiprows=skip_rows)
    return df

st.title("🎁 포인트 및 상품권 지급 관리 시스템")

menu = st.sidebar.radio("메뉴 선택", ["1. 데이터 업로드 및 관리", "2. 포인트 지급 대상 조회", "3. 상품권 지급 대상 조회"])
doc = authenticate_gspread()

# ==========================================
# 1. 데이터 업로드 및 관리 페이지
# ==========================================
if menu == "1. 데이터 업로드 및 관리":
    st.header("📂 데이터 업로드 및 구글 시트 연동")
    if doc is None: st.stop()

    # 1. 경리나라 수납 데이터
    st.subheader("1. 경리나라 수납 데이터 (첫 행 제외 및 13개열 추출)")
    receipt_file = st.file_uploader("경리나라 수납 파일 업로드", type=['xlsx', 'xls', 'csv'], key="receipt")
    if receipt_file:
        df_receipt = load_file_generic(receipt_file, skip_rows=0)
        if not df_receipt.empty and ('사업자' in str(df_receipt.iloc[0, 0])):
            df_receipt = df_receipt.iloc[1:].reset_index(drop=True)
        df_receipt_final = df_receipt.iloc[:, :13].copy()
        for c in range(df_receipt_final.shape[1], 13): df_receipt_final[c] = ""
        st.dataframe(df_receipt_final.head(3))
        if st.button("경리나라 수납 데이터 반영"):
            overwrite_google_sheet(doc, "경리나라 수납", df_receipt_final)
            st.success("✅ 반영 완료")

    st.divider()

    # 2. 추천 데이터 (3행 제외, M열 반영)
    st.subheader("2. 추천 데이터 (3행 제외, M열 → 추천일 반영)")
    referral_file = st.file_uploader("추천 파일 업로드", type=['xlsx', 'xls', 'csv'], key="referral")
    if referral_file:
        df_raw = load_file_generic(referral_file, skip_rows=3)
        if not df_raw.empty:
            target_indices = [0, 1, 2, 3, 12, col2idx('AT'), col2idx('AU'), col2idx('AV')]
            df_referral_final = df_raw.iloc[:, [i for i in target_indices if i < df_raw.shape[1]]].copy()
            df_referral_final.iloc[:, 0] = df_referral_final.iloc[:, 0].astype(str).str.replace('-', '', regex=False)
            for c in range(df_referral_final.shape[1], 8): df_referral_final[c] = ""
            st.dataframe(df_referral_final.head(5))
            if st.button("추천 데이터 누적 추가"):
                append_to_google_sheet(doc, "추천", df_referral_final)
                st.success("✅ 추천 데이터 누적 완료")

    st.divider()

    # 3. 위멤버스 가입 여부 데이터
    st.subheader("3. 위멤버스 가입 여부 데이터 (D, G, BQ열 추출)")
    wemembers_file = st.file_uploader("위멤버스 가입 여부 파일 업로드", type=['xlsx', 'xls', 'csv'], key="wemembers")
    if wemembers_file:
        df_we = load_file_generic(wemembers_file, skip_rows=0)
        if not df_we.empty and ('사업자' in str(df_we.iloc[0, 0])):
            df_we = df_we.iloc[1:].reset_index(drop=True)
        if df_we.shape[1] <= 5: target_indices = [0, 1, 2]
        else: target_indices = [col2idx('D'), col2idx('G'), col2idx('BQ')]
        df_we_final = df_we.iloc[:, [i for i in target_indices if i < df_we.shape[1]]].copy()
        if df_we_final.shape[1] > 0:
            df_we_final.iloc[:, 0] = df_we_final.iloc[:, 0].astype(str).str.replace('-', '', regex=False)
        for c in range(df_we_final.shape[1], 3): df_we_final[c] = ""
        st.dataframe(df_we_final.head(3))
        if st.button("위멤버스 데이터 시트에 반영"):
            overwrite_google_sheet(doc, "위멤버스 가입 여부", df_we_final)
            st.success("✅ 반영 완료")

# ==========================================
# 2. 포인트 지급 대상 조회 페이지
# ==========================================
elif menu == "2. 포인트 지급 대상 조회":
    st.header("🎯 포인트 지급 내역 산출")
    if doc:
        with st.spinner("데이터 분석 중..."):
            receipt_data = doc.worksheet("경리나라 수납").get_all_values()
            referral_data = doc.worksheet("추천").get_all_values()
            wemembers_data = doc.worksheet("위멤버스 가입 여부").get_all_values()
            rate_data