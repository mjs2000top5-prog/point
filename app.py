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
        # 데이터가 있는 경우만 업데이트
        if not df.empty:
            worksheet.update(df.values.tolist())
        return True
    except Exception as e:
        st.error(f"'{sheet_name}' 시트 업데이트 중 오류 발생: {e}")
        return False

def col2idx(col_str):
    """엑셀 열 문자(A, B, C...)를 인덱스 숫자(0, 1, 2...)로 변환"""
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

    # 1-1. 경리나라 수납 데이터 (요청사항 반영)
    st.subheader("1. 경리나라 수납 데이터 (첫 행 제외 & G, I, W, X, AA, AL, AM열 추출)")
    receipt_file = st.file_uploader("경리나라 수납 파일 업로드", type=['xlsx', 'xls', 'csv'], key="receipt")
    
    if receipt_file:
        # 첫 행 제외하고 읽기 (skiprows=1)
        df_receipt_raw = load_file_generic(receipt_file, skip_rows=1)
        
        if not df_receipt_raw.empty:
            try:
                # 추출할 열 목록 지정
                target_cols = ['G', 'I', 'W', 'X', 'AA', 'AL', 'AM']
                target_indices = [col2idx(c) for c in target_cols]
                
                # 원본 파일에 해당 열들이 존재하는지 확인하며 추출
                df_receipt_final = df_receipt_raw.iloc[:, [i for i in target_indices if i < df_receipt_raw.shape[1]]].copy()
                
                # G열(추출된 데이터의 첫 번째 열)에서 하이픈(-) 제거
                if df_receipt_final.shape[1] > 0:
                    df_receipt_final.iloc[:, 0] = df_receipt_final.iloc[:, 0].astype(str).str.replace('-', '', regex=False)
                
                st.write("업로드될 데이터 미리보기 (G, I, W, X, AA, AL, AM):")
                st.dataframe(df_receipt_final.head(5))
                
                if st.button("경리나라 수납 데이터 반영"):
                    with st.spinner("구글 시트에 반영 중..."):
                        if overwrite_google_sheet(doc, "경리나라 수납", df_receipt_final):
                            st.success("✅ 경리나라 수납 데이터 반영 완료 (요청하신 열만 추출됨)")
            except Exception as e:
                st.error(f"데이터 처리 중 오류 발생: {e}")

    st.divider()
    # (이하 추천 데이터, 위멤버스 데이터 등 기존 업로드 로직 유지...)