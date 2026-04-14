import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

# ==========================================
# 0. 기본 설정 및 구글 API 연동 (Secrets 적용)
# ==========================================
st.set_page_config(page_title="포인트 및 상품권 지급 관리", layout="wide")

SPREADSHEET_ID = '1N_LkqUzCgB4xrXT4FQNJ02S2NHnNSbvKTw_hyzIk9pQ'

def authenticate_gspread():
    try:
        # Streamlit Cloud의 Secrets에서 인증 정보 로드
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
        if not df.empty:
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
    """엑셀 열 문자를 인덱스 숫자로 변환 (A->0, AA->26 등)"""
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

    # 1-1. 경리나라 수납 데이터 (첫 행 제외 & 특정 열 추출)
    st.subheader("1. 경리나라 수납 데이터 (첫 행 제외 및 G, I, W, X, AA, AL, AM열)")
    receipt_file = st.file_uploader("경리나라 수납 파일 업로드", type=['xlsx', 'xls', 'csv'], key="receipt")
    if receipt_file:
        # 첫 행 제외 (skiprows=1)
        df_receipt_raw = load_file_generic(receipt_file, skip_rows=1)
        if not df_receipt_raw.empty:
            try:
                target_cols = ['G', 'I', 'W', 'X', 'AA', 'AL', 'AM']
                target_indices = [col2idx(c) for c in target_cols]
                # 요청하신 열만 추출
                df_receipt_final = df_receipt_raw.iloc[:, [i for i in target_indices if i < df_receipt_raw.shape[1]]].copy()
                # G열(추출 후 첫 번째 열) 하이픈 제거
                if df_receipt_final.shape[1] > 0:
                    df_receipt_final.iloc[:, 0] = df_receipt_final.iloc[:, 0].astype(str).str.replace('-', '', regex=False)
                
                st.dataframe(df_receipt_final.head(3))
                if st.button("경리나라 수납 데이터 반영"):
                    overwrite_google_sheet(doc, "경리나라 수납", df_receipt_final)
                    st.success("✅ 경리나라 수납 데이터 반영 완료")
            except Exception as e:
                st.error(f"데이터 추출 중 오류: {e}")

    st.divider()

    # 1-2. 추천 데이터 (3행 제외, C열->A열 복사, M열 반영)
    st.subheader("2. 추천 데이터 (3행 제외, C열을 A열로 복사)")
    referral_file = st.file_uploader("추천 파일 업로드", type=['xlsx', 'xls', 'csv'], key="referral")
    if referral_file:
        df_raw = load_file_generic(referral_file, skip_rows=3)
        if not df_raw.empty:
            if df_raw.shape[1] > 2:
                df_raw.iloc[:, 0] = df_raw.iloc[:, 2].astype(str).str.replace('-', '', regex=False)
            
            target_indices = [0, 1, 2, 3, 12, col2idx('AT'), col2idx('AU'), col2idx('AV')]
            df_referral_final = df_raw.iloc[:, [i for i in target_indices if i < df_raw.shape[1]]].copy()
            for c in range(df_referral_final.shape[1], 8): df_referral_final[c] = ""
            
            st.dataframe(df_referral_final.head(3))
            if st.button("추천 데이터 누적 추가"):
                append_to_google_sheet(doc, "추천", df_referral_final)
                st.success("✅ 추천 데이터 누적 완료")

    st.divider()

    # 1-3. 위멤버스 가입 여부 데이터
    st.subheader("3. 위멤버스 가입 여부 데이터 (D, G, BQ열 추출)")
    wemembers_file = st.file_uploader("위멤버스 가입 여부 파일 업로드", type=['xlsx', 'xls', 'csv'], key="wemembers")
    if wemembers_file:
        df_we = load_file_generic(wemembers_file, skip_rows=0)
        if not df_we.empty and '사업자' in str(df_we.iloc[0, 0]):
            df_we = df_we.iloc[1:].reset_index(drop=True)
            
        target_indices = [col2idx('D'), col2idx('G'), col2idx('BQ')]
        df_we_final = df_we.iloc[:, [i for i in target_indices if i < df_we.shape[1]]].copy()
        if df_we_final.shape[1] > 0:
            df_we_final.iloc[:, 0] = df_we_final.iloc[:, 0].astype(str).str.replace('-', '', regex=False)
        for c in range(df_we_final.shape[1], 3): df_we_final[c] = ""
        
        st.dataframe(df_we_final.head(3))
        if st.button("위멤버스 데이터 시트에 반영"):
            overwrite_google_sheet(doc, "위멤버스 가입 여부", df_we_final)
            st.success("✅ 위멤버스 반영 완료")

# ==========================================
# 2. 포인트 지급 대상 조회 및 보고서
# ==========================================
elif menu == "2. 포인트 지급 대상 조회":
    st.header("🎯 포인트 지급 내역 산출")
    if doc:
        with st.spinner("데이터 분석 중..."):
            receipt_data = doc.worksheet("경리나라 수납").get_all_values()
            referral_data = doc.worksheet("추천").get_all_values()
            wemembers_data = doc.worksheet("위멤버스 가입 여부").get_all_values()
            rate_data = doc.worksheet("적립율").get_all_values()
            
            if not receipt_data or not referral_data:
                st.warning("데이터가 부족합니다. 먼저 업로드해주세요.")
                st.stop()

            r_df = pd.DataFrame(receipt_data)
            ref_df = pd.DataFrame(referral_data)
            
            # 딕셔너리 구성 (하이픈 제거 매칭)
            rate_dict = {}
            for row in rate_data:
                if len(row) >= 2:
                    biz = str(row[0]).replace('-', '').strip()
                    try: rate_dict[biz] = float(str(row[1]).replace('%',''))
                    except: continue

            we_dict = {str(row[0]).replace('-', '').strip(): {'제품명': row[1], '비고': row[2]} 
                       for row in wemembers_data if len(row) >= 3}

        try:
            r_df.columns = [f"r_{i}" for i in range(len(r_df.columns))]
            ref_df.columns = [f"ref_{i}" for i in range(len(ref_df.columns))]
            
            # 필터링 및 매칭 (수납 G열 -> r_0, 추천 A열 -> ref_0)
            merged_df = pd.merge(r_df, ref_df, left_on="r_0", right_on="ref_0")

            results = []
            for _, row in merged_df.iterrows():
                rec_biz_raw = str(row.get("ref_6", ""))
                rec_biz_clean = rec_biz_raw.replace('-', '').strip()
                
                # 기한 필터링 (20251231)
                rec_date = str(row.get("ref_4", ""))
                if ''.join(filter(str.isdigit, rec_date)) > "20251231": continue

                we_info = we_dict.get(rec_biz_clean, {'제품명': '', '비고': ''})
                if not str(we_info['제품명']).strip(): continue

                # 포인트 계산 (수납 데이터 구조에 따라 r_인덱스 확인 필요)
                base_amt = pd.to_numeric(str(row.get("r_3", "0")).replace(',',''), errors='coerce') or 0
                rate = rate_dict.get(rec_biz_clean, 0.03)
                
                results.append({
                    "수납_사업자": row.get("r_0"),
                    "추천일": rec_date,
                    "위멤버스_제품명(G)": we_info['제품명'],
                    "위멤버스_비고(BQ)": we_info['비고'],
                    "추천자 회사명": row.get("ref_5"),
                    "추천자 사업자번호": rec_biz_raw,
                    "최종 지급포인트": base_amt * rate
                })

            final_df = pd.DataFrame(results)
            if not final_df.empty:
                st.dataframe(final_df, use_container_width=True)
                st.divider()
                st.subheader("📋 포인트 지급 내역 합계 보고서")
                summary = final_df.groupby(["위멤버스_비고(BQ)", "추천자 회사명", "추천자 사업자번호"])["최종 지급포인트"].sum().reset_index()
                st.dataframe(summary, use_container_width=True)
            else:
                st.info("조건을 충족하는 대상이 없습니다.")
        except Exception as e:
            st.error(f"오류 발생: {e}")

# ==========================================
# 3. 상품권 지급 대상 조회
# ==========================================
elif menu == "3. 상품권 지급 대상 조회":
    st.header("🎟️ 상품권 지급 대상 조회")
    if doc:
        receipt_data = doc.worksheet("경리나라 수납").get_all_values()
        referral_data = doc.worksheet("추천").get_all_values()
        wemembers_data = doc.worksheet("위멤버스 가입 여부").get_all_values()
        
        r_df = pd.DataFrame(receipt_data)
        ref_df = pd.DataFrame(referral_data)
        r_df.columns = [f"r_{i}" for i in range(len(r_df.columns))]
        ref_df.columns = [f"ref_{i}" for i in range(len(ref_df.columns))]
        
        # 1회차 수납 필터링 (원본 열 순서에 따라 r_인덱스 주의)
        # 예시로 r_5가 입금횟수라고 가정
        try:
            r_df["r_5"] = pd.to_numeric(r_df["r_5"], errors='coerce')
            filtered_r = r_df[r_df["r_5"] == 1].copy()
            merged_df = pd.merge(filtered_r, ref_df, left_on="r_0", right_on="ref_0")
            st.dataframe(merged_df, use_container_width=True)
        except:
            st.warning("데이터 구조를 확인해주세요.")