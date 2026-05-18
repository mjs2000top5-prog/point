import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

# ==========================================
# 0. 기본 유틸리티 함수 (가장 먼저 정의)
# ==========================================
def col2idx(col_str):
    """엑셀 열 문자를 인덱스 숫자로 변환 (A->0, E->4 등)"""
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

# ==========================================
# 1. 기본 설정 및 구글 API 연동 (Secrets 적용)
# ==========================================
st.set_page_config(page_title="포인트 및 상품권 지급 관리", layout="wide")

# 구글 스프레드시트 ID
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

# 시트 관리 공통 함수
def overwrite_google_sheet(doc, sheet_name, df):
    try:
        worksheet = doc.worksheet(sheet_name)
        worksheet.clear()
        df = df.fillna("")
        if not df.empty:
            worksheet.update(df.values.tolist())
        return True
    except Exception as e:
        st.error(f"'{sheet_name}' 업데이트 오류: {e}")
        return False

def append_to_google_sheet(doc, sheet_name, df):
    try:
        worksheet = doc.worksheet(sheet_name)
        df = df.fillna("")
        worksheet.append_rows(df.values.tolist())
        return True
    except Exception as e:
        st.error(f"'{sheet_name}' 누적 오류: {e}")
        return False

st.title("🎁 포인트 및 상품권 지급 관리 시스템")

menu = st.sidebar.radio("메뉴 선택", ["1. 데이터 업로드 및 관리", "2. 포인트 지급 대상 조회", "3. 상품권 지급 대상 조회"])
doc = authenticate_gspread()

# ==========================================
# 2. 데이터 업로드 및 관리 페이지
# ==========================================
if menu == "1. 데이터 업로드 및 관리":
    st.header("📂 데이터 업로드 및 관리")
    if doc is None: st.stop()

    # 1-1. 경리나라 수납 데이터
    st.subheader("1. 경리나라 수납 데이터 업로드 (순서: G, I, W, X, AA, AL, AM + 계산용 C, E)")
    if st.button("🗑️ 기존 수납 데이터 삭제", key="clear_r"):
        worksheet = doc.worksheet("경리나라 수납")
        worksheet.clear()
        st.success("수납 데이터 삭제 완료")
    
    receipt_file = st.file_uploader("수납 파일 업로드 (xlsx, csv)", type=['xlsx', 'xls', 'csv'], key="u1")
    if receipt_file:
        df_raw = load_file_generic(receipt_file, skip_rows=1)
        if not df_raw.empty:
            # 요청하신 순서대로 열 인덱스 정의
            target_cols = ['G', 'I', 'W', 'X', 'AA', 'AL', 'AM', 'C', 'E']
            target_indices = [col2idx(c) for c in target_cols]
            
            # 파일 내 존재하는 열만 필터링하여 추출
            available_indices = [i for i in target_indices if i < df_raw.shape[1]]
            df_receipt = df_raw.iloc[:, available_indices].copy()

            # 텍스트 클리닝 (신규, 부산 등 패턴 제거)
            def clean_text_patterns(val):
                if not isinstance(val, str): return val
                patterns = [r'\(신규\)', r'\(부산\)', r'외\s?\d+명', r'\(new\)', r'\(Busan\)', r'plus\s?\d+\s?people']
                for p in patterns:
                    val = re.sub(p, '', val)
                return val.strip()
            
            df_receipt = df_receipt.map(clean_text_patterns)

            # 사업자번호(첫 번째 열인 G열) 하이픈 제거
            df_receipt.iloc[:, 0] = df_receipt.iloc[:, 0].astype(str).str.replace('-', '', regex=False)

            st.write("📊 업로드 데이터 미리보기 (G, I, W, X, AA, AL, AM, C, E 순서):")
            st.dataframe(df_receipt.head(3))
            
            if st.button("경리나라 수납 시트 반영"):
                overwrite_google_sheet(doc, "경리나라 수납", df_receipt)
                st.success("지정한 순서로 데이터 반영을 완료했습니다.")

    st.divider()

    # 1-2. 추천 데이터
    st.subheader("2. 추천 데이터 업로드 (3행 제외)")
    referral_file = st.file_uploader("추천 파일 업로드", type=['xlsx', 'xls', 'csv'], key="u2")
    if referral_file:
        df_ref_raw = load_file_generic(referral_file, skip_rows=3)
        if not df_ref_raw.empty:
            df_ref_raw.iloc[:, 0] = df_ref_raw.iloc[:, 2].astype(str).str.replace('-', '', regex=False)
            t_idxs = [0, 1, 2, 3, 12, 45, 46, 47]
            df_ref_final = df_ref_raw.iloc[:, [i for i in t_idxs if i < df_ref_raw.shape[1]]].copy()
            st.dataframe(df_ref_final.head(3))
            if st.button("추천 데이터 누적 추가"):
                append_to_google_sheet(doc, "추천", df_ref_final)
                st.success("추천 데이터 누적 완료")

    st.divider()

    # 1-3. 위멤버스 가입 여부
    st.subheader("3. 위멤버스 가입 여부 업로드")
    if st.button("🗑️ 기존 위멤버스 데이터 삭제", key="clear_we"):
        worksheet = doc.worksheet("위멤버스 가입 여부")
        worksheet.clear()
        st.success("위멤버스 데이터 삭제 완료")
    
    we_file = st.file_uploader("위멤버스 파일 업로드", type=['xlsx', 'xls', 'csv'], key="u3")
    if we_file:
        df_we_raw = load_file_generic(we_file, skip_rows=0)
        if not df_we_raw.empty:
            df_we_raw.iloc[:, 3] = df_we_raw.iloc[:, 3].astype(str).str.replace('-', '', regex=False)
            target_cols = [3, 6, 69]
            df_we_final = df_we_raw.iloc[:, [i for i in target_cols if i < df_we_raw.shape[1]]].copy()
            st.dataframe(df_we_final.head(3))
            if st.button("위멤버스 시트 반영"):
                overwrite_google_sheet(doc, "위멤버스 가입 여부", df_we_final)
                st.success("위멤버스 데이터 반영 완료")

# ==========================================
# 데이터 처리 로직 (변경된 열 순서 반영)
# ==========================================
def get_processed_data(doc, filter_count_one=False):
    try:
        r_values = doc.worksheet("경리나라 수납").get_all_values()
        ref_values = doc.worksheet("추천").get_all_values()
        we_values = doc.worksheet("위멤버스 가입 여부").get_all_values()
        rate_values = doc.worksheet("적립율").get_all_values()
        
        if not r_values or not ref_values: return pd.DataFrame()

        r_df = pd.DataFrame(r_values).fillna("")
        ref_df = pd.DataFrame(ref_values).fillna("")
        
        # 적립율/위멤버스 매핑 데이터 준비 및 텍스트 데이터 예외 처리
        rate_dict = {}
        for row in rate_values:
            if len(row) >= 2:
                biz = str(row[0]).replace('-', '').strip()
                try:
                    val = float(str(row[1]).replace('%','').strip())
                    rate_dict[biz] = val/100 if val > 1 else val
                except ValueError:
                    continue
                    
        we_dict = {str(row[0]).replace('-', '').strip(): {'제품명': row[1], '비고': row[2] if len(row)>2 else ''} for row in we_values if len(row) > 0}

        results = []
        for _, ref_row in ref_df.iterrows():
            target_biz = str(ref_row[0]).replace('-', '').strip()
            # 시트의 0번 인덱스가 G열(사업자번호)
            matched_rows = r_df[r_df[0] == target_biz]
            
            for _, r_row in matched_rows.iterrows():
                # 인덱스 구조: 0:G, 1:I, 2:W, 3:X, 4:AA, 5:AL, 6:AM, 7:C, 8:E
                
                # 수납횟수 (AL열 -> 인덱스 5) - 헤더가 섞여 있어도 튕기지 않게 텍스트 예외 처리
                raw_count = str(r_row[5]).strip()
                pay_count = pd.to_numeric(raw_count, errors='coerce')
                if pd.isna(pay_count): continue
                if filter_count_one:
                    if pay_count != 1: continue
                else:
                    if pay_count <= 0 or pay_count >= 60: continue

                # 청구금액 (E열 -> 인덱스 8) - 헤더가 섞여 있어도 튕기지 않게 텍스트 예외 처리
                raw_bill = str(r_row[8]).replace(',','').strip()
                bill_amt = pd.to_numeric(raw_bill, errors='coerce')
                if pd.isna(bill_amt) or bill_amt < 40000: continue

                # 설치일 (C열 -> 인덱스 7)
                install_date_raw = str(r_row[7])
                install_clean = ''.join(filter(str.isdigit, install_date_raw))
                if install_clean and len(install_clean) >= 8 and int(install_clean[:8]) >= 20260101: continue
                
                # 추천일자 필터링
                rec_date = str(ref_row[4])
                if ''.join(filter(str.isdigit, rec_date)) > "20251231": continue

                rec_biz_raw = str(ref_row[6])
                rec_biz_clean = rec_biz_raw.replace('-', '').strip()
                we_info = we_dict.get(rec_biz_clean, {'제품명': '', '비고': ''})
                if not we_info['제품명'] or "위멤버스 베이직" in str(we_info['제품명']): continue

                rate = rate_dict.get(rec_biz_clean, 0.03)
                final_p = bill_amt * rate
                
                results.append({
                    "설치일(C)": install_date_raw,
                    "수납횟수(AL)": int(pay_count),
                    "수납사명": r_row[1], # I열
                    "수납사업자번호": target_biz,
                    "추천자회사": ref_row[5],
                    "추천자사업자": rec_biz_raw,
                    "위멤버스_제품": we_info['제품명'],
                    "적립율": rate,
                    "청구금액(E)": int(bill_amt),
                    "최종지급포인트": int(final_p)
                })
        return pd.DataFrame(results)
    except Exception as e:
        st.error(f"데이터 처리 오류: {e}")
        return pd.DataFrame()

# ==========================================
# 3. 메뉴별 화면 표시 로직
# ==========================================
if menu == "2. 포인트 지급 대상 조회":
    st.header("🎯 포인트 지급 내역 산출")
    if doc:
        with st.spinner("데이터 매칭 및 산출 중..."):
            final_df = get_processed_data(doc, filter_count_one=False)
            if not final_df.empty:
                st.dataframe(final_df, use_container_width=True)
                st.download_button("📥 상세 내역 다운로드 (CSV)", final_df.to_csv(index=False).encode('utf-8-sig'), "point_detail.csv")
                
                st.divider()
                st.subheader("📋 추천자별 합계 보고서")
                summary = final_df.groupby(["추천자회사", "추천자사업자"])["최종지급포인트"].sum().reset_index()
                st.dataframe(summary, use_container_width=True)
                st.download_button("📥 합계 보고서 다운로드 (CSV)", summary.to_csv(index=False).encode('utf-8-sig'), "point_summary.csv")
            else: 
                st.info("대상 데이터가 없습니다.")

elif menu == "3. 상품권 지급 대상 조회":
    st.header("🎟️ 상품권 지급 대상 (수납 1회)")
    if doc:
        with st.spinner("대상자 필터링 중..."):
            gift_df = get_processed_data(doc, filter_count_one=True)
            if not gift_df.empty:
                st.success(f"상품권 지급 대상: 총 {len(gift_df)}건")
                st.dataframe(gift_df, use_container_width=True)
                st.download_button("📥 상품권 지급 대상 다운로드 (CSV)", gift_df.to_csv(index=False).encode('utf-8-sig'), "gift_card.csv")
            else: 
                st.info("대상 데이터가 없습니다.")
