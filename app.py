import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ==========================================
# 0. 기본 설정 및 구글 API 연동 (Secrets)
# ==========================================
st.set_page_config(page_title="포인트 및 상품권 지급 관리", layout="wide")

# 구글 스프레드시트 ID
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

def clear_google_sheet(doc, sheet_name):
    try:
        worksheet = doc.worksheet(sheet_name)
        worksheet.clear()
        return True
    except Exception as e:
        st.error(f"'{sheet_name}' 시트 삭제 오류: {e}")
        return False

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

st.title("🎁 포인트 및 상품권 지급 관리 시스템")

menu = st.sidebar.radio("메뉴 선택", ["1. 데이터 업로드 및 관리", "2. 포인트 지급 대상 조회", "3. 상품권 지급 대상 조회"])
doc = authenticate_gspread()

# ==========================================
# 1. 데이터 업로드 및 관리 페이지
# ==========================================
if menu == "1. 데이터 업로드 및 관리":
    st.header("📂 데이터 업로드 및 관리")
    if doc is None: st.stop()

    # 1-1. 경리나라 수납 데이터 (A열 매칭을 위해 전체 열 업로드)
    st.subheader("1. 경리나라 수납 데이터 업로드 (첫 행 제외)")
    if st.button("🗑️ 기존 수납 데이터 일괄 삭제", key="c1"):
        if clear_google_sheet(doc, "경리나라 수납"):
            st.warning("경리나라 수납 데이터가 삭제되었습니다.")

    receipt_file = st.file_uploader("수납 파일 업로드 (xlsx, csv)", type=['xlsx', 'xls', 'csv'], key="u1")
    if receipt_file:
        df_receipt = load_file_generic(receipt_file, skip_rows=1)
        # 매칭용 A열 하이픈 제거
        if not df_receipt.empty:
            df_receipt.iloc[:, 0] = df_receipt.iloc[:, 0].astype(str).str.replace('-', '', regex=False)
            st.write("📊 업로드 데이터 미리보기 (A열 하이픈 제거됨):")
            st.dataframe(df_receipt.head(3))
            if st.button("경리나라 수납 시트 반영"):
                overwrite_google_sheet(doc, "경리나라 수납", df_receipt)
                st.success("반영 완료")

    st.divider()

    # 1-2. 추천 데이터 (3행 제외, C열->A열 복사, M열 반영)
    st.subheader("2. 추천 데이터 업로드 (3행 제외)")
    referral_file = st.file_uploader("추천 파일 업로드", type=['xlsx', 'xls', 'csv'], key="u2")
    if referral_file:
        df_ref_raw = load_file_generic(referral_file, skip_rows=3)
        if not df_ref_raw.empty:
            # C열(2번)을 A열(0번)로 복사 (하이픈 제거)
            df_ref_raw.iloc[:, 0] = df_ref_raw.iloc[:, 2].astype(str).str.replace('-', '', regex=False)
            # 필요한 열 추출: A, B, C, D, M(12), AT, AU, AV
            target_indices = [0, 1, 2, 3, 12, col2idx('AT'), col2idx('AU'), col2idx('AV')]
            df_ref_final = df_ref_raw.iloc[:, [i for i in target_indices if i < df_ref_raw.shape[1]]].copy()
            st.dataframe(df_ref_final.head(3))
            if st.button("추천 데이터 누적 추가"):
                append_to_google_sheet(doc, "추천", df_ref_final)
                st.success("누적 완료")

# ==========================================
# 2. 포인트 지급 대상 조회 및 보고서
# ==========================================
elif menu == "2. 포인트 지급 대상 조회":
    st.header("🎯 포인트 지급 내역 산출")
    if doc:
        with st.spinner("수납 시트 A열과 매칭하여 청구금액 계산 중..."):
            try:
                # 모든 시트 데이터 로드
                r_values = doc.worksheet("경리나라 수납").get_all_values()
                ref_values = doc.worksheet("추천").get_all_values()
                we_values = doc.worksheet("위멤버스 가입 여부").get_all_values()
                rate_values = doc.worksheet("적립율").get_all_values()
                
                if not r_values or not ref_values:
                    st.warning("수납 또는 추천 데이터가 부족합니다.")
                    st.stop()

                r_df = pd.DataFrame(r_values).fillna("0")
                ref_df = pd.DataFrame(ref_values).fillna("")
                
                # 적립율 딕셔너리 (기본값 0.03)
                rate_dict = {}
                for row in rate_values:
                    if len(row) >= 2:
                        biz = str(row[0]).replace('-', '').strip()
                        try:
                            v = str(row[1]).replace('%','').strip()
                            rate_dict[biz] = float(v)/100 if float(v) > 1 else float(v)
                        except: continue

                # 위멤버스 딕셔너리
                we_dict = {str(row[0]).replace('-', '').strip(): {'제품명': row[1], '비고': row[2]} 
                           for row in we_values if len(row) >= 3}

                # 산출 로직
                results = []
                for _, ref_row in ref_df.iterrows():
                    target_biz = str(ref_row[0]).replace('-', '').strip() # 추천 시트 A열
                    
                    # [매칭] 수납 시트의 A열(0번 인덱스)과 매칭
                    matched_r = r_df[r_df[0] == target_biz]
                    if matched_r.empty: continue
                    
                    for _, r_row in matched_r.iterrows():
                        rec_biz_raw = str(ref_row[6]) # 추천 시트 AU열
                        rec_biz_clean = rec_biz_raw.replace('-', '').strip()
                        
                        # 기한 필터링 (20251231 이후 제외)
                        rec_date = str(ref_row[4])
                        if ''.join(filter(str.isdigit, rec_date)) > "20251231": continue

                        we_info = we_dict.get(rec_biz_clean, {'제품명': '', '비고': ''})
                        if not str(we_info['제품명']).strip(): continue

                        # [청구금액] 수납 시트의 E열 (4번 인덱스)
                        raw_bill = str(r_row[4]).replace(',','').strip() if len(r_row) > 4 else "0"
                        bill_amt = pd.to_numeric(raw_bill, errors='coerce') or 0
                        
                        rate = rate_dict.get(rec_biz_clean, 0.03)
                        
                        results.append({
                            "수납사명": r_row[1], # 수납 시트 B열(상호)
                            "수납사업자번호": target_biz,
                            "추천자회사": ref_row[5], # 추천 AT열
                            "추천자사업자": rec_biz_raw, # 추천 AU열
                            "위멤버스_제품": we_info['제품명'],
                            "위멤버스_비고": we_info['비고'],
                            "적립율": rate,
                            "청구금액(E)": int(bill_amt),
                            "최종지급포인트": int(bill_amt * rate)
                        })

                final_df = pd.DataFrame(results)
                if not final_df.empty:
                    st.success(f"분석 완료: 총 {len(final_df)}건")
                    st.dataframe(final_df, use_container_width=True)
                    
                    st.divider()
                    st.subheader("📋 추천자별 합계 보고서")
                    summary = final_df.groupby(["추천자회사", "추천자사업자", "위멤버스_비고"])["최종지급포인트"].sum().reset_index()
                    st.dataframe(summary, use_container_width=True)
                    st.download_button("📥 결과 다운로드(CSV)", final_df.to_csv(index=False).encode('utf-8-sig'), "point_report.csv")
                else:
                    st.info("조건에 맞는 매칭 데이터가 없습니다.")
            except Exception as e:
                st.error(f"계산 중 오류 발생: {e}")

# ==========================================
# 3. 상품권 지급 대상 조회
# ==========================================
elif menu == "3. 상품권 지급 대상 조회":
    st.header("🎟️ 상품권 지급 대상 (1회차 수납)")
    if doc:
        with st.spinner("1회차 수납 데이터 조회 중..."):
            try:
                r_values = doc.worksheet("경리나라 수납").get_all_values()
                ref_values = doc.worksheet("추천").get_all_values()
                if not r_values: st.stop()
                
                r_df = pd.DataFrame(r_values)
                # AL열(37번 인덱스)이 입금횟수라고 가정하여 1회차 필터링
                r_df[37] = pd.to_numeric(r_df[37], errors='coerce')
                gift_targets = r_df[r_df[37] == 1].copy()
                
                # 추천 시트와 매칭하여 출력 (생략 가능하나 포인트 조회와 유사하게 구현)
                st.write("정상입금횟수가 1회인 대상자 목록입니다.")
                st.dataframe(gift_targets, use_container_width=True)
            except Exception as e:
                st.error(f"조회 중 오류: {e}")