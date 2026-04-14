import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ==========================================
# 0. 기본 설정 및 구글 API 연동 (Secrets 적용)
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

# 시트 관리 함수들
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

    st.subheader("1. 경리나라 수납 데이터 업로드 (첫 행 제외)")
    if st.button("🗑️ 기존 수납 데이터 일괄 삭제", key="clear_r"):
        clear_google_sheet(doc, "경리나라 수납")

    receipt_file = st.file_uploader("수납 파일 업로드", type=['xlsx', 'xls', 'csv'], key="u1")
    if receipt_file:
        df_receipt = load_file_generic(receipt_file, skip_rows=1)
        if not df_receipt.empty:
            df_receipt.iloc[:, 0] = df_receipt.iloc[:, 0].astype(str).str.replace('-', '', regex=False)
            st.dataframe(df_receipt.head(3))
            if st.button("경리나라 수납 시트 반영"):
                overwrite_google_sheet(doc, "경리나라 수납", df_receipt)
                st.success("반영 완료")

    st.divider()

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
                st.success("누적 완료")

# ==========================================
# 데이터 처리 공통 로직 함수 (포인트/상품권 공용)
# ==========================================
def get_processed_data(doc, filter_count_one=False):
    r_values = doc.worksheet("경리나라 수납").get_all_values()
    ref_values = doc.worksheet("추천").get_all_values()
    we_values = doc.worksheet("위멤버스 가입 여부").get_all_values()
    rate_values = doc.worksheet("적립율").get_all_values()
    
    if not r_values or not ref_values:
        return pd.DataFrame()

    r_df = pd.DataFrame(r_values).fillna("")
    ref_df = pd.DataFrame(ref_values).fillna("")
    
    # 사전 구축
    rate_dict = {}
    for row in rate_values:
        if len(row) >= 2:
            biz = str(row[0]).replace('-', '').strip()
            try:
                val = float(str(row[1]).replace('%','').strip())
                rate_dict[biz] = val/100 if val > 1 else val
            except: continue
    we_dict = {str(row[0]).replace('-', '').strip(): {'제품명': row[1], '비고': row[2]} for row in we_values if len(row) >= 3}

    results = []
    for _, ref_row in ref_df.iterrows():
        target_biz = str(ref_row[0]).replace('-', '').strip()
        matched_rows = r_df[r_df[0] == target_biz]
        
        for _, r_row in matched_rows.iterrows():
            # [필터 1] 수납횟수(G열)
            raw_count = str(r_row[6]).strip() if len(r_row) > 6 else "0"
            pay_count = pd.to_numeric(raw_count, errors='coerce')
            if pd.isna(pay_count): continue
            
            if filter_count_one:
                if pay_count != 1: continue # 상품권용: 무조건 1회만
            else:
                if pay_count <= 0 or pay_count >= 60: continue # 포인트용: 1~59회

            # [필터 2] 청구금액(E열) 4만 원 미만 제외
            raw_bill = str(r_row[4]).replace(',','').strip() if len(r_row) > 4 else "0"
            bill_amt = pd.to_numeric(raw_bill, errors='coerce')
            if pd.isna(bill_amt) or bill_amt < 40000: continue

            # [필터 3] 설치일(C열) 2026-01-01 이후 제외
            install_date_raw = str(r_row[2]) if len(r_row) > 2 else ""
            install_clean = ''.join(filter(str.isdigit, install_date_raw))
            if install_clean and len(install_clean) >= 8 and int(install_clean[:8]) >= 20260101: continue
            
            # [필터 4] 추천 기한 2025-12-31 이전
            rec_date = str(ref_row[4])
            if ''.join(filter(str.isdigit, rec_date)) > "20251231": continue

            # [필터 5] 위멤버스 제품 (베이직 제외)
            rec_biz_raw = str(ref_row[6])
            rec_biz_clean = rec_biz_raw.replace('-', '').strip()
            we_info = we_dict.get(rec_biz_clean, {'제품명': '', '비고': ''})
            we_product_name = str(we_info['제품명']).strip()
            if not we_product_name or "위멤버스 베이직" in we_product_name: continue

            rate = rate_dict.get(rec_biz_clean, 0.03)
            final_p = bill_amt * rate
            
            results.append({
                "설치일(C)": install_date_raw,
                "수납횟수(G)": int(pay_count),
                "수납사명": r_row[1],
                "수납사업자번호": target_biz,
                "추천자회사": ref_row[5],
                "추천자사업자": rec_biz_raw,
                "위멤버스_제품": we_product_name,
                "위멤버스_비고": we_info['비고'],
                "적립율": rate,
                "청구금액(E)": int(bill_amt),
                "최종지급포인트": int(final_p)
            })
    return pd.DataFrame(results)

# ==========================================
# 2. 포인트 지급 대상 조회
# ==========================================
if menu == "2. 포인트 지급 대상 조회":
    st.header("🎯 포인트 지급 내역 산출 (1~59회)")
    if doc:
        with st.spinner("데이터 분석 중..."):
            final_df = get_processed_data(doc, filter_count_one=False)
            if not final_df.empty:
                st.subheader("1. 상세 내역")
                st.dataframe(final_df, use_container_width=True)
                st.download_button("📥 상세 내역 다운로드", final_df.to_csv(index=False).encode('utf-8-sig'), "point_detail.csv")
                
                st.divider()
                st.subheader("2. 추천자별 합계")
                summary = final_df.groupby(["추천자회사", "추천자사업자", "위멤버스_비고"])["최종지급포인트"].sum().reset_index()
                st.dataframe(summary, use_container_width=True)
                st.download_button("📥 합계 보고서 다운로드", summary.to_csv(index=False).encode('utf-8-sig'), "point_summary.csv")
            else:
                st.info("조건을 충족하는 데이터가 없습니다.")

# ==========================================
# 3. 상품권 지급 대상 조회
# ==========================================
elif menu == "3. 상품권 지급 대상 조회":
    st.header("🎟️ 상품권 지급 대상 (수납횟수 1회만)")
    if doc:
        with st.spinner("수납 1회 대상자 필터링 중..."):
            # filter_count_one=True를 전달하여 수납횟수가 1인 것만 가져옴
            gift_df = get_processed_data(doc, filter_count_one=True)
            if not gift_df.empty:
                st.success(f"상품권 지급 대상: 총 {len(gift_df)}건")
                st.dataframe(gift_df, use_container_width=True)
                st.download_button("📥 상품권 지급 대상 다운로드", gift_df.to_csv(index=False).encode('utf-8-sig'), "gift_card_targets.csv")
            else:
                st.info("상품권 지급 대상(수납 1회 및 기타 조건 충족)이 없습니다.")