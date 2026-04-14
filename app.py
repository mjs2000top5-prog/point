import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ==========================================
# 0. 기본 설정 및 구글 API 연동 (Secrets)
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

    # 1-1. 경리나라 수납 데이터
    st.subheader("1. 경리나라 수납 (첫 행 제외 & G, I, E, W, X, AA, AL, AM 추출)")
    if st.button("🗑️ 경리나라 수납 데이터 일괄 삭제", key="c1"):
        if clear_google_sheet(doc, "경리나라 수납"):
            st.warning("경리나라 수납 데이터가 삭제되었습니다.")

    receipt_file = st.file_uploader("수납 파일 업로드", type=['xlsx', 'xls', 'csv'], key="u1")
    if receipt_file:
        df_raw = load_file_generic(receipt_file, skip_rows=1)
        # G(6), I(8), E(4), W(22), X(23), AA(26), AL(37), AM(38)
        t_cols = ['G', 'I', 'E', 'W', 'X', 'AA', 'AL', 'AM']
        t_idxs = [col2idx(c) for c in t_cols]
        df_final = df_raw.iloc[:, [i for i in t_idxs if i < df_raw.shape[1]]].copy()
        
        # G열 하이픈 제거
        if df_final.shape[1] > 0:
            df_final.iloc[:, 0] = df_final.iloc[:, 0].astype(str).str.replace('-', '', regex=False)
        
        st.write("미리보기 (3번째 열이 청구금액 E열입니다):")
        st.dataframe(df_final.head(3))
        if st.button("경리나라 수납 반영"):
            overwrite_google_sheet(doc, "경리나라 수납", df_final)
            st.success("반영 완료")

    st.divider()

    # 1-2. 추천 데이터
    st.subheader("2. 추천 데이터 (3행 제외, C->A 복사)")
    referral_file = st.file_uploader("추천 파일 업로드", type=['xlsx', 'xls', 'csv'], key="u2")
    if referral_file:
        df_raw = load_file_generic(referral_file, skip_rows=3)
        if not df_raw.empty:
            # C열을 A열로 하이픈 제거 후 복사
            if df_raw.shape[1] > 2:
                df_raw.iloc[:, 0] = df_raw.iloc[:, 2].astype(str).str.replace('-', '', regex=False)
            
            t_idxs = [0, 1, 2, 3, 12, col2idx('AT'), col2idx('AU'), col2idx('AV')]
            df_final = df_raw.iloc[:, [i for i in t_idxs if i < df_raw.shape[1]]].copy()
            for c in range(df_final.shape[1], 8): df_final[c] = ""
            st.dataframe(df_final.head(3))
            if st.button("추천 데이터 누적 추가"):
                append_to_google_sheet(doc, "추천", df_final)
                st.success("누적 완료")

    st.divider()

    # 1-3. 위멤버스 가입 여부
    st.subheader("3. 위멤버스 가입 여부 (D, G, BQ 추출)")
    if st.button("🗑️ 위멤버스 데이터 일괄 삭제", key="c3"):
        if clear_google_sheet(doc, "위멤버스 가입 여부"):
            st.warning("위멤버스 데이터가 삭제되었습니다.")

    we_file = st.file_uploader("위멤버스 파일 업로드", type=['xlsx', 'xls', 'csv'], key="u3")
    if we_file:
        df_raw = load_file_generic(we_file, skip_rows=0)
        if not df_raw.empty and '사업자' in str(df_raw.iloc[0, 0]):
            df_raw = df_raw.iloc[1:].reset_index(drop=True)
            
        t_idxs = [col2idx('D'), col2idx('G'), col2idx('BQ')]
        df_final = df_raw.iloc[:, [i for i in t_idxs if i < df_raw.shape[1]]].copy()
        df_final.iloc[:, 0] = df_final.iloc[:, 0].astype(str).str.replace('-', '', regex=False)
        for c in range(df_final.shape[1], 3): df_final[c] = ""
        st.dataframe(df_final.head(3))
        if st.button("위멤버스 반영"):
            overwrite_google_sheet(doc, "위멤버스 가입 여부", df_final)
            st.success("반영 완료")

# ==========================================
# 2. 포인트 지급 대상 조회 및 보고서
# ==========================================
elif menu == "2. 포인트 지급 대상 조회":
    st.header("🎯 포인트 지급 내역 산출")
    if doc:
        with st.spinner("데이터 분석 및 계산 중..."):
            try:
                r_data = doc.worksheet("경리나라 수납").get_all_values()
                ref_data = doc.worksheet("추천").get_all_values()
                we_data = doc.worksheet("위멤버스 가입 여부").get_all_values()
                rate_data = doc.worksheet("적립율").get_all_values()
                
                if not r_data or not ref_data:
                    st.info("수납 또는 추천 데이터가 구글 시트에 없습니다.")
                    st.stop()

                # 데이터프레임 변환 및 빈값 처리
                r_df = pd.DataFrame(r_data).fillna("")
                ref_df = pd.DataFrame(ref_data).fillna("")
                
                # 적립율 매핑
                rate_dict = {}
                for row in rate_data:
                    if len(row) >= 2:
                        biz = str(row[0]).replace('-', '').strip()
                        try:
                            v = str(row[1]).replace('%','').strip()
                            rate_dict[biz] = float(v)/100 if float(v) > 1 else float(v)
                        except: continue

                # 위멤버스 매핑
                we_dict = {str(row[0]).replace('-', '').strip(): {'제품명': row[1], '비고': row[2]} for row in we_data if len(row) >= 3}

                r_df.columns = [f"r_{i}" for i in range(len(r_df.columns))]
                ref_df.columns = [f"ref_{i}" for i in range(len(ref_df.columns))]
                
                # 수납 사업자(r_0) - 추천 대상 사업자(ref_0) 매칭
                merged = pd.merge(r_df, ref_df, left_on="r_0", right_on="ref_0")

                results = []
                for _, row in merged.iterrows():
                    rec_biz_raw = str(row.get("ref_6", ""))
                    rec_biz_clean = rec_biz_raw.replace('-', '').strip()
                    
                    # 기한 필터링 (20251231)
                    rec_date = str(row.get("ref_4", ""))
                    if ''.join(filter(str.isdigit, rec_date)) > "20251231": continue

                    we_info = we_dict.get(rec_biz_clean, {'제품명': '', '비고': ''})
                    if not str(we_info['제품명']).strip(): continue

                    # [오류 방지 로직] 청구금액(r_2) 처리
                    raw_bill = str(row.get("r_2", "0")).replace(',','').strip()
                    bill_amt = pd.to_numeric(raw_bill, errors='coerce')
                    
                    # 적립율 (기본 3%)
                    rate = rate_dict.get(rec_biz_clean, 0.03)
                    
                    # 계산 결과 안정화
                    final_point = 0
                    if pd.notna(bill_amt):
                        final_point = bill_amt * rate
                    
                    results.append({
                        "수납 회사명": row.get("r_1"),
                        "수납_사업자": row.get("r_0"),
                        "추천자 회사명": row.get("ref_5"),
                        "추천자 사업자번호": rec_biz_raw,
                        "위멤버스_제품명(G)": we_info['제품명'],
                        "위멤버스_비고(BQ)": we_info['비고'],
                        "적립율": rate,
                        "청구금액": bill_amt if pd.notna(bill_amt) else 0,
                        "최종 지급 포인트": int(final_point) if pd.notna(final_point) else 0
                    })

                final_df = pd.DataFrame(results)
                if not final_df.empty:
                    st.success(f"분석 완료: 총 {len(final_df)}건")
                    st.dataframe(final_df, use_container_width=True)
                    
                    st.divider()
                    st.subheader("📋 추천자별 합계 보고서")
                    summary = final_df.groupby(["추천자 회사명", "추천자 사업자번호", "위멤버스_비고(BQ)"])["최종 지급 포인트"].sum().reset_index()
                    st.dataframe(summary, use_container_width=True)
                    st.download_button("📥 상세 결과 다운로드", final_df.to_csv(index=False).encode('utf-8-sig'), "point_calc_result.csv")
                else:
                    st.info("조건을 만족하는 지급 대상 데이터가 없습니다.")
            except Exception as e:
                st.error(f"데이터 처리 중 오류 발생: {e}")

# ==========================================
# 3. 상품권 지급 대상 조회
# ==========================================
elif menu == "3. 상품권 지급 대상 조회":
    st.header("🎟️ 상품권 지급 대상 (1회차 수납)")
    if doc:
        with st.spinner("대상 조회 중..."):
            try:
                r_data = doc.worksheet("경리나라 수납").get_all_values()
                ref_data = doc.worksheet("추천").get_all_values()
                we_data = doc.worksheet("위멤버스 가입 여부").get_all_values()
                
                if not r_data: st.stop()
                
                r_df = pd.DataFrame(r_data).fillna("")
                ref_df = pd.DataFrame(ref_data).fillna("")
                r_df.columns = [f"r_{i}" for i in range(len(r_df.columns))]
                ref_df.columns = [f"ref_{i}" for i in range(len(ref_df.columns))]
                
                we_dict = {str(row[0]).replace('-', '').strip(): {'제품명': row[1], '비고': row[2]} for row in we_data if len(row) >= 3}

                # r_5(회차)가 1인 데이터 필터링
                r_df["r_5"] = pd.to_numeric(r_df["r_5"], errors='coerce')
                filtered_r = r_df[r_df["r_5"] == 1].copy()
                merged_g = pd.merge(filtered_r, ref_df, left_on="r_0", right_on="ref_0")

                if not merged_g.empty:
                    g_results = []
                    for _, row in merged_g.iterrows():
                        rec_biz = str(row.get("ref_6", "")).replace('-', '').strip()
                        we_info = we_dict.get(rec_biz, {'제품명': '', '비고': ''})
                        g_results.append({
                            "수납사명": row.get("r_1"),
                            "사업자번호": row.get("r_0"),
                            "회차": row.get("r_5"),
                            "추천자": row.get("ref_7"),
                            "위멤버스_비고": we_info['비고']
                        })
                    st.dataframe(pd.DataFrame(g_results), use_container_width=True)
                else:
                    st.info("1회차 수납 대상자가 없습니다.")
            except Exception as e:
                st.error(f"조회 중 오류 발생: {e}")