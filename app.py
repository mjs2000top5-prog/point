import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

# ==========================================
# 0. 기본 설정 및 구글 API 연동 (배포용 Secrets 적용)
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
        st.info("💡 배포 시 Settings > Secrets에 GCP 키를 올바르게 입력했는지 확인해주세요.")
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

def load_file_without_header(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, header=None, dtype=str)
    else:
        df = pd.read_excel(file, header=None, dtype=str)
    if not df.empty and ('사업자' in str(df.iloc[0, 0]) or '청구' in str(df.iloc[0, 0])):
        df = df.iloc[1:].reset_index(drop=True)
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

    st.subheader("1. 경리나라 수납 데이터 (첫 행 제외 및 13개열 추출)")
    receipt_file = st.file_uploader("경리나라 수납 파일 업로드", type=['xlsx', 'xls', 'csv'], key="receipt")
    if receipt_file:
        df_receipt = load_file_without_header(receipt_file)
        df_receipt_final = df_receipt.iloc[:, :13].copy()
        for c in range(df_receipt_final.shape[1], 13): df_receipt_final[c] = ""
        st.dataframe(df_receipt_final.head(3))
        if st.button("경리나라 수납 데이터 반영"):
            overwrite_google_sheet(doc, "경리나라 수납", df_receipt_final)
            st.success("✅ 수납 데이터 반영 완료")

    st.divider()

    st.subheader("2. 추천 데이터 (누적 - 추천일 포함)")
    referral_file = st.file_uploader("추천 파일 업로드", type=['xlsx', 'xls', 'csv'], key="referral")
    if referral_file:
        df_raw = load_file_without_header(referral_file)
        if df_raw.shape[1] > 2:
            df_raw.iloc[:, 0] = df_raw.iloc[:, 2].astype(str).str.replace('-', '', regex=False)
        
        if df_raw.shape[1] <= 15:
            target_indices = [0, 1, 2, 3, 4, 5, 6, 7]
        else:
            target_indices = [0, 1, 2, 3, 4, col2idx('AT'), col2idx('AU'), col2idx('AV')]
            
        df_referral_final = df_raw.iloc[:, [i for i in target_indices if i < df_raw.shape[1]]].copy()
        for c in range(df_referral_final.shape[1], 8): df_referral_final[c] = ""
        st.dataframe(df_referral_final.head(3))
        if st.button("추천 데이터 누적 추가"):
            append_to_google_sheet(doc, "추천", df_referral_final)
            st.success("✅ 추천 데이터 누적 완료")

    st.divider()

    st.subheader("3. 위멤버스 가입 여부 데이터 (D, G, BQ열 추출)")
    wemembers_file = st.file_uploader("위멤버스 가입 여부 파일 업로드", type=['xlsx', 'xls', 'csv'], key="wemembers")
    if wemembers_file:
        df_we = load_file_without_header(wemembers_file)
        if df_we.shape[1] <= 5: target_indices = [0, 1, 2]
        else: target_indices = [col2idx('D'), col2idx('G'), col2idx('BQ')]
            
        df_we_final = df_we.iloc[:, [i for i in target_indices if i < df_we.shape[1]]].copy()
        if df_we_final.shape[1] > 0:
            df_we_final.iloc[:, 0] = df_we_final.iloc[:, 0].astype(str).str.replace('-', '', regex=False)
        for c in range(df_we_final.shape[1], 3): df_we_final[c] = ""
        st.dataframe(df_we_final.head(3))
        if st.button("위멤버스 데이터 시트에 반영"):
            overwrite_google_sheet(doc, "위멤버스 가입 여부", df_we_final)
            st.success("✅ 위멤버스 반영 완료")

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
            rate_data = doc.worksheet("적립율").get_all_values()
            
            if len(receipt_data) < 1 or len(referral_data) < 1:
                st.warning("데이터가 부족합니다. 먼저 업로드해주세요.")
                st.stop()
            
            if len(receipt_data) > 0 and '사업자' in str(receipt_data[0][0]): receipt_data = receipt_data[1:]
            if len(referral_data) > 0 and '사업자' in str(referral_data[0][0]): referral_data = referral_data[1:]
            if len(wemembers_data) > 0 and '사업자' in str(wemembers_data[0][0]): wemembers_data = wemembers_data[1:]

            r_df = pd.DataFrame(receipt_data)
            ref_df = pd.DataFrame(referral_data)
            
            r_df.columns = [f"r_{i}" for i in range(len(r_df.columns))]
            for i in range(13): 
                if f"r_{i}" not in r_df.columns: r_df[f"r_{i}"] = ""
            ref_df.columns = [f"ref_{i}" for i in range(len(ref_df.columns))]
            for i in range(8):
                if f"ref_{i}" not in ref_df.columns: ref_df[f"ref_{i}"] = ""

            rate_dict = {str(row[0]).replace('-', '').strip(): float(str(row[1]).replace('%','')) for row in rate_data if len(row) >= 2 and '사업자' not in str(row[0])}
            
            we_dict = {}
            for row in wemembers_data:
                if len(row) >= 1:
                    biz_key = str(row[0]).replace('-', '').strip()
                    if biz_key:
                        we_dict[biz_key] = {'제품명': row[1] if len(row) > 1 else "", '비고': row[2] if len(row) > 2 else ""}

        try:
            r_df["r_5"] = pd.to_numeric(r_df["r_5"], errors='coerce')
            filtered_r = r_df[(r_df["r_5"].isna()) | (r_df["r_5"] < 60)].copy()
            filtered_r = filtered_r[filtered_r["r_12"].astype(str).str.strip() != "미가입자"]

            merged_df = pd.merge(filtered_r, ref_df, left_on="r_0", right_on="ref_0")

            if merged_df.empty:
                st.info("조건에 일치하는 지급 대상이 없습니다.")
            else:
                results = []
                for _, row in merged_df.iterrows():
                    rec_biz_raw = str(row.get("ref_6", ""))
                    rec_biz_clean = rec_biz_raw.replace('-', '').strip()
                    
                    rec_date_raw = str(row.get("ref_4", ""))
                    clean_date = ''.join(filter(str.isdigit, rec_date_raw))
                    if clean_date and int(clean_date) > 20251231: continue
                    
                    we_info = we_dict.get(rec_biz_clean, {'제품명': '', '비고': ''})
                    we_prod = str(we_info.get('제품명', '')).strip()
                    if not we_prod or we_prod.lower() == 'nan': continue

                    base_amt = pd.to_numeric(row.get("r_3", 0), errors='coerce') or 0
                    rate = rate_dict.get(rec_biz_clean, 0.03)
                    
                    res_row = {}
                    for i in [0, 1, 2, 3, 5]: 
                        res_row[f"수납_{chr(65+i)}열"] = row.get(f"r_{i}", "")
                    
                    res_row["추천일"] = rec_date_raw
                    res_row["위멤버스_제품명(G)"] = we_prod
                    res_row["위멤버스_비고(BQ)"] = we_info['비고']
                    res_row["추천자 회사명"] = str(row.get("ref_5", ""))
                    res_row["추천자 사업자번호"] = rec_biz_raw
                    res_row["추천자"] = str(row.get("ref_7", ""))
                    res_row["적립율"] = rate
                    res_row["최종 지급포인트"] = base_amt * rate
                    results.append(res_row)

                final_df = pd.DataFrame(results)
                if final_df.empty:
                    st.info("조건을 충족하는 대상이 없습니다.")
                else:
                    st.success(f"총 {len(final_df)} 건이 확인되었습니다.")
                    st.dataframe(final_df, use_container_width=True)
                    st.metric("총 지급 예정 포인트 합계", f"{final_df['최종 지급포인트'].sum():,.0f} P")
                    
                    # 📊 요약 보고서
                    st.divider()
                    st.subheader("📋 포인트 지급 내역 합계 보고서")
                    summary_df = final_df.groupby(["위멤버스_비고(BQ)", "추천자 회사명", "추천자 사업자번호"])["최종 지급포인트"].sum().reset_index()
                    st.dataframe(summary_df, use_container_width=True)
                    
                    st.download_button("📥 상세 결과 CSV 다운로드", final_df.to_csv(index=False).encode('utf-8-sig'), "point_result.csv")
        except Exception as e:
            st.error(f"오류 발생: {e}")

# ==========================================
# 3. 상품권 지급 대상 조회 페이지
# ==========================================
elif menu == "3. 상품권 지급 대상 조회":
    st.header("🎟️ 상품권 지급 대상 조회")
    if doc:
        with st.spinner("분석 중..."):
            receipt_data = doc.worksheet("경리나라 수납").get_all_values()
            referral_data = doc.worksheet("추천").get_all_values()
            wemembers_data = doc.worksheet("위멤버스 가입 여부").get_all_values()
            
            if len(receipt_data) > 0 and '사업자' in str(receipt_data[0][0]): receipt_data = receipt_data[1:]
            if len(referral_data) > 0 and '사업자' in str(referral_data[0][0]): referral_data = referral_data[1:]
            if len(wemembers_data) > 0 and '사업자' in str(wemembers_data[0][0]): wemembers_data = wemembers_data[1:]

            r_df = pd.DataFrame(receipt_data)
            ref_df = pd.DataFrame(referral_data)
            
            r_df.columns = [f"r_{i}" for i in range(len(r_df.columns))]
            for i in range(13): 
                if f"r_{i}" not in r_df.columns: r_df[f"r_{i}"] = ""
            ref_df.columns = [f"ref_{i}" for i in range(len(ref_df.columns))]
            for i in range(8):
                if f"ref_{i}" not in ref_df.columns: ref_df[f"ref_{i}"] = ""
            
            we_dict = {}
            for row in wemembers_data:
                if len(row) >= 1:
                    biz_key = str(row[0]).replace('-', '').strip()
                    if biz_key:
                        we_dict[biz_key] = {'제품명': row[1] if len(row) > 1 else "", '비고': row[2] if len(row) > 2 else ""}

        try:
            r_df["r_5"] = pd.to_numeric(r_df["r_5"], errors='coerce')
            filtered_r = r_df[r_df["r_5"] == 1].copy()

            merged_df = pd.merge(filtered_r, ref_df, left_on="r_0", right_on="ref_0")

            if merged_df.empty:
                st.info("지급 대상(1회차 수납)이 없습니다.")
            else:
                results = []
                for _, row in merged_df.iterrows():
                    rec_biz_raw = str(row.get("ref_6", ""))
                    rec_biz_clean = rec_biz_raw.replace('-', '').strip()
                    we_info = we_dict.get(rec_biz_clean, {'제품명': '', '비고': ''})
                    
                    res_row = {
                        "수납_A열": row.get("r_0", ""),
                        "수납_B열": row.get("r_1", ""),
                        "수납_C열": row.get("r_2", ""),
                        "수납_F열": row.get("r_5", ""),
                        "추천일": str(row.get("ref_4", "")),
                        "위멤버스_제품명(G)": we_info['제품명'],
                        "위멤버스_비고(BQ)": we_info['비고'],
                        "추천자사업자번호": rec_biz_raw,
                        "추천자": str(row.get("ref_7", ""))
                    }
                    results.append(res_row)

                final_df = pd.DataFrame(results)
                st.dataframe(final_df, use_container_width=True)
                st.download_button("📥 상품권 결과 다운로드", final_df.to_csv(index=False).encode('utf-8-sig'), "giftcard_result.csv")
        except Exception as e:
            st.error(f"오류 발생: {e}")