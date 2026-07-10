import streamlit as st
import pandas as pd
import pdfplumber
import re
from collections import defaultdict

# 웹페이지 기본 설정
st.set_page_config(page_title="보증금액 통합 비교 시스템", layout="wide")
st.title("📊 보증금액 교차 비교 시스템")

# ────────────────────────────────────────────────────────
# 🗂️ 사이드바 / 상단 메뉴를 통한 작업 선택 UI
# ────────────────────────────────────────────────────────
st.sidebar.header("⚙️ 작업 모드 선택")
mode = st.sidebar.radio(
    "실행할 분석 작업을 선택하세요:",
    ("📋 MW 보증 비교 (PDF vs 엑셀)", "🚗 쿠폰 보증 비교 (엑셀 vs 엑셀)")
)

# ────────────────────────────────────────────────────────
# 🛠️ [공통 함수] 엑셀 열 인덱스 또는 이름으로 안전하게 데이터 가져오기
# ────────────────────────────────────────────────────────
def get_col_by_idx_or_name(df, col_idx, possible_names):
    if col_idx < len(df.columns):
        return df.columns[col_idx]
    for name in possible_names:
        for col in df.columns:
            if name.upper() in str(col).upper().strip():
                return col
    return None

# ★ 정확한 반올림을 위한 헬퍼 함수 (0.5 이상 올림)
def round_half_up(value):
    return int(value + 0.5)

# ────────────────────────────────────────────────────────
# 1️⃣ [모드 1] MW 보증 비교 관련 로직
# ────────────────────────────────────────────────────────
def load_excel_mw(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.upper()
    col_claim_no = 'CLAIM NO'
    
    target_cols = ['공임청구액', '공임청구부가세', '부품청구액', '부품청구부가세']
    for col in target_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
    df['Excel_Total'] = (df.get('공임청구액', 0) + df.get('공임청구부가세', 0) + 
                         df.get('부품청구액', 0) + df.get('부품청구부가세', 0)).apply(round_half_up)
    
    excel_groups = defaultdict(list)
    for _, row in df.iterrows():
        claim_no = str(row.get(col_claim_no, '')).strip()
        if claim_no and claim_no != 'nan':
            excel_groups[claim_no].append(int(row['Excel_Total']))
    return excel_groups

def load_pdf_mw(uploaded_file):
    pdf_groups = defaultdict(list)
    with pdfplumber.open(uploaded_file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            if page_num % 2 != 0:
                continue
            text = page.extract_text()
            if not text:
                continue
            lines = text.split('\n')
            page_seen = set()
            for line in lines:
                line_stripped = line.strip()
                if line_stripped in page_seen:
                    continue
                page_seen.add(line_stripped)
                
                # ★ 중요 변경: 알파벳 개수([A-Z]+), 숫자 개수(\d+) 제한을 완전히 풀어 
                # 영어가 4개든 숫자가 5개든 어떤 조합이든 완벽하게 매칭합니다.
                match = re.search(r'([A-Z]+\d+)', line_stripped)
                if match:
                    rep_order = match.group(1)
                    parts = line_stripped.split()
                    try:
                        total_str = parts[-1].replace(',', '.')
                        pdf_total_with_vat = round_half_up(float(total_str) * 1.1)
                        pdf_groups[rep_order].append(pdf_total_with_vat)
                    except ValueError:
                        continue
    return pdf_groups

# ────────────────────────────────────────────────────────
# 2️⃣ [모드 2] 쿠폰 보증 비교 관련 로직
# ────────────────────────────────────────────────────────
def load_excel_coupon_a(uploaded_file):
    df = pd.read_excel(uploaded_file)
    
    col_car = get_col_by_idx_or_name(df, 3, ['차량번호', 'CAR', 'VEHICLE'])
    col_part = get_col_by_idx_or_name(df, 8, ['부품청구', '부품'])
    col_labour = get_col_by_idx_or_name(df, 9, ['공임청구', '공임'])
    
    if col_part: df[col_part] = pd.to_numeric(df[col_part], errors='coerce').fillna(0)
    if col_labour: df[col_labour] = pd.to_numeric(df[col_labour], errors='coerce').fillna(0)
    
    df['Calc_Total'] = ((df[col_part] + df[col_labour]) * 1.1).apply(round_half_up)
    
    a_groups = defaultdict(list)
    for _, row in df.iterrows():
        car_no = str(row[col_car]).strip() if col_car else 'Unknown'
        if car_no and car_no != 'nan':
            a_groups[car_no].append(int(row['Calc_Total']))
    return a_groups

def load_excel_coupon_b(uploaded_file):
    df = pd.read_excel(uploaded_file)
    
    col_car = get_col_by_idx_or_name(df, 6, ['차량번호', 'CAR', 'VEHICLE'])
    col_total = get_col_by_idx_or_name(df, 18, ['합계금액', '합계', 'TOTAL'])
    
    if col_total: df[col_total] = pd.to_numeric(df[col_total], errors='coerce').fillna(0)
    
    b_groups = defaultdict(list)
    for _, row in df.iterrows():
        car_no = str(row[col_car]).strip() if col_car else 'Unknown'
        if car_no and car_no != 'nan':
            b_groups[car_no].append(round_half_up(row[col_total]) if col_total else 0)
    return b_groups


# ────────────────────────────────────────────────────────
# 🖥️ 화면 조건별 렌더링
# ────────────────────────────────────────────────────────
if "MW 보증 비교" in mode:
    st.subheader("📋 MW 보증 비교 (PDF vs 엑셀)")
    st.write("PDF(홀수페이지)와 엑셀의 금액을 각각 계산 후 반올림 처리하여 순차 정렬 대조합니다. (글자 수 무제한 인식)")
    
    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("1. PDF 파일을 선택하세요 (.pdf)", type=["pdf"], key="mw_pdf")
    with col2:
        excel_file = st.file_uploader("2. 엑셀 파일을 선택하세요 (.xlsx)", type=["xlsx"], key="mw_excel")
        
    if pdf_file and excel_file:
        with st.spinner("MW 보증 데이터 교차 대조 중..."):
            excel_groups = load_excel_mw(excel_file)
            pdf_groups = load_pdf_mw(pdf_file)
            
            matched_results = []
            all_orders = sorted(list(set(list(excel_groups.keys()) + list(pdf_groups.keys()))))
            
            for order in all_orders:
                p_amts = pdf_groups.get(order, [])
                e_amts = excel_groups.get(order, [])
                max_len = max(len(p_amts), len(e_amts))
                
                for i in range(max_len):
                    p_amt = p_amts[i] if i < len(p_amts) else None
                    e_amt = e_amts[i] if i < len(e_amts) else None
                    order_label = f"{order} ({i+1})" if max_len > 1 else order
                    
                    if p_amt is not None and e_amt is not None:
                        diff = p_amt - e_amt
                        matched_results.append({
                            '주문번호': order_label, 'PDF 금액 (1.1배 반올림)': f"{p_amt:,}원", '엑셀 금액 (반올림)': f"{e_amt:,}원",
                            '차액': diff, '비고': "정확히 일치" if diff == 0 else f"불일치 ({diff:+,}원)"
                        })
                    elif p_amt is not None:
                        matched_results.append({
                            '주문번호': order_label, 'PDF 금액 (1.1배 반올림)': f"{p_amt:,}원", '엑셀 금액 (반올림)': "-",
                            '차액': p_amt, '비고': "★ 엑셀에 일치하는 항목 없음"
                        })
                    elif e_amt is not None:
                        matched_results.append({
                            '주문번호': order_label, 'PDF 금액 (1.1배 반올림)': "-", '엑셀 금액 (반올림)': f"{e_amt:,}원",
                            '차액': -e_amt, '비고': "★ PDF에 일치하는 항목 없음"
                        })
            
            res_df = pd.DataFrame(matched_results)
            total_diff = sum([item['차액'] for item in matched_results])
            
            st.subheader("📌 분석 요약 결과")
            m_col1, m_col2 = st.columns(2)
            m_col1.metric("총 대조 건수", f"{len(res_df)} 건")
            m_col2.metric("최종 총 차이 금액", f"{total_diff:,}원", delta=f"{total_diff:,}원" if total_diff != 0 else None)
            
            st.subheader("📋 상세 대조 내역")
            if not res_df.empty: res_df['차액'] = res_df['차액'].map(lambda x: f"{x:,}원" if isinstance(x, (int, float)) else x)
            st.dataframe(res_df, use_container_width=True)

else:
    st.subheader("🚗 쿠폰 보증 비교 (엑셀 vs 엑셀)")
    st.write("A파일(D열 차량번호, (I열+J열)*1.1 반올림)과 B파일(G열 차량번호, S열 합계금액 반올림)을 정밀 매칭합니다.")
    
    col1, col2 = st.columns(2)
    with col1:
        file_a = st.file_uploader("1. A 엑셀 파일을 선택하세요 (D, I, J행 포함)", type=["xlsx"], key="cp_a")
    with col2:
        file_b = st.file_uploader("2. B 엑셀 파일을 선택하세요 (G, S행 포함)", type=["xlsx"], key="cp_b")
        
    if file_a and file_b:
        with st.spinner("쿠폰 보증 엑셀 간 교차 대조 중..."):
            a_groups = load_excel_coupon_a(file_a)
            b_groups = load_excel_coupon_b(file_b)
            
            matched_results = []
            all_cars = sorted(list(set(list(a_groups.keys()) + list(b_groups.keys()))))
            
            for car in all_cars:
                a_amts = a_groups.get(car, [])
                b_amts = b_groups.get(car, [])
                max_len = max(len(a_amts), len(b_amts))
                
                for i in range(max_len):
                    a_amt = a_amts[i] if i < len(a_amts) else None
                    b_amt = b_amts[i] if i < len(b_amts) else None
                    car_label = f"{car} ({i+1})" if max_len > 1 else car
                    
                    if a_amt is not None and b_amt is not None:
                        diff = a_amt - b_amt
                        matched_results.append({
                            '차량번호': car_label, 'A파일 계산금액 (1.1배 반올림)': f"{a_amt:,}원", 'B파일 합계금액 (반올림)': f"{b_amt:,}원",
                            '차액': diff, '비고': "정확히 일치" if diff == 0 else f"불일치 ({diff:+,}원)"
                        })
                    elif a_amt is not None:
                        matched_results.append({
                            '차량번호': car_label, 'A파일 계산금액 (1.1배 반올림)': f"{a_amt:,}원", 'B파일 합계금액 (반올림)': "-",
                            '차액': a_amt, '비고': "★ B파일에 일치하는 항목 없음"
                        })
                    elif b_amt is not None:
                        matched_results.append({
                            '차량번호': car_label, 'A파일 계산금액 (1.1배 반올림)': "-", 'B파일 합계금액 (반올림)': f"{b_amt:,}원",
                            '차액': -b_amt, '비고': "★ A파일에 일치하는 항목 없음"
                        })
            
            res_df = pd.DataFrame(matched_results)
            total_diff = sum([item['차액'] for item in matched_results])
            
            st.subheader("📌 분석 요약 결과")
            m_col1, m_col2 = st.columns(2)
            m_col1.metric("총 대조 건수", f"{len(res_df)} 건")
            m_col2.metric("최종 총 차이 금액", f"{total_diff:,}원", delta=f"{total_diff:,}원" if total_diff != 0 else None)
            
            st.subheader("📋 상세 대조 내역 (차량번호 기준)")
            if not res_df.empty: res_df['차액'] = res_df['차액'].map(lambda x: f"{x:,}원" if isinstance(x, (int, float)) else x)
            st.dataframe(res_df, use_container_width=True)
