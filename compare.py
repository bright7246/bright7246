import streamlit as st
import pandas as pd
import pdfplumber
import re
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# 웹페이지 기본 설정
st.set_page_config(page_title="보증금액 통합 비교 시스템", layout="wide")
st.title("📊 보증금액 교차 비교 시스템")

# ────────────────────────────────────────────────────────
# 🗂️ 사이드바 / 상단 메뉴를 통한 작업 선택 UI
# ────────────────────────────────────────────────────────
st.sidebar.header("⚙️ 작업 모드 선택")
mode = st.sidebar.radio(
    "실행할 분석 작업을 선택하세요:",
    (
        "📋 MW 보증 비교 (PDF vs 엑셀)",
        "🚗 쿠폰 보증 비교 (엑셀 vs 엑셀)",
        "🔧 공임코드 비교"
    )
)

# ────────────────────────────────────────────────────────
# 🛠️ [공통 함수] 엑셀 상단 타이틀/빈줄을 건너뛰고 진짜 헤더 행 찾아 읽기
# ────────────────────────────────────────────────────────
def read_excel_smart_header(uploaded_file):
    uploaded_file.seek(0)
    df_raw = pd.read_excel(uploaded_file, header=None)
    
    header_row_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str)).upper()
        if 'CLAIM' in row_str or '차량' in row_str or '공임' in row_str:
            header_row_idx = idx
            break
            
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, header=header_row_idx)
    return df

def get_col_by_idx_or_name(df, col_idx, possible_names):
    if col_idx < len(df.columns):
        return df.columns[col_idx]
    for name in possible_names:
        for col in df.columns:
            if name.upper() in str(col).upper().strip():
                return col
    return None

def round_half_up(value):
    return int(value + 0.5)

# ────────────────────────────────────────────────────────
# 1️⃣ [모드 1] MW 보증 비교 관련 로직
# ────────────────────────────────────────────────────────
def load_excel_mw(uploaded_file):
    df = read_excel_smart_header(uploaded_file)
    df.columns = df.columns.astype(str).str.strip().str.upper()
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

def create_mw_excel_report(uploaded_file_mw, count, total_pdf, total_excel, total_diff):
    df_mw_raw = read_excel_smart_header(uploaded_file_mw)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "WARRANTY 수령내역"
    ws.print_title_rows = '1:3'
    
    target_headers = [
        "Claim No", "차량번호", "Job No", "완결일자", "청구일자",
        "공임청구액", "공임청구부가세", "부품청구액", "부품청구부가세",
        "공임입금액", "공임입금부가세", "부품입금액", "부품입금부가세"
    ]
    
    alias_dict = {
        "Claim No": ["CLAIM NO", "CLAIM", "클레임", "청구번호"],
        "차량번호": ["차량번호", "차량 번호", "CAR NO", "VEHICLE"],
        "Job No": ["JOB NO", "JOB", "작업번호"],
        "완결일자": ["완결일자", "완결일", "완결"],
        "청구일자": ["청구일자", "청구일"],
        "공임청구액": ["공임청구액", "공임청구", "공임 청구액"],
        "공임청구부가세": ["공임청구부가세", "공임청구 부가세", "공임부가세"],
        "부품청구액": ["부품청구액", "부품청구", "부품 청구액"],
        "부품청구부가세": ["부품청구부가세", "부품청구 부가세", "부품부가세"],
        "공임입금액": ["공임입금액", "공임입금", "공임승인액", "공임승인", "공임 입금액", "공임승인금액"],
        "공임입금부가세": ["공임입금부가세", "공임입금 부가세", "공임승인부가세"],
        "부품입금액": ["부품입금액", "부품입금", "부품승인액", "부품승인", "부품 입금액", "부품승인금액"],
        "부품입금부가세": ["부품입금부가세", "부품입금 부가세", "부품승인부가세"]
    }
    
    col_mapping = {}
    for th in target_headers:
        found_col = None
        possible_keywords = alias_dict.get(th, [th])
        for col in df_mw_raw.columns:
            clean_col = str(col).replace(" ", "").upper()
            for kw in possible_keywords:
                clean_kw = kw.replace(" ", "").upper()
                if clean_kw in clean_col:
                    found_col = col
                    break
            if found_col:
                break
        col_mapping[th] = found_col

    month_str = "6월"
    file_name = getattr(uploaded_file_mw, 'name', '')
    m_fn = re.search(r'20\d{2}(\d{2})', file_name)
    if m_fn:
        month_str = f"{int(m_fn.group(1))}월"
    else:
        for col in df_mw_raw.columns:
            if any(keyword in str(col).upper() for keyword in ['일자', 'DATE', '완결', '청구']):
                sample_dates = df_mw_raw[col].dropna().astype(str).tolist()
                for d in sample_dates:
                    m = re.search(r'-(\d{2})-', d) or re.search(r'/(\d{2})/', d)
                    if m:
                        month_str = f"{int(m.group(1))}월"
                        break
                if month_str != "6월":
                    break

    ws.merge_cells('A1:N1')
    ws['A1'] = f"{month_str} WARRANTY 수 령 내 역"
    ws['A1'].font = Font(size=22, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    header_font = Font(size=10, bold=True)
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws.row_dimensions[3].height = 25
    cell_a = ws.cell(row=3, column=1, value="No.")
    cell_a.font = header_font
    cell_a.alignment = header_align
    cell_a.border = thin_border
    
    for col_pos, h_name in enumerate(target_headers, 2):
        cell = ws.cell(row=3, column=col_pos, value=h_name)
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border

    current_row = 4
    no_counter = 1
    for _, row in df_mw_raw.iterrows():
        if row.dropna().empty:
            continue
            
        ws.row_dimensions[current_row].height = 20
        c_no = ws.cell(row=current_row, column=1, value=no_counter)
        c_no.alignment = Alignment(horizontal='center', vertical='center')
        c_no.border = thin_border
        
        for col_pos, h_name in enumerate(target_headers, 2):
            cell = ws.cell(row=current_row, column=col_pos)
            mapped_col = col_mapping.get(h_name)
            
            if mapped_col and mapped_col in row and not pd.isna(row[mapped_col]):
                val = row[mapped_col]
                if isinstance(val, pd.Timestamp):
                    val = val.strftime('%Y-%m-%d')
                elif isinstance(val, str) and len(val) >= 10 and '00:00:00' in val:
                    val = val.split()[0]
                    
                cell.value = val
                if isinstance(val, (int, float)):
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.value = ""
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
            cell.border = thin_border
            
        current_row += 1
        no_counter += 1

    ws.row_dimensions[current_row].height = 25
    
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
    c_sum = ws.cell(row=current_row, column=1, value="합계")
    c_sum.font = Font(bold=True)
    c_sum.alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(row=current_row, column=1).border = thin_border
    ws.cell(row=current_row, column=2).border = thin_border
    
    ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=4)
    c_cnt = ws.cell(row=current_row, column=3, value=f"댓수 : {count}")
    c_cnt.font = Font(bold=True)
    c_cnt.alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(row=current_row, column=3).border = thin_border
    ws.cell(row=current_row, column=4).border = thin_border
    
    ws.cell(row=current_row, column=5).border = thin_border
    ws.cell(row=current_row, column=6).border = thin_border
    
    c_g = ws.cell(row=current_row, column=7, value="총 실 수령액 :")
    c_g.font = Font(bold=True)
    c_g.alignment = Alignment(horizontal='center', vertical='center')
    c_g.border = thin_border
    
    c_h = ws.cell(row=current_row, column=8, value=total_pdf)
    c_h.font = Font(bold=True)
    c_h.number_format = '#,##0'
    c_h.alignment = Alignment(horizontal='right', vertical='center')
    c_h.border = thin_border
    
    c_i = ws.cell(row=current_row, column=9, value="총 청구 금액 :")
    c_i.font = Font(bold=True)
    c_i.alignment = Alignment(horizontal='center', vertical='center')
    c_i.border = thin_border
    
    c_j = ws.cell(row=current_row, column=10, value=total_excel)
    c_j.font = Font(bold=True)
    c_j.number_format = '#,##0'
    c_j.alignment = Alignment(horizontal='right', vertical='center')
    c_j.border = thin_border
    
    c_k = ws.cell(row=current_row, column=11, value="총 차액 :")
    c_k.font = Font(bold=True)
    c_k.alignment = Alignment(horizontal='center', vertical='center')
    c_k.border = thin_border
    
    c_l = ws.cell(row=current_row, column=12, value=total_diff)
    c_l.font = Font(bold=True)
    c_l.number_format = '#,##0'
    c_l.alignment = Alignment(horizontal='right', vertical='center')
    c_l.border = thin_border
    
    ws.merge_cells(start_row=current_row, start_column=13, end_row=current_row, end_column=14)
    c_mn = ws.cell(row=current_row, column=13, value="*부가세포함")
    c_mn.font = Font(bold=True)
    c_mn.alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(row=current_row, column=13).border = thin_border
    ws.cell(row=current_row, column=14).border = thin_border

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row > 3 and cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, month_str

# ────────────────────────────────────────────────────────
# 2️⃣ [모드 2] 쿠폰 보증 비교 관련 로직
# ────────────────────────────────────────────────────────
def load_excel_coupon_a(uploaded_file):
    df = read_excel_smart_header(uploaded_file)
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
    df = read_excel_smart_header(uploaded_file)
    col_car = get_col_by_idx_or_name(df, 6, ['차량번호', 'CAR', 'VEHICLE'])
    col_total = get_col_by_idx_or_name(df, 18, ['합계금액', '합계', 'TOTAL'])
    
    if col_total: df[col_total] = pd.to_numeric(df[col_total], errors='coerce').fillna(0)
    
    b_groups = defaultdict(list)
    for _, row in df.iterrows():
        car_no = str(row[col_car]).strip() if col_car else 'Unknown'
        if car_no and car_no != 'nan':
            b_groups[car_no].append(round_half_up(row[col_total]) if col_total else 0)
    return b_groups

def create_coupon_excel_report(uploaded_file_a, uploaded_file_b, count, total_b, total_a, total_diff):
    df_a_raw = read_excel_smart_header(uploaded_file_a)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "쿠폰 청구 현황"
    ws.print_title_rows = '1:3'
    
    month_str = "8월"
    file_b_name = getattr(uploaded_file_b, 'name', '')
    m_fn = re.search(r'20\d{2}(\d{2})', file_b_name)
    
    if m_fn:
        month_str = f"{int(m_fn.group(1))}월"
    else:
        for col in df_a_raw.columns:
            if any(keyword in str(col) for keyword in ['일자', 'DATE', '승인', '청구', '입고', '출고']):
                sample_dates = df_a_raw[col].dropna().astype(str).tolist()
                for d in sample_dates:
                    m = re.search(r'-(\d{2})-', d) or re.search(r'/(\d{2})/', d)
                    if m:
                        month_str = f"{int(m.group(1))}월"
                        break
                if month_str != "8월":
                    break

    ws.merge_cells('A1:N1')
    ws['A1'] = f"{month_str} 쿠폰 청구 현황"
    ws['A1'].font = Font(size=22, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    header_font = Font(size=10, bold=True)
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    headers = list(df_a_raw.columns)
    ws.row_dimensions[3].height = 25
    for col_idx, h_name in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=str(h_name))
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border

    current_row = 4
    for _, row in df_a_raw.iterrows():
        ws.row_dimensions[current_row].height = 20
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=current_row, column=col_idx)
            if pd.isna(val):
                cell.value = ""
            else:
                cell.value = val
                
            cell.border = thin_border
            if isinstance(val, (int, float)):
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')
        current_row += 1

    current_row += 2
    ws.row_dimensions[current_row].height = 30
    
    c_lbl1 = ws.cell(row=current_row, column=3, value="댓수 :")
    c_lbl1.font = Font(size=14, bold=True)
    c_lbl1.alignment = Alignment(horizontal='right', vertical='center')
    
    c_val1 = ws.cell(row=current_row, column=4, value=count)
    c_val1.font = Font(size=14, bold=True)
    c_val1.alignment = Alignment(horizontal='center', vertical='center')
    
    c_unit1 = ws.cell(row=current_row, column=5, value="대")
    c_unit1.font = Font(size=14, bold=True)
    c_unit1.alignment = Alignment(horizontal='left', vertical='center')

    c_lbl2 = ws.cell(row=current_row, column=7, value="총 청구 금액 :")
    c_lbl2.font = Font(size=14, bold=True)
    c_lbl2.alignment = Alignment(horizontal='right', vertical='center')

    c_val2 = ws.cell(row=current_row, column=9, value=total_b)
    c_val2.font = Font(size=14, bold=True)
    c_val2.number_format = '#,##0'
    c_val2.alignment = Alignment(horizontal='right', vertical='center')

    c_unit2 = ws.cell(row=current_row, column=10, value="원")
    c_unit2.font = Font(size=14, bold=True)
    c_unit2.alignment = Alignment(horizontal='left', vertical='center')

    c_vat1 = ws.cell(row=current_row, column=11, value="VAT 포함")
    c_vat1.font = Font(size=10, bold=True)
    c_vat1.alignment = Alignment(horizontal='left', vertical='center')

    current_row += 2

    ws.row_dimensions[current_row].height = 30
    
    c_lbl3 = ws.cell(row=current_row, column=3, value="차액 :")
    c_lbl3.font = Font(size=14, bold=True)
    c_lbl3.alignment = Alignment(horizontal='right', vertical='center')

    c_val3 = ws.cell(row=current_row, column=4, value=total_diff)
    c_val3.font = Font(size=14, bold=True)
    c_val3.number_format = '#,##0'
    c_val3.alignment = Alignment(horizontal='center', vertical='center')

    c_unit3 = ws.cell(row=current_row, column=5, value="원")
    c_unit3.font = Font(size=14, bold=True)
    c_unit3.alignment = Alignment(horizontal='left', vertical='center')

    c_lbl4 = ws.cell(row=current_row, column=7, value="총 입금 금액 :")
    c_lbl4.font = Font(size=14, bold=True)
    c_lbl4.alignment = Alignment(horizontal='right', vertical='center')

    c_val4 = ws.cell(row=current_row, column=9, value=total_a)
    c_val4.font = Font(size=14, bold=True)
    c_val4.number_format = '#,##0'
    c_val4.alignment = Alignment(horizontal='right', vertical='center')

    c_unit4 = ws.cell(row=current_row, column=10, value="원")
    c_unit4.font = Font(size=14웹 브라우저에서 바로 열어 간편하게 붙여넣고 비교할 수 있는 **단일 HTML 파일(공임코드 비교 도구)**입니다. 

아래 코드를 복사해 메모장에 붙여넣은 뒤, **`공임코드_비교.html`**로 저장하고 더블 클릭하여 실행하면 됩니다.

```html
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>공임코드 비교</title>
<style>
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    margin: 20px auto;
    max-width: 1000px;
    padding: 0 15px;
    color: #333;
  }
  h2 {
    text-align: center;
    margin-bottom: 20px;
  }
  .input-container {
    display: flex;
    gap: 20px;
    margin-bottom: 15px;
  }
  .input-box {
    flex: 1;
    display: flex;
    flex-direction: column;
  }
  .input-box label {
    font-weight: bold;
    margin-bottom: 8px;
  }
  textarea {
    width: 100%;
    height: 250px;
    padding: 10px;
    box-sizing: border-box;
    font-family: monospace;
    font-size: 13px;
    border: 1px solid #ccc;
    border-radius: 4px;
    resize: vertical;
    white-space: pre;
  }
  .btn-wrap {
    text-align: center;
    margin: 20px 0;
  }
  button {
    background-color: #007bff;
    color: white;
    border: none;
    padding: 12px 30px;
    font-size: 16px;
    font-weight: bold;
    border-radius: 4px;
    cursor: pointer;
  }
  button:hover {
    background-color: #0056b3;
  }
  .result-section {
    margin-top: 25px;
  }
  .summary {
    font-size: 15px;
    font-weight: bold;
    margin-bottom: 10px;
    padding: 10px;
    background-color: #f8f9fa;
    border-left: 4px solid #007bff;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 10px;
    font-size: 13px;
  }
  th, td {
    border: 1px solid #dee2e6;
    padding: 8px 12px;
    text-align: left;
  }
  th {
    background-color: #e9ecef;
  }
  .dup-code {
    color: #d9534f;
    font-weight: bold;
  }
  .no-data {
    color: #6c757d;
    font-style: italic;
  }
</style>
</head>
<body>

<h2>공임코드 비교</h2>

<div class="input-container">
  <div class="input-box">
    <label for="groupA">A 그룹 붙여넣기</label>
    <textarea id="groupA" placeholder="A 그룹 데이터를 붙여넣으세요."></textarea>
  </div>
  <div class="input-box">
    <label for="groupB">B 그룹 붙여넣기</label>
    <textarea id="groupB" placeholder="B 그룹 데이터를 붙여넣으세요."></textarea>
  </div>
</div>

<div class="btn-wrap">
  <button onclick="compareLaborCodes()">비교진행</button>
</div>

<div class="result-section" id="resultArea" style="display:none;">
  <div class="summary" id="summaryText"></div>
  <table>
    <thead>
      <tr>
        <th style="width: 18%;">중복 공임코드</th>
        <th style="width: 41%;">A 그룹 내용</th>
        <th style="width: 41%;">B 그룹 내용</th>
      </tr>
    </thead>
    <tbody id="resultBody"></tbody>
  </table>
</div>

<script>
function extractCode(line) {
  // 앞뒤 공백 및 보이지 않는 줄바꿈/특수공백 정리
  const cleanLine = line.trim();
  if (!cleanLine) return null;

  // 정규식: 문자 3개 - 문자 2개 - 문자 1~4개 패턴 검색
  const match = cleanLine.match(/([a-zA-Z0-9]{3}-[a-zA-Z0-9]{2}-[a-zA-Z0-9]{1,4})/);
  return match ? match[1] : null;
}

function parseGroup(text) {
  const lines = text.split('\n');
  const codeMap = new Map();

  lines.forEach(rawLine => {
    const code = extractCode(rawLine);
    if (code) {
      if (!codeMap.has(code)) {
        codeMap.set(code, []);
      }
      codeMap.get(code).push(rawLine.trim());
    }
  });

  return codeMap;
}

function compareLaborCodes() {
  const textA = document.getElementById('groupA').value;
  const textB = document.getElementById('groupB').value;

  const mapA = parseGroup(textA);
  const mapB = parseGroup(textB);

  const duplicates = [];
  mapA.forEach((linesA, code) => {
    if (mapB.has(code)) {
      duplicates.push({
        code: code,
        linesA: linesA,
        linesB: mapB.get(code)
      });
    }
  });

  const resultArea = document.getElementById('resultArea');
  const summaryText = document.getElementById('summaryText');
  const resultBody = document.getElementById('resultBody');

  resultBody.innerHTML = '';
  resultArea.style.display = 'block';

  if (duplicates.length === 0) {
    summaryText.innerText = '비교 결과: 중복된 공임코드가 없습니다.';
    resultBody.innerHTML = '<tr><td colspan="3" class="no-data" style="text-align:center;">일치하는 항목이 없습니다.</td></tr>';
    return;
  }

  summaryText.innerText = `총 ${duplicates.length}개의 중복 공임코드를 발견했습니다.`;

  duplicates.forEach(item => {
    const row = document.createElement('tr');

    const tdCode = document.createElement('td');
    tdCode.className = 'dup-code';
    tdCode.innerText = item.code;

    const tdA = document.createElement('td');
    tdA.innerHTML = item.linesA.join('<br>');

    const tdB = document.createElement('td');
    tdB.innerHTML = item.linesB.join('<br>');

    row.appendChild(tdCode);
    row.appendChild(tdA);
    row.appendChild(tdB);
    resultBody.appendChild(row);
  });
}
</script>

</body>
</html>
