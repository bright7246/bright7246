import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import pdfplumber
import re
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# 사이드바 초기 닫힘 상태로 설정
st.set_page_config(
    page_title="IRON WARRANTY", 
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 카카오톡 미리보기(OG Tag) 메타데이터 및 브라우저 자동 번역 충돌 방지 태그
st.markdown(
    """
    <head>
      <meta property="og:title" content="IRON WARRANTY">
      <meta property="og:description" content="아이언모터스 보증팀 지원 프로그램">
      <meta property="og:image" content="https://dummyimage.com/1200x630/0ea5e9/ffffff.png&text=IRON+WARRANTY">
      <meta property="og:image:width" content="1200">
      <meta property="og:image:height" content="630">
      <meta property="og:type" content="website">
      <meta name="google" content="notranslate">
    </head>
    <div translate="no"></div>
    """,
    unsafe_allow_html=True
)

# ────────────────────────────────────────────────────────
# 🎨 기본 UI 스타일링
# ────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    div[data-testid="stHorizontalBlock"] div.stButton > button {
        height: 65px !important;
        border-radius: 10px !important;
    }
    div[data-testid="stHorizontalBlock"] div.stButton > button p {
        font-size: 18px !important;
        font-weight: 700 !important;
        line-height: 1.3 !important;
    }
    div.share-btn-wrap div.stButton > button {
        height: 38px !important;
        min-height: 38px !important;
        padding: 4px 12px !important;
        border-radius: 8px !important;
        margin-top: 6px !important;
    }
    div.share-btn-wrap div.stButton > button p {
        font-size: 13px !important;
        font-weight: 600 !important;
        line-height: 1.2 !important;
    }
    button[kind="primary"], div.stDownloadButton > button {
        background-color: #0ea5e9 !important;
        border-color: #0ea5e9 !important;
        color: white !important;
    }
    button[kind="primary"]:hover, div.stDownloadButton > button:hover {
        background-color: #0284c7 !important;
        border-color: #0284c7 !important;
        color: white !important;
    }
    button[kind="primary"]:active, button[kind="primary"]:focus,
    div.stDownloadButton > button:active, div.stDownloadButton > button:focus {
        background-color: #0369a1 !important;
        border-color: #0369a1 !important;
        box-shadow: 0 0 0 0.2rem rgba(14, 165, 233, 0.4) !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

APP_URL = "https://bright7246-cg4cltxcy2z2ksgwbsod2p.streamlit.app"

# ────────────────────────────────────────────────────────
# 🔗 공유하기 다이얼로그
# ────────────────────────────────────────────────────────
@st.dialog("📱 프로그램 공유하기")
def share_modal():
    st.write("스마트폰 카메라로 아래 QR 코드를 비추면 즉시 접속할 수 있습니다.")
    qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=250x250&data={APP_URL}"
    col_img1, col_img2, col_img3 = st.columns([1, 2, 1])
    with col_img2:
        st.image(qr_url, caption="접속용 QR 코드", use_container_width=True)
    
    st.text_input("프로그램 접속 주소", value=APP_URL, disabled=True)
    
    copy_btn_html = f"""
    <div style="display: flex; justify-content: center; margin-top: 10px;">
        <button id="copy-btn" onclick="copyAppUrl()" style="
            background-color: #0ea5e9;
            color: white;
            border: none;
            padding: 10px 24px;
            font-size: 15px;
            font-weight: bold;
            border-radius: 8px;
            cursor: pointer;
            width: 100%;
            transition: 0.2s;
        ">📋 주소 복사하기</button>
    </div>
    <script>
    function copyAppUrl() {{
        navigator.clipboard.writeText('{APP_URL}').then(function() {{
            const btn = document.getElementById('copy-btn');
            btn.innerText = '✅ 복사 완료!';
            btn.style.backgroundColor = '#10b981';
            setTimeout(function() {{
                btn.innerText = '📋 주소 복사하기';
                btn.style.backgroundColor = '#0ea5e9';
            }}, 2000);
        }}).catch(function(err) {{
            alert('복사에 실패했습니다. 주소를 직접 드래그하여 복사해 주세요.');
        }});
    }}
    </script>
    """
    st.components.v1.html(copy_btn_html, height=65)

head_col1, head_col2 = st.columns([8.5, 1.5])
with head_col1:
    st.title("📊 아이언모터스 보증팀 지원 프로그램")
with head_col2:
    st.markdown('<div class="share-btn-wrap">', unsafe_allow_html=True)
    if st.button("🔗 공유 / QR", use_container_width=True):
        share_modal()
    st.markdown('</div>', unsafe_allow_html=True)

# ────────────────────────────────────────────────────────
# 🗂️ 상단 가로 메뉴 버튼 UI
# ────────────────────────────────────────────────────────
if "current_mode" not in st.session_state:
    st.session_state.current_mode = "MW 보증 비교"

nav_col1, nav_col2, nav_col3 = st.columns(3)

with nav_col1:
    btn_mw = st.button(
        "📋 MW 보증 비교 (PDF vs 엑셀)", 
        use_container_width=True, 
        type="primary" if st.session_state.current_mode == "MW 보증 비교" else "secondary"
    )
    if btn_mw:
        st.session_state.current_mode = "MW 보증 비교"
        st.rerun()

with nav_col2:
    btn_coupon = st.button(
        "🚗 쿠폰 보증 비교 (엑셀 vs 엑셀)", 
        use_container_width=True, 
        type="primary" if st.session_state.current_mode == "쿠폰 보증 비교" else "secondary"
    )
    if btn_coupon:
        st.session_state.current_mode = "쿠폰 보증 비교"
        st.rerun()

with nav_col3:
    btn_labor = st.button(
        "🔧 공임코드 비교 (중복 작업 검증)", 
        use_container_width=True, 
        type="primary" if st.session_state.current_mode == "공임코드 비교" else "secondary"
    )
    if btn_labor:
        st.session_state.current_mode = "공임코드 비교"
        st.rerun()

st.divider()

mode = st.session_state.current_mode

# ────────────────────────────────────────────────────────
# 🛠️ [공통 함수]
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

def find_col_smart(df, keywords, fallback_idx=None):
    for kw in keywords:
        kw_clean = str(kw).replace(" ", "").upper()
        for col in df.columns:
            col_clean = str(col).replace(" ", "").upper()
            if kw_clean in col_clean:
                return col
    if fallback_idx is not None and fallback_idx < len(df.columns):
        return df.columns[fallback_idx]
    return None

def round_half_up(value):
    return int(value + 0.5)

# ────────────────────────────────────────────────────────
# 📊 [컴포넌트 렌더링]
# ────────────────────────────────────────────────────────
def render_side_by_side_tables(df_main, df_diff=None, diff_title="🚨 차액 리스트 (100원 이상)"):
    main_headers = ["No."] + list(df_main.columns)
    
    main_tbody = []
    for idx, row in df_main.iterrows():
        is_total = "총합계" in str(row.iloc[0])
        tr_class = ' class="total-row"' if is_total else ''
        main_tbody.append(f'<tr{tr_class}>')
        main_tbody.append(f'<td class="col-no" onclick="copyCell(this)">{idx}</td>')
        for c_idx, val in enumerate(row):
            val_str = str(val)
            align_class = "col-id" if c_idx == 0 else ("col-diff" if c_idx == len(row)-1 else "col-amt")
            main_tbody.append(f'<td class="{align_class}" onclick="copyCell(this)">{val_str}</td>')
        main_tbody.append('</tr>')

    diff_section = ""
    if df_diff is not None:
        diff_headers = ["No."] + list(df_diff.columns)
        if len(df_diff) > 0:
            diff_tbody = []
            for idx, row in df_diff.iterrows():
                is_total = "총합계" in str(row.iloc[0])
                tr_class = ' class="total-row"' if is_total else ''
                diff_tbody.append(f'<tr{tr_class}>')
                diff_tbody.append(f'<td class="col-no" onclick="copyCell(this)">{idx}</td>')
                diff_tbody.append(f'<td class="col-id" onclick="copyCell(this)">{row.iloc[0]}</td>')
                diff_tbody.append(f'<td class="col-type" onclick="copyCell(this)">{row.iloc[1]}</td>')
                diff_tbody.append(f'<td class="col-desc" onclick="copyCell(this)">{row.iloc[2]}</td>')
                diff_color = "" if is_total else " diff-red"
                diff_tbody.append(f'<td class="col-diff{diff_color}" onclick="copyCell(this)">{row.iloc[3]}</td>')
                diff_tbody.append('</tr>')
            
            diff_section = f"""
            <div class="table-card">
                <div class="card-title">{diff_title}</div>
                <div class="scroll-wrap">
                    <table class="compact-table">
                        <thead>
                            <tr>
                                <th class="col-no">{diff_headers[0]}</th>
                                <th class="col-id">{diff_headers[1]}</th>
                                <th class="col-type">{diff_headers[2]}</th>
                                <th class="col-desc">{diff_headers[3]}</th>
                                <th class="col-diff">{diff_headers[4]}</th>
                            </tr>
                        </thead>
                        <tbody>{''.join(diff_tbody)}</tbody>
                    </table>
                </div>
            </div>
            """
        else:
            diff_section = f"""
            <div class="table-card">
                <div class="card-title">{diff_title}</div>
                <div style="padding: 16px; color: #10b981; font-weight: bold; background: #0f172a; border-radius: 6px; border: 1px solid #334155;">
                    ✅ 차액 100원 이상 발생 항목이 없습니다.
                </div>
            </div>
            """

    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8" />
    <style>
      * {{
        box-sizing: border-box;
        margin: 0;
        padding: 0;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      }}
      body {{
        background-color: transparent;
        color: #f8fafc;
        overflow-x: hidden;
      }}
      .flex-container {{
        display: flex;
        gap: 20px;
        align-items: flex-start;
        justify-content: flex-start;
      }}
      .table-card {{
        flex: 0 0 auto;
      }}
      .card-title {{
        font-
