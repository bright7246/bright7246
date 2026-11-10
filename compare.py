from collections import OrderedDict, defaultdict
import datetime
import calendar
import copy
import html as html_lib
import io
import json
import os
import re
import base64
import requests
import pdfplumber
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# ────────────────────────────────────────────────────────
# 🚗 볼보 공식 로고 이미지
# ────────────────────────────────────────────────────────
def get_brand_logo():
    if os.path.exists("logo.png"):
        try:
            with open("logo.png", "rb") as img_f:
                return f"data:image/png;base64,{base64.b64encode(img_f.read()).decode('utf-8')}"
        except Exception:
            pass
    svg_icon = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
      <circle cx="50" cy="50" r="38" fill="none" stroke="#111111" stroke-width="8"/>
      <path d="M 68 32 L 88 12 M 72 12 L 88 12 L 88 28" fill="none" stroke="#111111" stroke-width="8" stroke-linecap="square" stroke-linejoin="miter"/>
      <rect x="18" y="42" width="64" height="16" fill="#ffffff"/>
      <text x="50" y="55" font-family="'Arial Black', sans-serif" font-weight="900" font-size="14" fill="#111111" text-anchor="middle" letter-spacing="2">VOLVO</text>
    </svg>"""
    return f"data:image/svg+xml;utf8,{svg_icon}"

LOGO_DATA_URI = get_brand_logo()

st.set_page_config(
    page_title="IRON WARRANTY",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="collapsed",
)

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
    unsafe_allow_html=True,
)

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
    /* 버튼 개별 스타일: 버튼 key 로 자동 생성되는 st-key-<key> 클래스를 이용 */
    div[data-testid="stHorizontalBlock"] div[class*="st-key-sharebtn"] button {
        height: 38px !important;
        min-height: 38px !important;
        padding: 4px 12px !important;
        border-radius: 8px !important;
        margin-top: 6px !important;
    }
    div[data-testid="stHorizontalBlock"] div[class*="st-key-sharebtn"] button p {
        font-size: 13px !important;
        font-weight: 600 !important;
        line-height: 1.2 !important;
    }
    div[data-testid="stHorizontalBlock"] div[class*="st-key-swbtn_"] button {
        height: 38px !important;
        min-height: 38px !important;
        padding: 4px 8px !important;
        border-radius: 8px !important;
        margin-top: 6px !important;
        background-color: #0f172a !important;
        border: 1px solid #334155 !important;
        color: #f8fafc !important;
        white-space: nowrap !important;
    }
    div[data-testid="stHorizontalBlock"] div[class*="st-key-swbtn_"] button p {
        font-size: 13px !important;
        font-weight: 700 !important;
        line-height: 1.2 !important;
    }
    div[data-testid="stHorizontalBlock"] div[class*="st-key-swbtn_"] button:hover {
        background-color: #1e293b !important;
        border-color: #0ea5e9 !important;
        color: #38bdf8 !important;
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
    div[data-baseweb="input"] label p, div[data-baseweb="textarea"] label p, div.stTextInput label p, div.stTextArea label p, div.stSelectbox label p, div.stDateInput label p {
        font-size: 19px !important;
        font-weight: 700 !important;
        color: #f1f5f9 !important;
    }
    .custom-sub-label {
        font-size: 17px;
        font-weight: 700;
        color: #f1f5f9;
        margin-bottom: 8px;
    }
    div[data-testid="stExpander"] div[data-testid="stButton"] > button {
        height: 48px !important;
    }
    div[data-testid="stExpander"] div[data-testid="stButton"] > button p {
        font-size: 16px !important;
        font-weight: 700 !important;
    }
    .event-card {
        background-color: #0f172a;
        border: 1px solid #334155;
        border-radius: 8px;
        padding: 10px 14px;
        margin-bottom: 8px;
        min-height: 72px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .event-card-title {
        font-size: 14px;
        font-weight: 700;
        color: #f8fafc;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .event-card-memo {
        font-size: 12px;
        color: #94a3b8;
        margin-top: 4px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    div[data-testid="stHorizontalBlock"] div[class*="st-key-delev_"] button {
        height: 72px !important;
        min-height: 72px !important;
        border-radius: 8px !important;
        background-color: #1e293b !important;
        border: 1px solid #475569 !important;
        color: #ef4444 !important;
        padding: 0 !important;
        margin-bottom: 8px !important;
    }
    div[data-testid="stHorizontalBlock"] div[class*="st-key-delev_"] button p {
        font-size: 13px !important;
        font-weight: 700 !important;
    }
    div[data-testid="stHorizontalBlock"] div[class*="st-key-delev_"] button:hover {
        background-color: #7f1d1d !important;
        border-color: #ef4444 !important;
        color: #ffffff !important;
    }
    div[class*="st-key-cal_single_day"] {
        margin-top: 10px;
    }
    div[class*="st-key-cal_single_day"] label p {
        font-size: 15px !important;
        font-weight: 600 !important;
        color: #38bdf8 !important;
    }
    .brand-title-wrap {
        display: flex;
        align-items: center;
        gap: 14px;
    }
    .brand-logo-img {
        width: 48px;
        height: 48px;
        border-radius: 50%;
        background-color: #ffffff;
        padding: 3px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.3);
        flex-shrink: 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

APP_URL = "https://iron-warranty-app.streamlit.app"

# 화면 상단 제목 (서버별로 다르게 사용)
#   main  : 아이언모터스 보증팀 지원 프로그램
#   test  : TEST 용 서버
#   test2 : 클로드 테스트용 서버
APP_TITLE = "클로드 테스트용 서버"

# ────────────────────────────────────────────────────────
# ☁️ GitHub 영구 저장소 연동 모듈
# ────────────────────────────────────────────────────────
def get_github_auth():
    token = st.secrets.get("GITHUB_TOKEN", None)
    repo = st.secrets.get("GITHUB_REPO", None)
    if token and repo:
        headers = {
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json"
        }
        return repo, headers
    return None, None

def _local_load(filename, default_data):
    if os.path.exists(filename):
        try:
            with open(filename, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return copy.deepcopy(default_data)

def _local_save(filename, data):
    try:
        with open(filename, "w", encoding="utf-8") as f:
            f.write(json.dumps(data, ensure_ascii=False, indent=2))
        return True
    except Exception:
        return False

def _github_get(filename):
    """GitHub 에서 파일을 읽어옵니다.
    반환: (상태, 데이터, sha, 오류메시지)  상태 = "ok" / "missing" / "error" / "nolink"
    """
    repo, headers = get_github_auth()
    if not (repo and headers):
        return "nolink", None, None, "GitHub 연동 정보(GITHUB_TOKEN / GITHUB_REPO)가 없습니다."
    url = f"https://api.github.com/repos/{repo}/contents/{filename}"
    try:
        res = requests.get(url, headers=headers, timeout=8)
    except requests.RequestException as e:
        return "error", None, None, f"GitHub 연결 실패 ({type(e).__name__})"
    if res.status_code == 200:
        try:
            body = res.json()
            decoded = base64.b64decode(body.get("content", "")).decode("utf-8")
            return "ok", json.loads(decoded), body.get("sha"), ""
        except Exception as e:
            return "error", None, None, f"GitHub 데이터 해석 실패 ({type(e).__name__})"
    if res.status_code == 404:
        return "missing", None, None, ""
    return "error", None, None, f"GitHub 응답 오류 (코드 {res.status_code})"

@st.cache_data(ttl=30, show_spinner=False)
def load_shared_data(filename, default_data):
    """모든 접속자가 공유하는 데이터(S/W 이력, 캘린더)를 읽어옵니다.
    30초마다 GitHub 에서 최신본을 다시 읽어, 다른 사람이 등록한 내용도 반영됩니다."""
    status, data, _, _ = _github_get(filename)
    if status == "ok":
        return data
    if status == "missing":
        return copy.deepcopy(default_data)
    return _local_load(filename, default_data)

def update_shared_data(filename, default_data, mutate_fn):
    """저장 직전에 GitHub 최신본을 다시 읽고, 그 위에 변경사항(mutate_fn)만 적용해 저장합니다.
    → 여러 명이 동시에 사용해도 다른 사람이 등록한 내용을 덮어쓰지 않습니다.
    mutate_fn(data) 는 실제로 변경했으면 True, 변경할 대상이 없으면 False 를 반환합니다.
    반환: (결과, 메시지)  결과 = "success" / "warning" / "info" / "error"
    """
    repo, headers = get_github_auth()

    # GitHub 연동 정보가 없는 경우 : 서버 로컬 파일에만 저장
    if not (repo and headers):
        data = _local_load(filename, default_data)
        if not mutate_fn(data):
            load_shared_data.clear()
            return "info", "이미 삭제되었거나 변경된 항목입니다. 최신 내용으로 새로 고쳤습니다."
        _local_save(filename, data)
        load_shared_data.clear()
        return "warning", (
            "GitHub 연동 정보가 없어 서버에 임시로만 저장되었습니다. "
            "서버가 재시작되면 사라질 수 있습니다."
        )

    url = f"https://api.github.com/repos/{repo}/contents/{filename}"
    for _ in range(3):
        status, data, sha, err_msg = _github_get(filename)
        if status == "error":
            return "error", f"저장 실패 : {err_msg}. 잠시 후 다시 시도해 주세요."
        if status == "missing":
            data, sha = copy.deepcopy(default_data), None

        if not mutate_fn(data):
            load_shared_data.clear()
            return "info", "이미 삭제되었거나 변경된 항목입니다. 최신 내용으로 새로 고쳤습니다."

        content_str = json.dumps(data, ensure_ascii=False, indent=2)
        payload = {
            "message": f"Update {filename} via App",
            "content": base64.b64encode(content_str.encode("utf-8")).decode("utf-8"),
        }
        if sha:
            payload["sha"] = sha
        try:
            res = requests.put(url, headers=headers, json=payload, timeout=10)
        except requests.RequestException as e:
            return "error", f"저장 실패 : GitHub 연결 실패 ({type(e).__name__}). 잠시 후 다시 시도해 주세요."

        if res.status_code in (200, 201):
            _local_save(filename, data)
            load_shared_data.clear()
            return "success", ""
        if res.status_code in (409, 422):
            # 그 사이 다른 사람이 먼저 저장함 → 최신본을 다시 읽어 재시도
            continue
        return "error", f"저장 실패 : GitHub 응답 오류 (코드 {res.status_code}). 잠시 후 다시 시도해 주세요."

    return "error", "다른 사용자와 동시에 저장되어 저장하지 못했습니다. 다시 시도해 주세요."

def set_flash(level, message):
    """st.rerun() 후에도 보이도록 알림 메시지를 보관합니다."""
    st.session_state.flash_msg = (level, message)

def save_shared_and_rerun(filename, default_data, mutate_fn, success_msg):
    result, msg = update_shared_data(filename, default_data, mutate_fn)
    set_flash(result, success_msg if result == "success" else msg)
    st.rerun()

DEFAULT_SW_DATA = {
    "VOLVO": [{"version": "5.2.16", "date": datetime.date.today().strftime("%Y-%m-%d"), "memo": ""}],
    "V.ELEC": [{"version": "3.0.34", "date": datetime.date.today().strftime("%Y-%m-%d"), "memo": ""}],
    "POL": [{"version": "4.2.14", "date": datetime.date.today().strftime("%Y-%m-%d"), "memo": ""}],
}

DEFAULT_CALENDAR_DATA = [
    {
        "start_date": f"{datetime.date.today().year:04d}-{datetime.date.today().month:02d}-10",
        "end_date": f"{datetime.date.today().year:04d}-{datetime.date.today().month:02d}-10",
        "title": "MW 보증 청구 마감",
        "category": "MW 마감",
        "memo": "DMS 및 PDF 대조 완료 확인",
    }
]

# 매 실행마다 공유 데이터를 다시 읽음 (30초 캐시) → 다른 접속자의 변경사항도 반영
st.session_state.sw_history = load_shared_data("sw_data.json", DEFAULT_SW_DATA)
st.session_state.calendar_events = load_shared_data("calendar_data.json", DEFAULT_CALENDAR_DATA)

def latest_sw_version(car_key):
    """해당 차종의 최신 S/W 버전 (없으면 '-')"""
    history = st.session_state.sw_history.get(car_key)
    return history[0]["version"] if history else "-"

@st.dialog("🚗 S/W 버전 관리")
def manage_sw_dialog(car_key):
    st.markdown(f"### ⚙️ **{car_key}** S/W 버전 기록")
    st.info(f"현재 등록된 최신 버전 : **{latest_sw_version(car_key)}**")

    new_ver_input = st.text_input(
        "새 S/W 버전 입력",
        placeholder="예: 5.2.17",
        key=f"input_sw_ver_{car_key}",
    )

    new_memo_input = st.text_area(
        "메모 (특이사항 / 변경내역)",
        placeholder="특이사항이 있을 경우 작성해 주세요",
        height=75,
        key=f"input_sw_memo_{car_key}",
    )

    if st.button("신규 버전 등록하기", type="primary", use_container_width=True):
        if new_ver_input.strip():
            new_item = {
                "version": new_ver_input.strip(),
                "date": datetime.date.today().strftime("%Y-%m-%d"),
                "memo": new_memo_input.strip(),
            }

            def _add_version(data):
                data.setdefault(car_key, []).insert(0, new_item)
                return True

            save_shared_and_rerun(
                "sw_data.json", DEFAULT_SW_DATA, _add_version,
                f"✅ {car_key} S/W 버전이 영구 저장되었습니다!",
            )
        else:
            st.warning("⚠️ 버전을 입력해 주세요.")

    st.divider()
    st.markdown("#### 📜 최근 등록 이력 (최신 5개)")
    history_list = st.session_state.sw_history.get(car_key, [])[:5]

    if history_list:
        for idx, item in enumerate(history_list):
            c_h_info, c_h_del = st.columns([8.2, 1.8])
            with c_h_info:
                badge_text = " <span style='color:#38bdf8; font-weight:bold;'>(최신)</span>" if idx == 0 else ""
                memo_display = f"<div style='font-size:12px; color:#cbd5e1; margin-top:4px;'>📝 {item.get('memo')}</div>" if item.get("memo") else ""
                st.markdown(
                    f"**{idx + 1}. 버전: `{item['version']}`** {badge_text}  \n"
                    f"<span style='color:#94a3b8; font-size:12px;'>등록일자: {item['date']}</span>"
                    f"{memo_display}",
                    unsafe_allow_html=True,
                )
            with c_h_del:
                if st.button("삭제", key=f"del_sw_{car_key}_{idx}", use_container_width=True):
                    # 순번이 아니라 내용으로 찾아서 삭제 (그 사이 목록이 바뀌어도 안전)
                    def _del_version(data, target=dict(item)):
                        items = data.get(car_key, [])
                        if target in items:
                            items.remove(target)
                            return True
                        return False

                    save_shared_and_rerun(
                        "sw_data.json", DEFAULT_SW_DATA, _del_version,
                        f"🗑️ {car_key} S/W 버전 기록이 삭제되었습니다.",
                    )
            st.markdown(
                "<hr style='border:0; border-top:1px solid #334155; margin:6px 0;'>",
                unsafe_allow_html=True,
            )
    else:
        st.info("등록된 버전 이력이 없습니다.")

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

# ────────────────────────────────────────────────────────
# 🔝 상단 헤더
# ────────────────────────────────────────────────────────
SW_BUTTONS = [("VOLVO", "swbtn_volvo"), ("V.ELEC", "swbtn_velec"), ("POL", "swbtn_pol")]

head_col1, col_sw_lbl, *sw_cols, head_col2 = st.columns(
    [4.3, 0.6, 1.4, 1.4, 1.3, 1.0]
)

with head_col1:
    st.markdown(
        f"""
        <div class="brand-title-wrap">
          <img src="{LOGO_DATA_URI}" class="brand-logo-img" alt="VOLVO" />
          <h1 style="margin: 0; padding: 0; font-size: 1.85rem; font-weight: 800; color: #f8fafc;">{APP_TITLE}</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col_sw_lbl:
    st.markdown(
        '<div style="height:38px; margin-top:6px; display:flex; align-items:center; justify-content:center; background-color:#1e293b; border:1px solid #475569; border-radius:8px; font-weight:800; font-size:13px; color:#38bdf8;">S/W</div>',
        unsafe_allow_html=True,
    )

for sw_col, (car_key, btn_key) in zip(sw_cols, SW_BUTTONS):
    with sw_col:
        if st.button(f"{car_key}  {latest_sw_version(car_key)}", use_container_width=True, key=btn_key):
            manage_sw_dialog(car_key)

with head_col2:
    if st.button("🔗 공유 / QR", use_container_width=True, key="sharebtn"):
        share_modal()

today_date = datetime.date.today()

for _state_key, _default in {
    "current_mode": "MW 보증 비교",
    "reset_trigger": 0,
    "show_group_c": False,
    "cal_year": today_date.year,
    "cal_month": today_date.month,
}.items():
    if _state_key not in st.session_state:
        st.session_state[_state_key] = _default

# (모드 이름, 버튼 문구, 모드 전환 시 첨부파일 초기화 여부)
NAV_MODES = [
    ("MW 보증 비교", "📋 MW 보증 비교\n(PDF vs 엑셀)", True),
    ("쿠폰 보증 비교", "🚗 쿠폰 보증 비교\n(엑셀 vs 엑셀)", True),
    ("공임코드 비교", "🔧 공임코드 비교\n(중복 작업 검증)", True),
    ("캘린더", "📅 캘린더\n(보증/업무 일정 관리)", False),
]

for nav_col, (nav_mode, nav_label, nav_reset) in zip(st.columns(len(NAV_MODES)), NAV_MODES):
    with nav_col:
        is_current = st.session_state.current_mode == nav_mode
        if st.button(nav_label, use_container_width=True, type="primary" if is_current else "secondary"):
            if not is_current:
                st.session_state.current_mode = nav_mode
                if nav_reset:
                    st.session_state.reset_trigger += 1
                st.rerun()

st.divider()

# 저장 결과 알림 (st.rerun 이후에도 표시)
_flash = st.session_state.pop("flash_msg", None)
if _flash:
    _level, _message = _flash
    {"success": st.success, "warning": st.warning, "info": st.info}.get(_level, st.error)(_message)

mode = st.session_state.current_mode

# ────────────────────────────────────────────────────────
# 🛠️️ [공통 함수]
# ────────────────────────────────────────────────────────
def read_excel_smart_header(uploaded_file):
    uploaded_file.seek(0)
    df_raw = pd.read_excel(uploaded_file, header=None)
    header_row_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str)).upper()
        if "CLAIM" in row_str or "차량" in row_str or "공임" in row_str:
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
    try:
        val_float = float(value)
    except (ValueError, TypeError):
        return 0
    if val_float >= 0:
        return int(val_float + 0.5)
    else:
        return int(val_float - 0.5)

def parse_pdf_amount(val_str):
    if not val_str:
        return 0.0
    s = str(val_str).strip()
    is_negative = False
    if s.endswith("-"):
        is_negative = True
        s = s[:-1].strip()
    elif s.startswith("-"):
        is_negative = True
        s = s[1:].strip()

    s = s.replace(" ", "").replace(",", ".")
    try:
        val = float(s)
        return -val if is_negative else val
    except ValueError:
        return 0.0


# 3개 로더(MW 엑셀, 쿠폰 A, 쿠폰 B)가 공통으로 쓰는 열 검색 키워드
CLAIM_TYPE_KEYWORDS = ["CLAIM TYPE", "CLAIMTYPE", "청구유형", "클레임유형", "TYPE", "유형"]
DESC_KEYWORDS = ["제목", "TITLE", "SUBJECT", "내용", "작업내용", "수리내용", "DESCRIPTION", "REMARK", "비고"]

def find_type_desc_cols(df):
    """Claim Type 열과 제목(작업내용) 열을 찾습니다."""
    col_r = find_col_smart(df, CLAIM_TYPE_KEYWORDS, fallback_idx=17)
    col_v = find_col_smart(df, DESC_KEYWORDS, fallback_idx=21)
    return col_r, col_v

def to_numeric_cols(df, cols):
    """찾은 열들을 숫자로 변환 (숫자가 아니면 0)"""
    for c in cols:
        if c:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

def cell_str(row, col, default):
    """행에서 열 값을 문자열로 꺼냅니다. 열이 없으면 기본값."""
    return str(row.get(col, default)).strip() if col and col in row else default

def clean_text(value):
    """비어있거나 'nan' 이면 '-' 로 표시"""
    return value if value and value != "nan" else "-"

# ────────────────────────────────────────────────────────
# 📊 [테이블 렌더링]
# ────────────────────────────────────────────────────────
TABLE_COMMON_CSS = """
  * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
  body { background-color: transparent; color: #f8fafc; overflow-x: hidden; }
  /* 왼쪽 상세 표는 내용 폭 그대로, 오른쪽(차액 / 0원 리스트)은 남은 폭에 맞춰 줄어듦 */
  .flex-container { display: flex; gap: 14px; align-items: flex-start; justify-content: flex-start; flex-wrap: nowrap; width: 100%; }
  .table-card { flex: 0 0 auto; min-width: 0; }
  .card-title { font-size: 16px; font-weight: bold; margin-bottom: 8px; color: #f1f5f9; white-space: nowrap; }
  .scroll-wrap { max-height: 1120px; overflow-y: auto; overflow-x: auto; border: 1px solid #334155; border-radius: 6px; }
  .side-stack { display: flex; flex-direction: column; gap: 16px; flex: 1 1 0; min-width: 0; }
  .side-stack .table-card { width: 100%; }
  .side-stack .scroll-wrap { max-height: 520px; }
  .side-stack .compact-table { width: 100%; }
  .empty-box { padding: 14px; color: #10b981; font-weight: bold; background: #0f172a; border-radius: 6px; border: 1px solid #334155; font-size: 14px; }
  .scroll-wrap::-webkit-scrollbar { width: 8px; height: 8px; }
  .scroll-wrap::-webkit-scrollbar-track { background: #0f172a; }
  .scroll-wrap::-webkit-scrollbar-thumb { background: #334155; border-radius: 4px; }
  .compact-table { border-collapse: collapse; width: max-content; font-size: 13.5px; user-select: text; }
  .compact-table thead th { position: sticky; top: 0; background-color: #1e293b; color: #ffffff; padding: 8px 9px; font-weight: 700; border-bottom: 2px solid #475569; border-right: 1px solid #334155; white-space: nowrap; z-index: 2; }
  .compact-table tbody td { padding: 7px 9px; border-bottom: 1px solid #334155; border-right: 1px solid #334155; white-space: nowrap; transition: background-color 0.15s ease; }
  
  .compact-table tbody td.copyable { cursor: pointer; }
  .compact-table tbody td.copyable:hover { background-color: rgba(14, 165, 233, 0.25) !important; }
  
  .compact-table tbody td.state-1 { background-color: rgba(14, 165, 233, 0.45) !important; color: #ffffff !important; }
  .compact-table tbody td.state-2 { background-color: rgba(239, 68, 68, 0.45) !important; color: #ffffff !important; }
  
  .total-row { background-color: #0f172a !important; font-weight: bold; color: #38bdf8 !important; }
  .diff-red { color: #ef4444 !important; font-weight: bold; }
  .col-no { width: 1%; min-width: 38px; text-align: center; font-weight: bold; }
  .col-id { width: 1%; min-width: 92px; text-align: center; }
  .col-amt { min-width: 108px; text-align: right; }
  .col-type { width: 1%; min-width: 80px; text-align: center; color: #38bdf8; }
  /* 제목 : 남은 폭을 모두 쓰고, 넘치면 … 으로 줄임 (마우스를 올리면 전체 제목 표시) */
  .col-desc { min-width: 140px; max-width: 0; text-align: left; overflow: hidden; text-overflow: ellipsis; }
  th.col-desc { max-width: none; }
  .col-diff { width: 1%; min-width: 84px; text-align: right; }
  #toast { visibility: hidden; position: fixed; top: 14px; left: 50%; transform: translateX(-50%); background-color: #0284c7; color: #ffffff; padding: 9px 18px; border-radius: 6px; font-weight: bold; font-size: 14px; z-index: 999999; box-shadow: 0 4px 12px rgba(0,0,0,0.5); }
  #toast.show { visibility: visible; animation: fadein 0.2s, fadeout 0.3s 1.1s; }
  @keyframes fadein { from { opacity: 0; top: 0px; } to { opacity: 1; top: 14px; } }
  @keyframes fadeout { from { opacity: 1; top: 14px; } to { opacity: 0; top: 0px; } }
"""

TABLE_COMMON_JS = """
  function toggleCellColor(el) {
    const text = el.innerText.trim();
    if (!text || text === '-') return;

    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); } catch(e) { navigator.clipboard.writeText(text); }
    document.body.removeChild(ta);

    let currentState = el.getAttribute('data-click-state') || '0';
    if (currentState === '0') {
      el.setAttribute('data-click-state', '1');
      el.classList.add('state-1');
    } else if (currentState === '1') {
      el.setAttribute('data-click-state', '2');
      el.classList.remove('state-1');
      el.classList.add('state-2');
    } else {
      el.setAttribute('data-click-state', '0');
      el.classList.remove('state-2');
    }

    const toast = document.getElementById('toast');
    toast.innerText = '📋 복사 완료: ' + text;
    toast.className = 'show';
    setTimeout(function() { toast.className = ''; }, 1400);
  }
"""

def render_side_by_side_tables(df_main, df_diff, df_zero=None, zero_title=""):
    """왼쪽 : 상세 대조 내역 / 오른쪽 : 차액 리스트 (MW · 쿠폰 공용)
    df_zero 를 주면 차액 리스트 아래에 'PDF 0원 리스트' 를 추가로 표시합니다. (MW 전용)"""
    main_tbody = []
    for idx, row in df_main.iterrows():
        is_total = "총합계" in str(row.iloc[0])
        tr_class = ' class="total-row"' if is_total else ""
        main_tbody.append(f"<tr{tr_class}>")
        main_tbody.append(f'<td class="col-no">{idx}</td>')
        for c_idx, val in enumerate(row):
            val_str = str(val)
            if c_idx in [0, 1] and not is_total and val_str != "-":
                align_class = "col-id copyable" if c_idx == 0 else "col-amt copyable"
                main_tbody.append(f'<td class="{align_class}" onclick="toggleCellColor(this)">{val_str}</td>')
            else:
                align_class = "col-id" if c_idx == 0 else ("col-diff" if c_idx == len(row) - 1 else "col-amt")
                main_tbody.append(f'<td class="{align_class}">{val_str}</td>')
        main_tbody.append("</tr>")

    if df_diff is not None and len(df_diff) > 0:
        diff_tbody = []
        for idx, row in df_diff.iterrows():
            is_total = "총합계" in str(row.iloc[0])
            tr_class = ' class="total-row"' if is_total else ""
            id_cell = (
                f'<td class="col-id">{row.iloc[0]}</td>' if is_total
                else f'<td class="col-id copyable" onclick="toggleCellColor(this)">{row.iloc[0]}</td>'
            )
            diff_color = "" if is_total else " diff-red"
            diff_tbody.append(
                f'<tr{tr_class}><td class="col-no">{idx}</td>{id_cell}'
                f'<td class="col-type">{row.iloc[1]}</td>'
                f'<td class="col-desc" title="{html_lib.escape(str(row.iloc[2]), quote=True)}">{row.iloc[2]}</td>'
                f'<td class="col-diff{diff_color}">{row.iloc[3]}</td></tr>'
            )

        diff_th_html = "".join(
            f'<th class="{cls}">{name}</th>'
            for cls, name in zip(["col-no", "col-id", "col-type", "col-desc", "col-diff"], ["No."] + list(df_diff.columns))
        )
        diff_body = f"""
            <div class="scroll-wrap">
              <table class="compact-table">
                <thead><tr>{diff_th_html}</tr></thead>
                <tbody>{"".join(diff_tbody)}</tbody>
              </table>
            </div>"""
    else:
        diff_body = """
            <div class="empty-box">
              ✅ 차액 100원 이상 발생 항목이 없습니다.
            </div>"""

    zero_section = ""
    if df_zero is not None:
        if len(df_zero) > 0:
            zero_tbody = "".join(
                f'<tr><td class="col-no">{idx}</td>'
                f'<td class="col-id copyable" onclick="toggleCellColor(this)">{row.iloc[0]}</td>'
                f'<td class="col-type">{row.iloc[1]}</td>'
                f'<td class="col-desc" title="{html_lib.escape(str(row.iloc[2]), quote=True)}">{row.iloc[2]}</td></tr>'
                for idx, row in df_zero.iterrows()
            )
            zero_th_html = "".join(
                f'<th class="{cls}">{name}</th>'
                for cls, name in zip(["col-no", "col-id", "col-type", "col-desc"], ["No."] + list(df_zero.columns))
            )
            zero_body = f"""
            <div class="scroll-wrap">
              <table class="compact-table">
                <thead><tr>{zero_th_html}</tr></thead>
                <tbody>{zero_tbody}</tbody>
              </table>
            </div>"""
        else:
            zero_body = """
            <div class="empty-box">
              ✅ PDF 에 0원으로 등록된 항목이 없습니다.
            </div>"""
        zero_section = f"""
          <div class="table-card">
            <div class="card-title">{zero_title} ({len(df_zero)}건)</div>{zero_body}
          </div>"""

    main_th_html = "".join(
        f'<th class="{cls}">{name}</th>'
        for cls, name in zip(["col-no", "col-id", "col-amt", "col-amt", "col-diff"], ["No."] + list(df_main.columns))
    )

    full_html = f"""
      <!DOCTYPE html><html><head><meta charset="utf-8" />
      <style>{TABLE_COMMON_CSS}</style></head>
      <body>
        <div id="toast">📋 복사 완료!</div>
        <div class="flex-container">
          <div class="table-card">
            <div class="card-title">📋 상세 대조 내역</div>
            <div class="scroll-wrap">
              <table class="compact-table">
                <thead><tr>{main_th_html}</tr></thead>
                <tbody>{"".join(main_tbody)}</tbody>
              </table>
            </div>
          </div>
          <div class="side-stack">
            <div class="table-card">
              <div class="card-title">🚨 차액 리스트 (100원 이상)</div>{diff_body}
            </div>{zero_section}
          </div>
        </div>
        <script>{TABLE_COMMON_JS}</script>
      </body></html>
    """

    # 높이 : 왼쪽 상세 표와 오른쪽(차액 + 0원 리스트) 중 더 긴 쪽에 맞춤
    def side_block_height(df):
        if df is None:
            return 0
        rows = len(df) if len(df) > 0 else 1
        return min(520, (rows + 1) * 36) + 50

    right_height = side_block_height(df_diff if df_diff is not None else pd.DataFrame())
    if df_zero is not None:
        right_height += 20 + side_block_height(df_zero)
    calc_height = min(1160, max(320, len(df_main) * 36 + 90, right_height + 20))
    components.html(full_html, height=calc_height, scrolling=False)

# ────────────────────────────────────────────────────────
# 📑 [엑셀 서식 공통 도우미]
# ────────────────────────────────────────────────────────
def _thin_border(color):
    side = Side(style="thin", color=color)
    return Border(left=side, right=side, top=side, bottom=side)

BORDER_BLACK = _thin_border("000000")
BORDER_GRAY = _thin_border("A6A6A6")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
FMT_INT = "#,##0"

def xl_cell(ws, row, col, value=None, font=None, align=None, fill=None, border=None, fmt=None):
    """셀 하나에 값과 서식을 한 번에 지정합니다. (None 인 항목은 건드리지 않음)"""
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    if font is not None:
        c.font = font
    if align is not None:
        c.alignment = align
    if fill is not None:
        c.fill = fill
    if border is not None:
        c.border = border
    if fmt is not None:
        c.number_format = fmt
    return c

def amount_fmt(value):
    """소수점이 있는 금액은 소수 2자리까지, 아니면 정수로 표시"""
    return "#,##0.00" if isinstance(value, float) and value % 1 != 0 else FMT_INT

def write_report_title(ws, text):
    ws.merge_cells("A1:N1")
    xl_cell(ws, 1, 1, text, font=Font(size=22, bold=True), align=ALIGN_CENTER)
    ws.row_dimensions[1].height = 40

def autofit_columns(ws, padding, min_width):
    """4행 이후 데이터 길이에 맞춰 열 너비 자동 조정"""
    for col in ws.columns:
        max_len = 0
        for cell in col:
            if cell.row > 3 and cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + padding, min_width)

def workbook_to_bytes(wb):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ────────────────────────────────────────────────────────
# 1️⃣ [모드 1] MW 보증 비교
# ────────────────────────────────────────────────────────
def load_excel_mw(uploaded_file):
    df = read_excel_smart_header(uploaded_file)
    col_claim_no = find_col_smart(
        df, ["CLAIM NO", "CLAIM_NO", "클레임번호", "청구번호", "CLAIM"], fallback_idx=0
    )

    c_labor = find_col_smart(df, ["공임청구액", "공임청구", "공임 청구액"])
    c_labor_vat = find_col_smart(df, ["공임청구부가세", "공임청구 부가세", "공임부가세"])
    c_part = find_col_smart(df, ["부품청구액", "부품청구", "부품 청구액"])
    c_part_vat = find_col_smart(df, ["부품청구부가세", "부품청구 부가세", "부품부가세"])
    amount_cols = [c_labor, c_labor_vat, c_part, c_part_vat]
    to_numeric_cols(df, amount_cols)

    raw_labor_sum = int(round_half_up(df[c_labor].sum())) if c_labor else 0
    raw_part_sum = int(round_half_up(df[c_part].sum())) if c_part else 0

    total_series = pd.Series(0, index=df.index)
    for c in amount_cols:
        if c:
            total_series = total_series + df[c]
    df["Excel_Total"] = total_series.apply(round_half_up)

    col_r, col_v = find_type_desc_cols(df)
    col_job_no = find_col_smart(df, ["JOB NO", "JOB_NO", "JOB", "작업번호"], fallback_idx=2)
    col_ro_no = find_col_smart(df, ["R/O NO", "RO NO", "RO_NO", "RO", "차량번호", "주문번호"], fallback_idx=1)

    def amount_of(row, col):
        return int(round_half_up(row.get(col, 0))) if col and col in row else 0

    excel_groups = defaultdict(list)
    for _, row in df.iterrows():
        claim_no = cell_str(row, col_claim_no, "")
        if claim_no and claim_no != "nan":
            excel_groups[claim_no].append({
                "amount": int(row.get("Excel_Total", 0)),
                "claim_type": clean_text(cell_str(row, col_r, "-")),
                "v_desc": clean_text(cell_str(row, col_v, "-")),
                "labor": amount_of(row, c_labor),
                "labor_vat": amount_of(row, c_labor_vat),
                "part": amount_of(row, c_part),
                "part_vat": amount_of(row, c_part_vat),
                "job_no": cell_str(row, col_job_no, "1"),
                "ro_no": cell_str(row, col_ro_no, ""),
            })
    return excel_groups, raw_labor_sum, raw_part_sum

def load_pdf_mw(uploaded_file):
    pdf_groups = defaultdict(list)
    line_labour_total = 0.0
    line_material_total = 0.0
    exact_labour_summary = 0
    exact_material_summary = 0

    with pdfplumber.open(uploaded_file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            if page_num % 2 != 0:
                continue
            text = page.extract_text()
            if not text:
                continue
            lines = text.split("\n")
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

                    if parts and parts[-1] == "-" and len(parts) >= 2:
                        parts[-2] = parts[-2] + "-"
                        parts.pop()

                    try:
                        total_float = parse_pdf_amount(parts[-1])
                        pdf_total_with_vat = round_half_up(total_float * 1.1)

                        ro_no = parts[1] if len(parts) >= 2 else ""
                        job_no = parts[2] if len(parts) >= 3 else "1"
                        clm_type = parts[3] if len(parts) >= 4 else ""

                        l_cost = parse_pdf_amount(parts[-4]) if len(parts) >= 7 else (parse_pdf_amount(parts[-3]) if len(parts) >= 6 else 0.0)
                        m_cost = parse_pdf_amount(parts[-3]) if len(parts) >= 6 else 0.0
                        sublet = parse_pdf_amount(parts[-2]) if len(parts) >= 6 else 0.0

                        pdf_groups[rep_order].append({
                            "total_vat": pdf_total_with_vat,
                            "ro_no": ro_no,
                            "job_no": job_no,
                            "clm_type": clm_type,
                            "labour_cost": l_cost,
                            "material": m_cost,
                            "sublet": sublet,
                            "total": total_float
                        })

                        if len(parts) >= 6:
                            line_labour_total += l_cost
                            line_material_total += m_cost
                    except ValueError:
                        continue

        pages_to_check = pdf.pages[-2:] if len(pdf.pages) >= 2 else pdf.pages
        for p in reversed(pages_to_check):
            p_text = p.extract_text() or ""
            for line in p_text.split("\n"):
                l_clean = line.strip()
                if "***" in l_clean:
                    after_stars = l_clean.split("***")[-1]
                    parts = after_stars.split()

                    idx_pt = 0
                    merged_parts = []
                    while idx_pt < len(parts):
                        if parts[idx_pt] == "-" and merged_parts:
                            merged_parts[-1] += "-"
                        else:
                            merged_parts.append(parts[idx_pt])
                        idx_pt += 1

                    numeric_tokens = []
                    for pt in merged_parts:
                        cleaned = pt.replace(" ", "")
                        if re.match(r'^-?\d+,\d{2}-?$', cleaned):
                            numeric_tokens.append(cleaned)

                    if len(numeric_tokens) >= 2:
                        try:
                            exact_labour_summary = round_half_up(parse_pdf_amount(numeric_tokens[0]))
                            exact_material_summary = round_half_up(parse_pdf_amount(numeric_tokens[1]))
                            break
                        except Exception:
                            pass
            if exact_labour_summary != 0 and exact_material_summary != 0:
                break

    final_pdf_labour = exact_labour_summary if exact_labour_summary != 0 else round_half_up(line_labour_total)
    final_pdf_material = exact_material_summary if exact_material_summary != 0 else round_half_up(line_material_total)

    return pdf_groups, int(final_pdf_labour), int(final_pdf_material)

# ────────────────────────────────────────────────────────
# 📊 [MW 세부 대조 비교문서 (PDF vs DMS) 생성]
# ────────────────────────────────────────────────────────
def create_mw_comparison_matrix_excel(excel_groups, pdf_groups, month_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "비교 문서"

    # 📌 3번째 줄까지 틀고정 (4행부터 스크롤되도록 고정)
    ws.freeze_panes = "A4"

    font_main_title = Font(name="맑은 고딕", size=11, bold=True)
    font_header = Font(name="맑은 고딕", size=9, bold=True)
    font_data = Font(name="맑은 고딕", size=9)
    font_bold = Font(name="맑은 고딕", size=9, bold=True)

    fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_header_gray = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    fill_diff_col = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
    fill_diff_title = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    # 1. 메인 타이틀 바 (2행) : (병합범위, 제목, 글꼴, 배경, 시작열, 끝열)
    title_blocks = [
        ("B2:J2", "CREDIT PDF 6755", font_main_title, fill_yellow, 2, 10),
        ("L2:N2", "PDF - DMS", font_bold, fill_diff_title, 12, 14),
        ("P2:X2", "DMS FILE", font_main_title, fill_yellow, 16, 24),
    ]
    for merge_range, title, font, fill, start_col, end_col in title_blocks:
        ws.merge_cells(merge_range)
        xl_cell(ws, 2, start_col, title, font=font, align=ALIGN_CENTER, fill=fill)
        for col in range(start_col, end_col + 1):
            ws.cell(row=2, column=col).border = BORDER_GRAY

    # 2. 서브 헤더 바 (3행)
    header_groups = [
        (2, ["Claim No", "R/O No", "Job No", "Claim Type", "Labour cost", "Meterial", "sublet", "Total", "TOTAL"]),
        (12, ["공임차액", "부품차액", "총 차액"]),
        (16, ["Claim No", "R/O No", "Job No", "Claim Type", "공임청구액", "공임청구부가세", "부품청구액", "부품청구부가세", "TOTAL"]),
    ]
    for start_col, headers in header_groups:
        for i, h in enumerate(headers, start=start_col):
            xl_cell(ws, 3, i, h, font=font_header, align=ALIGN_CENTER, fill=fill_header_gray, border=BORDER_GRAY)

    # 3. 데이터 로우 작성 (4행부터)
    all_claims = sorted(set(excel_groups) | set(pdf_groups))
    cur_row = 4

    for claim in all_claims:
        p_list = pdf_groups.get(claim, [])
        e_list = excel_groups.get(claim, [])

        for idx in range(max(len(p_list), len(e_list))):
            p_data = p_list[idx] if idx < len(p_list) else None
            e_data = e_list[idx] if idx < len(e_list) else None

            pdf_vat_total = p_data["total_vat"] if p_data else 0
            dms_vat_total = e_data["amount"] if e_data else 0
            pdf_labor_vat = round_half_up(p_data["labour_cost"] * 1.1) if p_data else 0
            pdf_part_vat = round_half_up(p_data["material"] * 1.1) if p_data else 0
            dms_labor_total = (e_data["labor"] + e_data["labor_vat"]) if e_data else 0
            dms_part_total = (e_data["part"] + e_data["part_vat"]) if e_data else 0

            # 한쪽에만 있는 경우 없는 쪽은 0 으로 계산
            diff_values = [
                pdf_labor_vat - dms_labor_total,
                pdf_part_vat - dms_part_total,
                pdf_vat_total - dms_vat_total,
            ]

            # ◀ PDF 영역 (B~J)
            p_texts = [claim, p_data["ro_no"], p_data["job_no"], p_data["clm_type"]] if p_data else [""] * 4
            for i, v in enumerate(p_texts):
                xl_cell(ws, cur_row, 2 + i, v, font=font_data, align=ALIGN_CENTER, border=BORDER_GRAY)
            for i, key in enumerate(["labour_cost", "material", "sublet", "total"]):
                v = p_data[key] if p_data else ""
                xl_cell(ws, cur_row, 6 + i, v, font=font_data, align=ALIGN_RIGHT, border=BORDER_GRAY, fmt=amount_fmt(v))
            xl_cell(ws, cur_row, 10, pdf_vat_total if p_data else "", font=font_bold, align=ALIGN_RIGHT,
                    border=BORDER_GRAY, fmt=FMT_INT)

            # ■ 차액 영역 (L~N)
            for i, v in enumerate(diff_values):
                xl_cell(ws, cur_row, 12 + i, v, font=font_bold, align=ALIGN_RIGHT, fill=fill_diff_col,
                        border=BORDER_GRAY, fmt=FMT_INT)

            # ▶ DMS 영역 (P~X)
            e_texts = [claim, e_data["ro_no"], e_data["job_no"], e_data["claim_type"]] if e_data else [""] * 4
            for i, v in enumerate(e_texts):
                xl_cell(ws, cur_row, 16 + i, v, font=font_data, align=ALIGN_CENTER, border=BORDER_GRAY)
            for i, key in enumerate(["labor", "labor_vat", "part", "part_vat"]):
                xl_cell(ws, cur_row, 20 + i, e_data[key] if e_data else "", font=font_data, align=ALIGN_RIGHT,
                        border=BORDER_GRAY, fmt=FMT_INT)
            xl_cell(ws, cur_row, 24, dms_vat_total if e_data else "", font=font_bold, align=ALIGN_RIGHT,
                    border=BORDER_GRAY, fmt=FMT_INT)

            cur_row += 1

    col_widths = {
        "B": 11, "C": 12, "D": 8, "E": 11, "F": 13, "G": 14, "H": 8, "I": 14, "J": 13,
        "L": 12, "M": 12, "N": 12,
        "P": 11, "Q": 15, "R": 8, "S": 11, "T": 12, "U": 13, "V": 12, "W": 13, "X": 13
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    return workbook_to_bytes(wb)

def detect_month_from_filename(file_name):
    """파일명에서 월 찾기 : '10월' 표기 → 'YYYYMM' 숫자 순서. 못 찾으면 None"""
    m_direct = re.search(r"(\d{1,2})월", file_name)
    if m_direct:
        return f"{int(m_direct.group(1))}월"
    m_fn = re.search(r"20\d{2}(\d{2})", file_name)
    if m_fn and 1 <= int(m_fn.group(1)) <= 12:
        return f"{int(m_fn.group(1))}월"
    return None

def create_mw_excel_report(uploaded_file_mw, count, total_pdf, total_excel, total_diff):
    df_mw_raw = read_excel_smart_header(uploaded_file_mw)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "WARRANTY 수령내역"
    ws.print_title_rows = "1:3"

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
        "부품입금부가세": ["부품입금부가세", "부품입금 부가세", "부품승인부가세"],
    }
    target_headers = list(alias_dict)
    col_mapping = {th: find_col_smart(df_mw_raw, alias_dict[th]) for th in target_headers}

    month_str = detect_month_from_filename(getattr(uploaded_file_mw, "name", "")) or "X월"
    write_report_title(ws, f"{month_str} WARRANTY 수 령 내 역")

    header_font = Font(size=10, bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[3].height = 25
    for col_pos, h_name in enumerate(["No."] + target_headers, 1):
        xl_cell(ws, 3, col_pos, h_name, font=header_font, align=header_align, border=BORDER_BLACK)

    current_row = 4
    no_counter = 1
    for _, row in df_mw_raw.iterrows():
        if row.dropna().empty:
            continue
        ws.row_dimensions[current_row].height = 20
        xl_cell(ws, current_row, 1, no_counter, align=ALIGN_CENTER, border=BORDER_BLACK)

        for col_pos, h_name in enumerate(target_headers, 2):
            cell = xl_cell(ws, current_row, col_pos, border=BORDER_BLACK, align=ALIGN_CENTER)
            mapped_col = col_mapping.get(h_name)
            if mapped_col and mapped_col in row and not pd.isna(row[mapped_col]):
                val = row[mapped_col]
                if isinstance(val, pd.Timestamp):
                    val = val.strftime("%Y-%m-%d")
                elif isinstance(val, str) and len(val) >= 10 and "00:00:00" in val:
                    val = val.split()[0]
                cell.value = val
                if isinstance(val, (int, float)):
                    cell.number_format = FMT_INT
                    cell.alignment = ALIGN_RIGHT
            else:
                cell.value = ""
        current_row += 1
        no_counter += 1

    # 합계 행
    ws.row_dimensions[current_row].height = 25
    for start_col, end_col in [(1, 2), (3, 4), (13, 14)]:
        ws.merge_cells(start_row=current_row, start_column=start_col, end_row=current_row, end_column=end_col)
    for col in range(1, 15):
        ws.cell(row=current_row, column=col).border = BORDER_BLACK

    bold = Font(bold=True)
    summary_cells = [
        (1, "합계", None),
        (3, f"댓수 : {count}", None),
        (7, "총 실 수령액 :", None),
        (8, total_pdf, FMT_INT),
        (9, "총 청구 금액 :", None),
        (10, total_excel, FMT_INT),
        (11, "총 차액 :", None),
        (12, total_diff, FMT_INT),
        (13, "*부가세포함", None),
    ]
    for col, value, fmt in summary_cells:
        xl_cell(ws, current_row, col, value, font=bold, fmt=fmt,
                align=ALIGN_RIGHT if fmt else ALIGN_CENTER)

    autofit_columns(ws, padding=4, min_width=12)
    return workbook_to_bytes(wb), month_str

# ────────────────────────────────────────────────────────
# 2️⃣ [모드 2] 쿠폰 보증 비교
# ────────────────────────────────────────────────────────
def _group_coupon_rows(df, col_car, amount_fn):
    """차량번호별로 금액 / Claim Type / 제목을 묶습니다. (쿠폰 A·B 공용)"""
    col_r, col_v = find_type_desc_cols(df)
    groups = defaultdict(list)
    for _, row in df.iterrows():
        car_no = cell_str(row, col_car, "Unknown")
        if car_no and car_no != "nan":
            groups[car_no].append({
                "amount": amount_fn(row),
                "claim_type": clean_text(cell_str(row, col_r, "-")),
                "v_desc": clean_text(cell_str(row, col_v, "-")),
            })
    return groups

def load_excel_coupon_a(uploaded_file):
    df = read_excel_smart_header(uploaded_file)
    col_car = find_col_smart(
        df, ["차량번호", "CAR NO", "CAR_NO", "VEHICLE"], fallback_idx=3
    )
    col_part = find_col_smart(df, ["부품청구", "부품"], fallback_idx=8)
    col_labour = find_col_smart(df, ["공임청구", "공임"], fallback_idx=9)
    to_numeric_cols(df, [col_part, col_labour])
    df["Calc_Total"] = (
        ((df[col_part] if col_part else 0) + (df[col_labour] if col_labour else 0)) * 1.1
    ).apply(round_half_up)
    return _group_coupon_rows(df, col_car, lambda row: int(row["Calc_Total"]))

def load_excel_coupon_b(uploaded_file):
    df = read_excel_smart_header(uploaded_file)
    col_car = find_col_smart(
        df, ["차량번호", "CAR NO", "CAR_NO", "VEHICLE"], fallback_idx=6
    )
    col_total = find_col_smart(df, ["합계금액", "합계", "TOTAL"], fallback_idx=18)
    to_numeric_cols(df, [col_total])
    return _group_coupon_rows(df, col_car, lambda row: round_half_up(row[col_total]) if col_total else 0)

def detect_coupon_month(uploaded_file_b, df_a_raw):
    """월 판별 순서 : ① DMS 파일명(예: 쿠폰청구관리_20260818…) → ② 공지 파일의 날짜 열 → ③ 판별 불가 시 "X월" """
    m_fn = re.search(r"20\d{2}(\d{2})", getattr(uploaded_file_b, "name", ""))
    if m_fn and 1 <= int(m_fn.group(1)) <= 12:
        return f"{int(m_fn.group(1))}월"
    for col in df_a_raw.columns:
        if any(keyword in str(col) for keyword in ["일자", "DATE", "승인", "청구", "입고", "출고"]):
            for d in df_a_raw[col].dropna().astype(str).tolist():
                m = re.search(r"-(\d{2})-", d) or re.search(r"/(\d{2})/", d)
                if m and 1 <= int(m.group(1)) <= 12:
                    return f"{int(m.group(1))}월"
    return "X월"

def create_coupon_excel_report(uploaded_file_a, uploaded_file_b, count, total_b, total_a, total_diff):
    df_a_raw = read_excel_smart_header(uploaded_file_a)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "쿠폰 청구 현황"
    ws.print_title_rows = "1:3"

    month_str = detect_coupon_month(uploaded_file_b, df_a_raw)
    write_report_title(ws, f"{month_str} 쿠폰 청구 현황")

    header_font = Font(size=10, bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[3].height = 25
    for col_idx, h_name in enumerate(df_a_raw.columns, 1):
        xl_cell(ws, 3, col_idx, str(h_name), font=header_font, align=header_align, border=BORDER_BLACK)

    current_row = 4
    for _, row in df_a_raw.iterrows():
        ws.row_dimensions[current_row].height = 20
        for col_idx, val in enumerate(row, 1):
            is_num = isinstance(val, (int, float))
            xl_cell(ws, current_row, col_idx, "" if pd.isna(val) else val, border=BORDER_BLACK,
                    fmt=FMT_INT if is_num else None, align=ALIGN_RIGHT if is_num else ALIGN_CENTER)
        current_row += 1

    # 하단 요약 (2줄) : (라벨, 값, 값 정렬, 값 서식, 단위, VAT 표시)
    font_big = Font(size=14, bold=True)
    font_small = Font(size=10, bold=True)
    summary_rows = [
        [("댓수 :", count, ALIGN_CENTER, None, "대", False), ("총 청구 금액 :", total_b, ALIGN_RIGHT, FMT_INT, "원", True)],
        [("차액 :", total_diff, ALIGN_CENTER, FMT_INT, "원", False), ("총 입금 금액 :", total_a, ALIGN_RIGHT, FMT_INT, "원", True)],
    ]
    # 각 묶음의 (라벨 열, 값 열, 단위 열)
    positions = [(3, 4, 5), (7, 9, 10)]
    for summary in summary_rows:
        current_row += 2
        ws.row_dimensions[current_row].height = 30
        for (label, value, value_align, fmt, unit, show_vat), (c_lbl, c_val, c_unit) in zip(summary, positions):
            xl_cell(ws, current_row, c_lbl, label, font=font_big, align=ALIGN_RIGHT)
            xl_cell(ws, current_row, c_val, value, font=font_big, align=value_align, fmt=fmt)
            xl_cell(ws, current_row, c_unit, unit, font=font_big, align=ALIGN_LEFT)
            if show_vat:
                xl_cell(ws, current_row, c_unit + 1, "VAT 포함", font=font_small, align=ALIGN_LEFT)

    autofit_columns(ws, padding=5, min_width=14)
    return workbook_to_bytes(wb), month_str

# ────────────────────────────────────────────────────────
# 3️⃣ [모드 3] 공임코드 비교
# ────────────────────────────────────────────────────────
def parse_labor_lines(text):
    code_map = OrderedDict()
    if not text:
        return code_map
    pattern = re.compile(r"([A-Za-z0-9]{1,4}-[A-Za-z0-9]{1,4}-[A-Za-z0-9]{1,4})")
    for line in text.split("\n"):
        line_clean = line.strip()
        if not line_clean:
            continue
        match = pattern.search(line_clean)
        if match:
            code = match.group(1).upper()
            if code not in code_map:
                code_map[code] = []
            code_map[code].append(line_clean)
    return code_map

# ────────────────────────────────────────────────────────
# ⚡ [대조 계산] 같은 파일이면 결과를 저장해 두고 재사용 (버튼 클릭 시 재계산 방지)
# ────────────────────────────────────────────────────────
class NamedBytesIO(io.BytesIO):
    """업로드 파일처럼 .name 속성을 갖는 메모리 파일 (월 판별 등에 파일명 사용)"""
    def __init__(self, data, name=""):
        super().__init__(data)
        self.name = name

def _numbered_index(df):
    # 1, 2, 3 … 순번 + 마지막 총합계 행은 빈칸
    df.index = [str(i) for i in range(1, len(df))] + [""]
    return df

# 계산 결과 캐시 버전 : 대조/보고서 계산 방식을 바꾸면 이 숫자를 1 올립니다.
# (서버가 새 코드로 바뀌어도 예전 코드로 계산해 둔 결과를 재사용하지 않도록)
COMPUTE_VERSION = 2

def compare_groups(groups_1, groups_2, amount_key_1, amount_key_2, key_col, col_1, col_2, meta_fn):
    """두 자료를 키(주문번호/차량번호)별로 맞대어 금액을 대조합니다. (MW · 쿠폰 공용)
    - 같은 키가 여러 건이면 순서대로 1:1 매칭, 라벨에 (1), (2) … 표시
    - 차액 = 1번 자료 금액 - 2번 자료 금액, |차액| 100원 이상은 차액 리스트에 추가
    - meta_fn(item_1, item_2) → (Claim Type, 제목)
    """
    matched_results = []
    diff_over_100_results = []
    zero_1_results = []   # 1번 자료(MW 에서는 PDF)에 0원으로 등록된 건
    total_1_sum = 0
    total_2_sum = 0
    total_diff_100_sum = 0

    for key_item in sorted(set(groups_1) | set(groups_2)):
        items_1 = groups_1.get(key_item, [])
        items_2 = groups_2.get(key_item, [])
        max_len = max(len(items_1), len(items_2))

        for i in range(max_len):
            item_1 = items_1[i] if i < len(items_1) else None
            item_2 = items_2[i] if i < len(items_2) else None
            val_1 = item_1[amount_key_1] if item_1 else None
            val_2 = item_2[amount_key_2] if item_2 else None
            r_val, v_val = meta_fn(item_1, item_2)
            label_item = f"{key_item} ({i+1})" if max_len > 1 else key_item

            total_1_sum += val_1 or 0
            total_2_sum += val_2 or 0
            diff_val = (val_1 or 0) - (val_2 or 0)

            matched_results.append({
                key_col: label_item,
                col_1: f"{val_1:,}원" if val_1 is not None else "-",
                col_2: f"{val_2:,}원" if val_2 is not None else "-",
                "차액": f"{diff_val:,}원",
            })

            if val_1 == 0:
                zero_1_results.append({key_col: label_item, "Claim Type": r_val, "제목": v_val})

            if abs(diff_val) >= 100:
                total_diff_100_sum += diff_val
                diff_over_100_results.append({
                    key_col: label_item,
                    "Claim Type": r_val,
                    "제목": v_val,
                    "차액": f"{diff_val:,}원",
                })

    total_diff_sum = total_1_sum - total_2_sum
    total_cnt = len(matched_results)

    matched_results.append({
        key_col: "★ 총합계",
        col_1: f"{total_1_sum:,}원",
        col_2: f"{total_2_sum:,}원",
        "차액": f"{total_diff_sum:,}원",
    })
    res_df = _numbered_index(pd.DataFrame(matched_results))

    if diff_over_100_results:
        diff_over_100_results.append({
            key_col: "★ 총합계",
            "Claim Type": "-",
            "제목": "-",
            "차액": f"{total_diff_100_sum:,}원",
        })
        diff_df = _numbered_index(pd.DataFrame(diff_over_100_results))
    else:
        diff_df = pd.DataFrame(columns=[key_col, "Claim Type", "제목", "차액"])

    zero_df = pd.DataFrame(zero_1_results, columns=[key_col, "Claim Type", "제목"])
    zero_df.index = [str(i) for i in range(1, len(zero_df) + 1)]

    return {
        "res_df": res_df,
        "diff_df": diff_df,
        "zero_df": zero_df,
        "total_cnt": total_cnt,
        "total_diff_sum": total_diff_sum,
        "total_1_sum": total_1_sum,
        "total_2_sum": total_2_sum,
    }

@st.cache_data(max_entries=5, show_spinner=False)
def compute_mw(pdf_bytes, pdf_name, xlsx_bytes, xlsx_name, version=COMPUTE_VERSION):
    f_pdf = NamedBytesIO(pdf_bytes, pdf_name)
    f_xlsx = NamedBytesIO(xlsx_bytes, xlsx_name)

    excel_groups, raw_excel_labor, raw_excel_part = load_excel_mw(f_xlsx)
    pdf_groups, raw_pdf_labour, raw_pdf_material = load_pdf_mw(f_pdf)

    r = compare_groups(
        pdf_groups, excel_groups, "total_vat", "amount",
        "주문번호", "PDF 금액 (실 수령액)", "DMS 금액 (청구 금액)",
        meta_fn=lambda p, e: (
            e["claim_type"] if e else (p["clm_type"] if p else "-"),
            e["v_desc"] if e else "-",
        ),
    )

    excel_data, month_name = create_mw_excel_report(
        f_xlsx, r["total_cnt"], r["total_1_sum"], r["total_2_sum"], r["total_diff_sum"]
    )
    matrix_excel_data = create_mw_comparison_matrix_excel(excel_groups, pdf_groups, month_name)

    r["detail_metrics"] = [
        ("청구 공임 합계", raw_excel_labor), ("입금 공임 합계", raw_pdf_labour),
        ("청구 부품 합계", raw_excel_part), ("입금 부품 합계", raw_pdf_material),
    ]
    r["downloads"] = [
        (f"📥 [{month_name} WARRANTY 수령내역] 엑셀 보고서 다운로드",
         excel_data.getvalue(), f"{month_name}_WARRANTY_수령내역_보고서.xlsx"),
        (f"📊 [{month_name} MW 세부 대조 비교문서 (PDF vs DMS)] 다운로드",
         matrix_excel_data.getvalue(), f"{month_name}_MW_세부대조_비교문서.xlsx"),
    ]
    return r

@st.cache_data(max_entries=5, show_spinner=False)
def compute_coupon(a_bytes, a_name, b_bytes, b_name, version=COMPUTE_VERSION):
    f_a = NamedBytesIO(a_bytes, a_name)
    f_b = NamedBytesIO(b_bytes, b_name)

    r = compare_groups(
        load_excel_coupon_a(f_a), load_excel_coupon_b(f_b), "amount", "amount",
        "차량번호", "공지된 쿠폰 금액 ( 입금 금액 )", "DMS 쿠폰파일 ( 청구 금액 ) ",
        meta_fn=lambda a, b: (
            b["claim_type"] if b else (a["claim_type"] if a else "-"),
            b["v_desc"] if b else (a["v_desc"] if a else "-"),
        ),
    )

    excel_data, month_name = create_coupon_excel_report(
        f_a, f_b, r["total_cnt"], r["total_2_sum"], r["total_1_sum"], r["total_diff_sum"]
    )
    r["detail_metrics"] = []
    r["downloads"] = [
        (f"📥 [{month_name} 쿠폰 청구 현황] 엑셀 보고서 다운로드",
         excel_data.getvalue(), f"{month_name}_쿠폰_청구_현황_보고서.xlsx"),
    ]
    return r

# ────────────────────────────────────────────────────────
# 🖥️ 본문 화면 렌더링
# ────────────────────────────────────────────────────────
COMPARE_MODES = {
    "MW 보증 비교": {
        "prefix": "MW",
        "desc": "PDF(홀수페이지)와 엑셀",
        # (안내 문구, 업로드 라벨, 파일 형식)
        "uploads": [
            ("1. PDF 파일을 선택하세요 (예시 : DEALER_CREDITNOTE_6755)", "PDF 파일 업로드", "pdf"),
            ("2. 엑셀 파일을 선택하세요 (예시 : 보증청구현황 [ 항목 조정 가능 ]_2026)", "엑셀 파일 업로드", "xlsx"),
        ],
        "label_1": "PDF 총 합계 (입금 합계)",
        "zero_title": "📭 PDF 0원 리스트",
        "label_2": "DMS 총 합계 (청구 합계)",
        "compute": compute_mw,
    },
    "쿠폰 보증 비교": {
        "prefix": "쿠폰",
        "desc": "공지된 쿠폰 파일과 DMS 엑셀",
        "uploads": [
            ("1. 공지된 쿠폰 파일을 선택하세요 (예시 : IR_JJ_Aug)", "공지 쿠폰 파일 업로드", "xlsx"),
            ("2. DMS 쿠폰파일을 선택하세요 (예시 : 쿠폰청구관리_20260818085441)", "DMS 쿠폰 파일 업로드", "xlsx"),
        ],
        "label_1": "공지 쿠폰 (입금 합계)",
        "label_2": "DMS 쿠폰 (청구 합계)",
        "compute": compute_coupon,
    },
}

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

if mode in COMPARE_MODES:
    cfg = COMPARE_MODES[mode]

    st.markdown(
        f"### 🔍 {cfg['desc']}의"
        " 금액을 각각 계산 후 반올림 처리하여 대조합니다."
    )
    st.write("")

    left_col, right_col = st.columns([3.2, 6.8], gap="medium")

    with left_col:
        uploaded = []
        for n, (guide, upload_label, file_type) in enumerate(cfg["uploads"], 1):
            st.markdown(f'<div class="custom-sub-label">{guide}</div>', unsafe_allow_html=True)
            uploaded.append(st.file_uploader(
                upload_label,
                type=[file_type],
                key=f"file_{n}_{st.session_state.reset_trigger}",
                label_visibility="collapsed",
            ))
        f1, f2 = uploaded

        st.write("")
        if st.button("🗑️ 첨부파일 초기화", use_container_width=True):
            st.session_state.reset_trigger += 1
            st.rerun()

    if f1 and f2:
        with st.spinner(f"{cfg['prefix']} 보증 데이터 교차 대조 중..."):
            r = cfg["compute"](f1.getvalue(), f1.name, f2.getvalue(), f2.name, COMPUTE_VERSION)

        with right_col:
            if cfg.get("zero_title") and "zero_df" in r:
                render_side_by_side_tables(r["res_df"], r["diff_df"], r["zero_df"], cfg["zero_title"])
            else:
                render_side_by_side_tables(r["res_df"], r["diff_df"])

        with left_col:
            st.divider()
            st.subheader("📌 분석 요약 결과")
            sub_c1, sub_c2 = st.columns(2)
            sub_c1.metric("총 대조 건수", f"{r['total_cnt']} 건")
            sub_c2.metric(
                "최종 총 차이 금액",
                f"{r['total_diff_sum']:,}원",
                delta=(f"{r['total_diff_sum']:,}원" if r["total_diff_sum"] != 0 else None),
            )

            sub_c3, sub_c4 = st.columns(2)
            sub_c3.metric(cfg["label_2"], f"{r['total_2_sum']:,}원")
            sub_c4.metric(cfg["label_1"], f"{r['total_1_sum']:,}원")

            if r["detail_metrics"]:
                st.write("")
                st.markdown(
                    "<div style='font-size: 15px; font-weight: 700; color: #38bdf8; margin-bottom: 6px;'>🔧 공임 / 부품 세부 내역 (VAT 제외)</div>",
                    unsafe_allow_html=True,
                )
                metrics = r["detail_metrics"]
                for i in range(0, len(metrics), 2):
                    for m_col, (m_label, m_value) in zip(st.columns(2), metrics[i:i + 2]):
                        m_col.metric(m_label, f"{m_value:,}원")

            st.write("")
            # 📌 다운로드 버튼 가로 나란히 배치
            for dl_col, (dl_label, dl_data, dl_name) in zip(st.columns(len(r["downloads"])), r["downloads"]):
                with dl_col:
                    st.download_button(
                        label=dl_label,
                        data=dl_data,
                        file_name=dl_name,
                        mime=XLSX_MIME,
                        use_container_width=True,
                    )
    else:
        with right_col:
            st.info("👈 좌측에서 두 파일을 모두 선택하시면 우측에 상세 대조 내역과 차액 리스트가 표시됩니다.")

elif mode == "공임코드 비교":
    top_col1, top_col2 = st.columns([7.5, 2.5])
    with top_col1:
        st.markdown("### 🔍 VF01 에서 확인된 공임코드들의 중복값을 비교합니다.")
    with top_col2:
        show_c = st.session_state.show_group_c
        if st.button("➖ 그룹 C 제거하기" if show_c else "➕ 그룹 C 추가하기", use_container_width=True):
            st.session_state.show_group_c = not show_c
            st.rerun()

    st.write("")

    # 그룹별 (공임코드 예시, 공임내역 예시)
    LABOR_PLACEHOLDERS = {
        "A": ("26215", "coolant drain"),
        "B": ("26010", "coolant drain"),
        "C": ("26300", "inspection"),
    }
    groups = ["A", "B", "C"] if st.session_state.show_group_c else ["A", "B"]
    three_cols = len(groups) == 3
    label_margin = "margin-top: 8px; margin-bottom: 4px;" if three_cols else "margin-top: 10px; margin-bottom: 6px;"

    group_inputs = {}
    for g_col, g in zip(st.columns(len(groups)), groups):
        ph_code, ph_desc = LABOR_PLACEHOLDERS[g]
        with g_col:
            sub_1, sub_2 = st.columns([1.2, 1.8])
            with sub_1:
                code = st.text_input(f"공임코드 ({g})", placeholder=f"예시: {ph_code}", key=f"code_{g.lower()}")
            with sub_2:
                desc = st.text_input("공임내역", placeholder=f"예시: {ph_desc}", key=f"desc_{g.lower()}")
            st.markdown(
                f"<div style='font-size: 19px; font-weight: 700; {label_margin} color: #f1f5f9;'>공임상세내역</div>",
                unsafe_allow_html=True,
            )
            text = st.text_area(
                f"공임상세내역 ({g})",
                height=350 if three_cols else 380,
                placeholder="예시 : 900-00-B   Engine hood open 1 7",
                label_visibility="collapsed",
                key=f"textarea_{g.lower()}",
            )
            group_inputs[g] = (code.strip(), desc.strip(), text)

    st.write("")
    start_compare = st.button("🔍 비교진행", use_container_width=True, type="primary")

    if start_compare:
        # 공임코드를 입력했으면 상세내역 맨 위에 한 줄로 추가
        combined_texts = {
            g: (f"{code} {desc}\n" + text if code else text)
            for g, (code, desc, text) in group_inputs.items()
        }

        if not combined_texts["A"].strip() and not combined_texts["B"].strip():
            st.warning("⚠️ 공임코드를 입력하거나 내용을 붙여넣어 주세요.")
        else:
            maps = {g: parse_labor_lines(t) for g, t in combined_texts.items()}

            # 입력 순서(A → B → C)대로, 2개 이상 그룹에 있는 코드
            all_codes = dict.fromkeys(code for m in maps.values() for code in m)
            duplicate_codes = [
                code for code in all_codes if sum(code in m for m in maps.values()) >= 2
            ]

            st.divider()

            res_col_left, res_col_right = st.columns([3, 7], gap="large")

            with res_col_left:
                st.markdown("### 📌 비교 요약 결과")
                for g, m in maps.items():
                    unique_cnt = len([c for c in m if not any(c in other for og, other in maps.items() if og != g)])
                    sum_col1, sum_col2 = st.columns(2)
                    with sum_col1:
                        st.metric(f"{g}그룹 총 건수", f"{len(m)} 건")
                    with sum_col2:
                        st.metric(f"{g}그룹 고유", f"{unique_cnt} 건")

                st.write("")
                st.metric(
                    "중복된 공임코드",
                    f"{len(duplicate_codes)} 건",
                    delta=("중복 발견" if duplicate_codes else None),
                )

            with res_col_right:
                st.markdown("### 🚨 중복 발견 내역")
                if duplicate_codes:
                    group_headers = [
                        f"{code} {desc}".strip() or f"{g}그룹 내용"
                        for g, (code, desc, _) in group_inputs.items()
                    ]
                    th_tags = "".join(f"<th>{h}</th>" for h in ["No.", "중복 공임코드"] + group_headers)

                    main_tbody = []
                    for idx, code in enumerate(duplicate_codes, 1):
                        group_cells = "".join(
                            f'<td class="col-amt">{" | ".join(m[code]) if code in m else "-"}</td>'
                            for m in maps.values()
                        )
                        main_tbody.append(
                            f'<tr><td class="col-no">{idx}</td><td class="col-id copyable" onclick="copyCell(this)">{code}</td>{group_cells}</tr>'
                        )

                    css_dup = """
                      * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
                      body { background-color: transparent; color: #f8fafc; }
                      table { border-collapse: collapse; width: 100%; font-size: 15px; user-select: text; }
                      th { background-color: #1e293b; color: #fff; padding: 10px 12px; border: 1px solid #334155; text-align: center; }
                      td { padding: 8px 12px; border: 1px solid #334155; }
                      td.copyable { cursor: pointer; }
                      td.copyable:hover { background-color: rgba(14, 165, 233, 0.2) !important; }
                      #toast { visibility: hidden; position: fixed; top: 10px; left: 50%; transform: translateX(-50%); background-color: #0284c7; color: #fff; padding: 8px 16px; border-radius: 6px; font-weight: bold; z-index: 99999; }
                      #toast.show { visibility: visible; }
                    """

                    js_dup = """
                      function copyCell(el) {
                        const text = el.innerText.trim();
                        if (!text || text === '-') return;
                        navigator.clipboard.writeText(text);
                        const toast = document.getElementById('toast');
                        toast.innerText = '📋 복사 완료: ' + text;
                        toast.className = 'show';
                        setTimeout(function() { toast.className = ''; }, 1200);
                      }
                    """

                    table_html = (
                        '<!DOCTYPE html><html><head><meta charset="utf-8" />'
                        f"<style>{css_dup}</style></head><body>"
                        '<div id="toast">📋 복사 완료!</div>'
                        f"<table><thead><tr>{th_tags}</tr></thead>"
                        f'<tbody>{"".join(main_tbody)}</tbody></table>'
                        f"<script>{js_dup}</script>"
                        "</body></html>"
                    )
                    components.html(
                        table_html, height=min(600, len(duplicate_codes) * 44 + 80), scrolling=True
                    )
                else:
                    st.success("✅ 비교 그룹 간에 중복된 공임코드가 없습니다.")

elif mode == "캘린더":
    st.markdown("### 📅 보증팀 주요 마감 및 업무 일정 관리")
    st.write("")

    def shift_month(delta):
        """delta = -1 : 이전 달, +1 : 다음 달"""
        total = st.session_state.cal_year * 12 + (st.session_state.cal_month - 1) + delta
        st.session_state.cal_year, month_idx = divmod(total, 12)
        st.session_state.cal_month = month_idx + 1
        st.rerun()

    ctrl_c1, ctrl_c2, ctrl_c3, ctrl_c4 = st.columns([1.5, 4, 1.5, 3])
    with ctrl_c1:
        if st.button("◀ 이전 달", use_container_width=True):
            shift_month(-1)

    with ctrl_c2:
        st.markdown(
            f"<h3 style='text-align: center; margin: 0; color: #f1f5f9;'>"
            f"{st.session_state.cal_year}년 {st.session_state.cal_month}월"
            f"</h3>",
            unsafe_allow_html=True,
        )

    with ctrl_c3:
        if st.button("다음 달 ▶", use_container_width=True):
            shift_month(1)

    with ctrl_c4:
        if st.button("📌 오늘 날짜로 이동", use_container_width=True):
            today_now = datetime.date.today()
            st.session_state.cal_year = today_now.year
            st.session_state.cal_month = today_now.month
            st.rerun()

    st.write("")

    # 업무 구분 : (달력 칩 스타일, 목록 아이콘)
    CATEGORY_INFO = {
        "MW 마감": ("chip-mw", "🟦"),
        "쿠폰 정산": ("chip-coupon", "🟩"),
        "본사 청구": ("chip-claim", "🟨"),
        "휴가/당직": ("chip-vacation", "🟥"),
        "기타": ("chip-etc", "⬜"),
    }
    category_styles = {name: info[0] for name, info in CATEGORY_INFO.items()}

    def event_range(ev):
        """일정의 (시작일, 종료일) 문자열. 예전 형식('date' 하나)도 지원"""
        s_d = ev.get("start_date", ev.get("date", ""))
        return s_d, ev.get("end_date", s_d)

    # 📌 일요일 시작 요일로 설정하여 요일 밀림 방지
    calendar.setfirstweekday(calendar.SUNDAY)
    month_cal = calendar.monthcalendar(
        st.session_state.cal_year, st.session_state.cal_month
    )
    today = datetime.date.today()

    events_by_date = defaultdict(list)
    for ev in st.session_state.calendar_events:
        s_date_str, e_date_str = event_range(ev)
        try:
            d_cur = datetime.datetime.strptime(s_date_str, "%Y-%m-%d").date()
            d_end = datetime.datetime.strptime(e_date_str, "%Y-%m-%d").date()
            if d_cur > d_end:
                d_cur, d_end = d_end, d_cur
            while d_cur <= d_end:
                events_by_date[d_cur.strftime("%Y-%m-%d")].append(ev)
                d_cur += datetime.timedelta(days=1)
        except Exception:
            events_by_date[s_date_str].append(ev)

    cal_inner_css = """
      * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
      body { background-color: transparent; color: #f8fafc; overflow: hidden; }
      .cal-container {
        width: 100%;
        border: 1px solid #334155;
        border-radius: 8px;
        overflow: hidden;
        background-color: #0b0f19;
      }
      .cal-header-row {
        display: grid;
        grid-template-columns: repeat(7, 1fr);
        background-color: #1e293b;
        border-bottom: 1px solid #334155;
        text-align: center;
        font-weight: bold;
        padding: 9px 0;
        color: #cbd5e1;
        font-size: 15px;
      }
      .cal-grid {
        display: grid;
        grid-template-columns: repeat(7, 1fr);
      }
      .cal-cell {
        min-height: 85px;
        max-height: 95px;
        border-right: 1px solid #334155;
        border-bottom: 1px solid #334155;
        padding: 6px;
        background-color: #0b0f19;
        overflow-y: auto;
      }
      .cal-cell::-webkit-scrollbar { width: 4px; }
      .cal-cell::-webkit-scrollbar-thumb { background: #334155; border-radius: 2px; }
      .cal-cell:nth-child(7n) {
        border-right: none;
      }
      .cal-cell.other-month {
        background-color: #07090e;
        color: #475569;
      }
      .cal-day-num {
        font-size: 14px;
        font-weight: bold;
        margin-bottom: 4px;
        display: inline-block;
      }
      .cal-today {
        background-color: #0284c7;
        color: #ffffff;
        border-radius: 50%;
        width: 22px;
        height: 22px;
        text-align: center;
        line-height: 22px;
      }
      .cal-sun { color: #f87171; }
      .cal-sat { color: #60a5fa; }
      .event-chip {
        font-size: 11px;
        padding: 2px 5px;
        border-radius: 4px;
        margin-bottom: 3px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        font-weight: 600;
        display: block;
      }
      .chip-mw { background-color: rgba(14, 165, 233, 0.25); color: #38bdf8; border: 1px solid rgba(14, 165, 233, 0.4); }
      .chip-coupon { background-color: rgba(34, 197, 94, 0.25); color: #4ade80; border: 1px solid rgba(34, 197, 94, 0.4); }
      .chip-claim { background-color: rgba(234, 179, 8, 0.25); color: #facc15; border: 1px solid rgba(234, 179, 8, 0.4); }
      .chip-vacation { background-color: rgba(244, 63, 94, 0.25); color: #fb7185; border: 1px solid rgba(244, 63, 94, 0.4); }
      .chip-etc { background-color: rgba(148, 163, 184, 0.25); color: #cbd5e1; border: 1px solid rgba(148, 163, 184, 0.4); }
    """

    cal_html = [
        '<!DOCTYPE html><html><head><meta charset="utf-8" />',
        f"<style>{cal_inner_css}</style></head><body>",
        '<div class="cal-container">',
        '<div class="cal-header-row">',
    ]

    days_labels = [
        '<span class="cal-sun">일</span>',
        "월",
        "화",
        "수",
        "목",
        "금",
        '<span class="cal-sat">토</span>',
    ]
    for dl in days_labels:
        cal_html.append(f"<div>{dl}</div>")
    cal_html.append("</div>")

    cal_html.append('<div class="cal-grid">')
    for week in month_cal:
        for day_idx, day_num in enumerate(week):
            if day_num == 0:
                cal_html.append('<div class="cal-cell other-month"></div>')
            else:
                current_d = datetime.date(
                    st.session_state.cal_year, st.session_state.cal_month, day_num
                )
                date_str = current_d.strftime("%Y-%m-%d")

                if current_d == today:
                    num_html = f'<div class="cal-day-num cal-today">{day_num}</div>'
                elif day_idx == 0:
                    num_html = f'<div class="cal-day-num cal-sun">{day_num}</div>'
                elif day_idx == 6:
                    num_html = f'<div class="cal-day-num cal-sat">{day_num}</div>'
                else:
                    num_html = f'<div class="cal-day-num">{day_num}</div>'

                chips_html = []
                for ev in events_by_date.get(date_str, []):
                    cat_cls = category_styles.get(ev.get("category", "기타"), "chip-etc")
                    chips_html.append(
                        f'<div class="event-chip {cat_cls}" title="{ev.get("memo", "")}">'
                        f'[{ev.get("category")}] {ev.get("title")}'
                        "</div>"
                    )

                cal_html.append(f'<div class="cal-cell">{num_html}{"".join(chips_html)}</div>')
    cal_html.append("</div></div></body></html>")

    num_weeks = len(month_cal)
    calc_iframe_height = num_weeks * 95 + 50
    components.html("".join(cal_html), height=calc_iframe_height, scrolling=False)

    st.divider()

    bottom_col_left, bottom_col_right = st.columns([3.2, 6.8], gap="large")

    with bottom_col_left:
        st.markdown("#### ✏️ 일정 등록 / 관리")
        with st.expander("➕ 새 일정 등록하기", expanded=True):
            default_day = datetime.date.today()

            head_date_col, chk_col = st.columns([6.5, 3.5])
            with head_date_col:
                st.markdown(
                    "<div style='font-size: 19px; font-weight: 700; color: #f1f5f9; margin-top: 10px;'>날짜 선택</div>",
                    unsafe_allow_html=True,
                )
            with chk_col:
                is_single_day = st.checkbox("하루예약", value=False, key="cal_single_day")

            if is_single_day:
                start_date_val = st.date_input("선택 일자", value=default_day, key="cal_start_date_single")
                end_date_val = start_date_val
            else:
                d_col1, d_col2 = st.columns(2)
                with d_col1:
                    start_date_val = st.date_input("시작일", value=default_day, key="cal_start_date_range")
                with d_col2:
                    end_date_val = st.date_input(
                        "종료일",
                        value=start_date_val,
                        min_value=start_date_val,
                        key="cal_end_date_range",
                    )

            new_title = st.text_input("일정 제목", placeholder="예: 여름 휴가 / 9월 쿠폰 정산", key="cal_add_title")
            new_cat = st.selectbox(
                "업무 구분",
                list(CATEGORY_INFO),
                key="cal_add_cat",
            )
            new_memo = st.text_area("상세 메모 (선택)", height=70, placeholder="특이사항 입력", key="cal_add_memo")

            if st.button("등록하기", type="primary", use_container_width=True):
                if new_title.strip():
                    new_event = {
                        "start_date": start_date_val.strftime("%Y-%m-%d"),
                        "end_date": end_date_val.strftime("%Y-%m-%d"),
                        "title": new_title.strip(),
                        "category": new_cat,
                        "memo": new_memo.strip(),
                    }

                    def _add_event(data):
                        data.append(new_event)
                        return True

                    save_shared_and_rerun(
                        "calendar_data.json", DEFAULT_CALENDAR_DATA, _add_event,
                        "✅ 일정이 영구 저장되었습니다!",
                    )
                else:
                    st.warning("⚠️ 일정 제목을 입력해 주세요.")


    with bottom_col_right:
        cur_year = st.session_state.cal_year
        cur_month = st.session_state.cal_month
        month_start_str = f"{cur_year:04d}-{cur_month:02d}-01"
        last_day_num = calendar.monthrange(cur_year, cur_month)[1]
        month_end_str = f"{cur_year:04d}-{cur_month:02d}-{last_day_num:02d}"

        # 이번 달에 걸쳐 있는 일정만, 시작일 순으로
        month_events = [
            ev for ev in st.session_state.calendar_events
            if event_range(ev)[0] <= month_end_str and event_range(ev)[1] >= month_start_str
        ]
        month_events.sort(key=lambda ev: event_range(ev)[0])

        st.markdown(
            f"#### 📋 {st.session_state.cal_month}월 등록된 일정 목록"
            f" <span style='font-size: 15px; color: #38bdf8; font-weight: normal; margin-left: 8px;'>({len(month_events)}건)</span>",
            unsafe_allow_html=True,
        )

        with st.container(height=360):
            if month_events:
                grid_cols = st.columns(2, gap="medium")
                for i, ev in enumerate(month_events):
                    with grid_cols[i % 2]:
                        c_card, c_del = st.columns([8.2, 1.8])
                        with c_card:
                            category_icon = CATEGORY_INFO.get(ev.get("category"), (None, "📌"))[1]
                            s_d, e_d = event_range(ev)
                            date_display = s_d if s_d == e_d else f"{s_d} ~ {e_d}"
                            memo_txt = ev.get("memo", "").strip() or "-"

                            st.markdown(
                                f"""
                                <div class="event-card">
                                  <div class="event-card-title">
                                    {category_icon} <span>{date_display}</span> <span style="color:#94a3b8;">|</span> <span>[{ev['category']}] {ev['title']}</span>
                                  </div>
                                  <div class="event-card-memo">{memo_txt}</div>
                                </div>
                                """,
                                unsafe_allow_html=True,
                            )
                        with c_del:
                            if st.button("삭제", key=f"delev_{i}", use_container_width=True):
                                # 순번이 아니라 내용으로 찾아서 삭제 (그 사이 목록이 바뀌어도 안전)
                                def _del_event(data, target=dict(ev)):
                                    if target in data:
                                        data.remove(target)
                                        return True
                                    return False

                                save_shared_and_rerun(
                                    "calendar_data.json", DEFAULT_CALENDAR_DATA, _del_event,
                                    "🗑️ 일정이 삭제되었습니다.",
                                )
            else:
                st.info("해당 월에 등록된 일정이 없습니다.")
