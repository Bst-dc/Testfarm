import streamlit as st
import sqlite3 as _sqlite3
import db_adapter
from db_schema import SQLITE_DDL, PG_DDL, DB_TABLES
import pandas as pd
import os
import io
import glob
import shutil
import tempfile
import html
import re
import hashlib
import base64
from datetime import datetime

# ========== 로고/아이콘 ==========
_ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")

def _load_data_uri(filename, mime):
    with open(os.path.join(_ASSETS_DIR, filename), "rb") as f:
        return f"data:{mime};base64," + base64.b64encode(f.read()).decode("ascii")

MEDAL_ICON_DATA_URI = _load_data_uri("medal_icon.gif", "image/gif")

def clean_excel_text(val):
    """엑셀에서 읽은 값을 문자열로 정리한다.
    이표번호처럼 큰 숫자가 들어가는 열은, 같은 열의 다른 행에 빈 칸이 하나만 있어도
    pandas가 열 전체를 float로 읽어 "410002123456.0"처럼 끝에 .0이 붙어버린다.
    정수값인 float는 .0을 떼고 정수 문자열로 되돌려 이 문제를 막는다."""
    if val is None:
        return None
    if pd.isna(val):
        return None
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    text = str(val).strip()
    if not text or text.lower() == "nan" or text == "None":
        return None
    return text

def format_thousands_input(key):
    """숫자를 입력받는 text_input의 on_change 콜백.
    입력한 값에서 콤마를 뗀 뒤 다시 천단위 콤마를 붙여 session_state에 되돌려 놓는다."""
    raw = st.session_state.get(key, "")
    cleaned = raw.replace(",", "").strip()
    if not cleaned:
        return
    try:
        num = float(cleaned)
    except ValueError:
        return
    st.session_state[key] = f"{int(num):,}" if num == int(num) else f"{num:,.2f}"

def parse_thousands_input(key):
    """format_thousands_input 으로 콤마가 붙은 텍스트를 다시 숫자로 되돌린다."""
    cleaned = st.session_state.get(key, "").replace(",", "").strip()
    try:
        return float(cleaned) if cleaned else 0.0
    except ValueError:
        return 0.0

# ========== 데이터 저장 위치 ==========
# 로컬 PC와 클라우드(영구 볼륨)에서 같은 코드가 돌아가도록 경로를 환경변수로 분리한다.
#   - 로컬 윈도우 : 기본값 (사용자 폴더)/시험농장DB
#   - 클라우드    : 환경변수 ERP_DB_DIR 에 볼륨 경로(/data) 지정
# 구글 드라이브 폴더에는 두지 않는다. 동기화 프로그램이 파일을 잠가 잠금 오류와 손상을 일으킨다.
#
# DATABASE_URL(secrets.toml 또는 환경변수)이 있으면 SQLite 파일 대신 Supabase(PostgreSQL)를 쓴다.
# 이때 농장 DB 파일 경로(erp_sunsan.db)는 스키마 이름(sunsan)을 정하는 용도로만 쓰이고 실제 파일은 없다.
# 그래서 DB가 있는지는 os.path.exists 가 아니라 db_exists() 로 확인해야 한다.
USE_PG = db_adapter.enabled()
sqlite3 = db_adapter if USE_PG else _sqlite3

DB_DIR =os.environ.get("ERP_DB_DIR") or os.path.join(os.path.expanduser("~"), "시험농장DB")
BACKUP_DIR = os.path.join(DB_DIR, "backup")
os.makedirs(BACKUP_DIR, exist_ok=True)
BACKUP_KEEP = 30  # 농장별 보관 백업 개수

# 사이드바의 '지금 백업 만들기'로 만든 백업을 모아 두는 폴더 (기본: 바탕화면\AI\시험농장백업파일).
# 환경변수 ERP_MANUAL_BACKUP_DIR 로 바꿀 수 있고, 만들 수 없는 경로면 기본 백업 폴더를 쓴다.
MANUAL_BACKUP_DIR = os.environ.get("ERP_MANUAL_BACKUP_DIR") or os.path.join(
    os.path.expanduser("~"), "OneDrive", "Desktop", "AI", "시험농장백업파일"
)
try:
    os.makedirs(MANUAL_BACKUP_DIR, exist_ok=True)
except OSError:
    MANUAL_BACKUP_DIR = BACKUP_DIR

st.set_page_config(page_title="대구축협 시험농장 관리 시스템", layout="wide", page_icon="🐂")

APP_CSS = """
<style>
@import url("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css");

:root {
    --farm-primary: #2E6B57;
    --farm-primary-dark: #23533F;
    --farm-accent: #2F5D8A;
    --farm-bg: #F3F6F4;
    --farm-card: #FFFFFF;
    --farm-ink: #1F2A24;
    --farm-muted: #5F6F66;
    --farm-line: #D5E0D9;
    --farm-soft: #EAF2ED;
    --farm-shadow: 0 1px 2px rgba(20, 40, 30, 0.04), 0 6px 18px rgba(20, 40, 30, 0.06);
}

html, body, [class*="st-"] {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, 'Segoe UI',
        'Apple SD Gothic Neo', 'Noto Sans KR', 'Malgun Gothic', sans-serif !important;
}
/* 위 폰트 강제 적용이 Streamlit 내장 아이콘(사이드바 접기 화살표, expander 화살표 등)의
   전용 아이콘 폰트까지 덮어써서 "keyboard_double_arrow_left" 같은 글자가 그대로 보이는
   문제를 막기 위해 아이콘 요소는 원래 아이콘 폰트로 되돌린다. */
[data-testid="stIconMaterial"] {
    font-family: "Material Symbols Rounded" !important;
}

/* ---------- 전체 배경 / 여백 ---------- */
.stApp {
    background:
        radial-gradient(900px 420px at 0% 0%, rgba(46, 107, 87, 0.08) 0%, rgba(46, 107, 87, 0) 70%),
        radial-gradient(900px 420px at 100% 0%, rgba(47, 93, 138, 0.07) 0%, rgba(47, 93, 138, 0) 70%),
        var(--farm-bg);
    color: var(--farm-ink);
}
.block-container, [data-testid="stMainBlockContainer"] {
    padding-top: 4.6rem !important;
    padding-bottom: 4rem !important;
    padding-left: 2.5rem !important;
    padding-right: 2.5rem !important;
    max-width: 1680px;
}
h1, h2, h3, h4, h5 { color: var(--farm-ink); letter-spacing: -0.01em; }
h2, h3 { font-weight: 800 !important; }
hr { border-color: var(--farm-line) !important; opacity: 0.8; }

/* ---------- 사이드바 ---------- */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #FFFFFF 0%, #F1F6F3 100%) !important;
    border-right: 1px solid var(--farm-line);
}
section[data-testid="stSidebar"] * { color: var(--farm-ink); }

/* 사이드바 라디오 버튼(농장 선택)을 카드형 메뉴로
   (React Aria 기반 라디오: label[data-testid="stRadioOption"], 선택 상태는 data-selected="true") */
section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"] {
    padding: 14px 18px !important;
    background-color: #FFFFFF !important;
    border: 1px solid var(--farm-line) !important;
    border-radius: 12px !important;
    margin-bottom: 10px !important;
    cursor: pointer !important;
    transition: all 0.15s ease !important;
    box-shadow: 0 1px 2px rgba(20, 40, 30, 0.04) !important;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"]:hover {
    border-color: var(--farm-primary) !important;
    background-color: var(--farm-soft) !important;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"][data-selected="true"] {
    background: linear-gradient(135deg, var(--farm-primary) 0%, var(--farm-accent) 100%) !important;
    border-color: transparent !important;
    box-shadow: 0 6px 14px rgba(46, 107, 87, 0.28) !important;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"][data-selected="true"] p {
    color: #FFFFFF !important;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"] p {
    font-size: 19px !important;
    font-weight: 800 !important;
    margin: 0 !important;
}
/* 라디오 동그라미 숨기기 (텍스트만 돋보이게) */
section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"] > div > div:first-child {
    display: none !important;
}
section[data-testid="stSidebar"] div[data-testid="stWidgetLabel"] p {
    font-size: 15px !important;
    font-weight: 800 !important;
    color: var(--farm-muted) !important;
    padding-bottom: 6px !important;
}

/* ---------- 핵심 지표 카드 ---------- */
div[data-testid="stMetric"] {
    background: var(--farm-card);
    border: 1px solid var(--farm-line);
    border-top: 4px solid var(--farm-primary);
    border-radius: 14px;
    padding: 16px 20px !important;
    box-shadow: var(--farm-shadow);
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}
div[data-testid="stMetric"]:hover {
    transform: translateY(-2px);
    box-shadow: 0 10px 24px rgba(20, 40, 30, 0.10);
}
div[data-testid="stMetric"] label, div[data-testid="stMetricLabel"] p {
    color: var(--farm-muted) !important;
    font-weight: 700 !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: clip !important;
}
div[data-testid="stMetricValue"] {
    font-size: 1.75rem !important;
    font-weight: 800 !important;
    color: var(--farm-ink) !important;
    font-variant-numeric: tabular-nums;
}
div[data-testid="stMetricValue"] > div {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: clip !important;
}

/* ---------- 탭 (Pill 형) ---------- */
div[data-testid="stTabs"] div[role="tablist"] {
    gap: 8px;
    border-bottom: none;
    padding-bottom: 8px;
    flex-wrap: nowrap !important;
    overflow-x: auto !important;
    overflow-y: hidden;
    scrollbar-width: thin;
}
div[data-testid="stTab"] {
    background-color: #FFFFFF !important;
    border: 1px solid var(--farm-line) !important;
    border-radius: 999px !important;
    padding: 9px 20px !important;
    box-shadow: 0 1px 2px rgba(20, 40, 30, 0.04);
    transition: all 0.15s ease;
}
div[data-testid="stTab"]:hover {
    border-color: var(--farm-primary) !important;
    background-color: var(--farm-soft) !important;
}
div[data-testid="stTab"][aria-selected="true"] {
    background: linear-gradient(135deg, var(--farm-primary) 0%, var(--farm-accent) 100%) !important;
    border-color: transparent !important;
    box-shadow: 0 4px 12px rgba(46, 107, 87, 0.28);
}
div[data-testid="stTab"][aria-selected="true"] * { color: #FFFFFF !important; }
div[data-testid="stTab"] [data-testid="stMarkdownContainer"] p {
    font-size: 16px !important;
    font-weight: 700 !important;
    margin: 0;
}
/* 탭 활성화 시 기본 하단 표시선 제거 (실제 DOM: stTab 안의 .react-aria-SelectionIndicator) */
div[data-testid="stTab"] .react-aria-SelectionIndicator { display: none !important; }

/* ---------- 버튼: 주요 액션(primary) / 보조(secondary) 위계 ---------- */
button[data-testid^="stBaseButton-primary"] {
    background: linear-gradient(135deg, var(--farm-primary) 0%, var(--farm-accent) 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    font-weight: 700 !important;
    box-shadow: 0 4px 12px rgba(46, 107, 87, 0.25) !important;
    transition: transform 0.12s ease, box-shadow 0.12s ease, filter 0.12s ease;
}
button[data-testid^="stBaseButton-primary"]:hover {
    transform: translateY(-1px);
    filter: brightness(1.06);
    box-shadow: 0 8px 18px rgba(46, 107, 87, 0.30) !important;
}
button[data-testid^="stBaseButton-primary"] p { color: #FFFFFF !important; }
button[data-testid^="stBaseButton-secondary"] {
    background: #FFFFFF !important;
    color: var(--farm-ink) !important;
    border: 1px solid #C5D3CB !important;
    font-weight: 600 !important;
    transition: all 0.12s ease;
}
button[data-testid^="stBaseButton-secondary"]:hover {
    border-color: var(--farm-primary) !important;
    background: var(--farm-soft) !important;
    color: var(--farm-primary-dark) !important;
}

/* ---------- 입력창 ---------- */
div[data-testid="stSelectbox"] div[role="group"],
div[data-testid="stMultiSelect"] div[role="group"],
.stTextInput input, .stNumberInput input, .stDateInput input, .stTextArea textarea {
    background-color: #FFFFFF !important;
    border-radius: 10px !important;
}

/* ---------- 카드형 폼 / expander ---------- */
div[data-testid="stForm"] {
    background: var(--farm-card);
    border: 1px solid var(--farm-line) !important;
    border-radius: 16px !important;
    box-shadow: var(--farm-shadow);
    padding: 20px 22px !important;
}
div[data-testid="stExpander"] details {
    background: var(--farm-card);
    border: 1px solid var(--farm-line) !important;
    border-radius: 14px !important;
    box-shadow: 0 1px 2px rgba(20, 40, 30, 0.04);
}
div[data-testid="stExpander"] summary p { font-weight: 700; }

/* ---------- 표 ---------- */
div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
    border: 1px solid var(--farm-line);
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 1px 3px rgba(20, 40, 30, 0.04);
    background: #FFFFFF;
}

/* ---------- 알림 ---------- */
div[data-testid="stAlert"] > div { border-radius: 12px !important; }

/* ================= 레이아웃 공통 값 (C) ================= */
:root {
    --farm-radius: 12px;      /* 카드·폼·표·expander 모서리 */
    --farm-btn-h: 2.6rem;     /* 버튼 높이 */
}
div[data-testid="stMetric"], div[data-testid="stForm"], div[data-testid="stExpander"] details,
div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
    border-radius: var(--farm-radius) !important;
}
div[data-testid="stMetric"], div[data-testid="stForm"] { box-shadow: var(--farm-shadow) !important; }
.stButton > button, .stFormSubmitButton > button, .stDownloadButton > button {
    min-height: var(--farm-btn-h);
    border-radius: 10px !important;
}
.farm-title { font-size: 1.9rem; line-height: 1.25; word-break: keep-all; }
.farm-title .title-short { display: none; }

/* ---------- 1. 가로 스크롤 방지 ---------- */
html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] { overflow-x: hidden !important; }
[data-testid="stMainBlockContainer"] {
    max-width: min(1680px, 100%) !important;
    container-type: inline-size;      /* 아래 @container 규칙의 기준 = 본문 폭 (사이드바 제외) */
    container-name: farm-main;
}
.block-container * { scroll-margin-top: 5rem; }

/* ---------- 2. KPI 카드: 숫자는 폭에 맞게 크기 자동 축소, 라벨 줄바꿈 허용 ---------- */
div[data-testid="stMetricValue"] { font-size: clamp(1.05rem, 2.3cqi, 1.75rem) !important; }
div[data-testid="stMetricLabel"] p { font-size: clamp(0.8rem, 1.2cqi, 0.95rem) !important; line-height: 1.3; }
div[data-testid="stMetric"] { padding: 14px 16px !important; min-width: 0; }

/* ---------- 4. 탭 버튼 줄바꿈 금지 ---------- */
div[data-testid="stTab"] { flex: 0 0 auto !important; white-space: nowrap; }

/* ---------- 11. 위험한 버튼 (key 가 danger_ 로 시작) ---------- */
[class*="st-key-danger_"] button {
    background: #FFF5F5 !important;
    border: 1px solid #E8A9A4 !important;
    color: #B42318 !important;
}
[class*="st-key-danger_"] button p { color: #B42318 !important; font-weight: 700 !important; }
[class*="st-key-danger_"] button:hover:not(:disabled) { background: #FDE7E5 !important; border-color: #B42318 !important; }
[class*="st-key-danger_"] button:disabled { opacity: 0.55; }

/* ---------- 17. 준비중 화면 ---------- */
.farm-coming { background: var(--farm-card); border: 1px solid var(--farm-line); border-radius: var(--farm-radius);
               box-shadow: var(--farm-shadow); padding: 22px 24px; }
.farm-coming h3 { margin: 0 0 6px 0; }
.farm-coming p { color: var(--farm-muted); margin: 0 0 16px 0; }
.farm-badge { display: inline-block; font-size: 0.8rem; font-weight: 800; color: #9A5B00; background: #FFF3D6;
              border: 1px solid #F2D49B; border-radius: 999px; padding: 2px 10px; vertical-align: middle; margin-left: 6px; }
.farm-coming-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; }
.farm-coming-card { border: 1px dashed var(--farm-line); border-radius: 10px; padding: 14px 16px; background: var(--farm-soft); }
.farm-coming-card b { display: block; font-size: 1.05rem; margin-bottom: 4px; }
.farm-coming-card span { color: var(--farm-muted); font-size: 0.88rem; }

/* ---------- C. 사이드바 간격 줄이기 ---------- */
section[data-testid="stSidebar"] hr { margin: 0.5rem 0 !important; }
section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"] {
    padding: 12px 16px !important;
    margin-bottom: 8px !important;
}
section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] { padding-top: 0.5rem; padding-bottom: 1rem; }

/* ================= 반응형: 본문 폭 기준 ================= */
/* 본문 폭 900px 미만 (예: 창 800px + 사이드바 열림 → 본문 약 450px) */
@container farm-main (max-width: 900px) {
    /* 5. 2열 화면은 위아래 1열로 */
    div[data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 100% !important; width: 100% !important; min-width: 100% !important;
    }
    /* 6. 열 안의 작은 칸(폼 입력칸, 필터)은 한 줄 최대 2칸 */
    div[data-testid="stColumn"] div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"],
    [class*="st-key-grid2_"] div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 calc(50% - 0.5rem) !important; width: calc(50% - 0.5rem) !important; min-width: calc(50% - 0.5rem) !important;
    }
    /* 2. KPI: 600~900px 은 3개씩 */
    [class*="st-key-kpi_"] div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 calc(33.333% - 0.67rem) !important; width: calc(33.333% - 0.67rem) !important;
        min-width: calc(33.333% - 0.67rem) !important;
    }
    /* 3. 본문 제목: 로고는 사이드바에만, 제목은 짧게 */
    .farm-title-logo { display: none !important; }
    .farm-title { font-size: 1.35rem; }
    .farm-title .title-long { display: none; }
    .farm-title .title-short { display: inline; }
    .farm-coming-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
}
/* 본문 폭 600px 미만: KPI 2개씩 */
@container farm-main (max-width: 600px) {
    [class*="st-key-kpi_"] div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 calc(50% - 0.5rem) !important; width: calc(50% - 0.5rem) !important; min-width: calc(50% - 0.5rem) !important;
    }
}
/* 창 폭 기준: 좁은 창에서는 본문 좌우 여백을 줄인다 */
@media (max-width: 900px) {
    .block-container, [data-testid="stMainBlockContainer"] {
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
}
@media (max-width: 600px) {
    .block-container, [data-testid="stMainBlockContainer"] { padding-top: 4rem !important; }
}
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)


def farm_dataframe(data=None, **kwargs):
    """st.dataframe 과 같다. 빈 값(None)을 'None' 대신 '-' 로 보여 준다."""
    kwargs.setdefault("placeholder", "-")
    return st.dataframe(data, **kwargs)


def farm_data_editor(data, **kwargs):
    """st.data_editor 와 같다. 빈 값(None)을 'None' 대신 '-' 로 보여 준다 (입력표는 placeholder="" 로 비워 둔다)."""
    kwargs.setdefault("placeholder", "-")
    return st.data_editor(data, **kwargs)


def notify(message, icon="✅"):
    """st.rerun() 은 화면을 곧바로 다시 그려서 직전의 st.success 가 사용자에게 보이지 않는다.
    다음 실행에서 토스트로 띄우도록 세션에 쌓아 둔다."""
    st.session_state.setdefault("_pending_toasts", []).append((message, icon))


def show_pending_toasts():
    for message, icon in st.session_state.pop("_pending_toasts", []):
        st.toast(message, icon=icon)


def settlement_month_input(conn, tables, label, key, view_keys):
    """등록 폼 위의 '정산연월' 입력칸. 폼 밖에 두어야 값을 바꾸는 즉시 화면이 다시 그려지고,
    그때 아래 등록 내역 표들의 조회 연월(view_keys)과 월말 정산 대상 연월도 같은 달로 맞춘다.
    처음에는 내역이 있는 가장 최근 달(없으면 이번 달)로 시작해 아래 표와 맞춘다."""
    if key not in st.session_state:
        latest = max(
            (m for m in (conn.execute(f"SELECT MAX(settlement_month) FROM {t}").fetchone()[0] for t in tables) if m),
            default=None,
        )
        st.session_state[key] = latest or datetime.now().strftime('%Y-%m')

    def _sync():
        month = str(st.session_state[key]).strip()
        if re.fullmatch(r"\d{4}-\d{2}", month):
            for view_key in view_keys:
                st.session_state[view_key] = month
            st.session_state["calc_target_month"] = month

    return st.text_input(label, key=key, on_change=_sync, help="형식: YYYY-MM")


def month_view_select(conn, table, key, default_month):
    """등록 내역 표 위에 '조회 연월' 선택 상자를 그리고 고른 연월을 돌려준다.
    선택지는 표에 실제로 있는 연월 + 기본 연월(최근 등록한 달)이다.
    처음 열면 내역이 있는 가장 최근 달을 보여 준다(없으면 기본 연월)."""
    months = [r[0] for r in conn.execute(
        f"SELECT DISTINCT settlement_month FROM {table} WHERE settlement_month IS NOT NULL"
    ).fetchall()]
    if key not in st.session_state:
        st.session_state[key] = max(months) if months else default_month
    months = sorted(set(months) | {default_month, st.session_state[key]}, reverse=True)
    left, _ = st.columns([1, 3])  # 짧은 선택값이 칸 전체로 늘어나지 않게
    return left.selectbox("조회 연월", months, key=key)


def item_entry_table(key, codes, names, columns, info=None, name_label="품목"):
    """여러 품목/항목을 한 번에 등록하는 입력표. 전체 품목을 줄마다 미리 나열해 두고,
    이번에 들어온(쓴) 품목 옆에만 숫자를 넣게 한다. 숫자를 넣지 않은 줄은 무시된다.
    (편집표의 글자 칸에 바로 한글을 치면 첫 글자가 영문으로 들어가는 문제가 있어 숫자만 입력받는다.)
    columns 는 입력받을 숫자 열 {열 이름: column_config},
    info 는 참고용으로만 보여 줄 열 {열 이름: (값 목록, column_config)}.
    반환값은 숫자가 하나라도 들어간 줄만 (코드 열 포함). 등록이 끝나면 reset_item_entry(key) 로 비운다."""
    info = info or {}
    ver = st.session_state.setdefault(f"_{key}_ver", 0)
    df = pd.DataFrame({"코드": list(codes), "항목": list(names)})
    for col, (values, _) in info.items():
        df[col] = list(values)
    for col in columns:
        df[col] = pd.Series([None] * len(df), dtype="float")
    # 품목 구성이 바뀌면(품목 추가·삭제) 줄이 달라지므로 편집표 키도 바꾼다.
    sig = hashlib.md5("|".join(map(str, codes)).encode("utf-8")).hexdigest()[:10]
    edited = farm_data_editor(
        df, width="stretch", hide_index=True, num_rows="fixed", placeholder="",
        disabled=["코드", "항목", *info], column_order=["항목", *info, *columns],
        key=f"{key}_editor_{ver}_{sig}",
        column_config={
            "항목": st.column_config.TextColumn(name_label),
            **{col: cfg for col, (_, cfg) in info.items()},
            **columns,
        },
    )
    filled = edited[list(columns)].apply(pd.to_numeric, errors="coerce").fillna(0).gt(0).any(axis=1)
    return edited[filled]


def reset_item_entry(key):
    st.session_state[f"_{key}_ver"] = st.session_state.get(f"_{key}_ver", 0) + 1


def normalize_item_name(name):
    """품목명 중복 비교용: 띄어쓰기와 대소문자를 무시한다 ('팔공 1' == '팔공1')."""
    return re.sub(r"\s+", "", str(name or "")).lower()


@st.dialog("⚠️ 중복된 품목이 있습니다")
def duplicate_item_dialog(message, dup_df):
    st.warning(message)
    farm_dataframe(dup_df, width="stretch", hide_index=True)
    st.caption("품목명은 띄어쓰기·대소문자를 무시하고 비교합니다. 다른 품목이면 이름을 구분되게 바꿔 주세요.")
    if st.button("확인", type="primary", width="stretch"):
        st.rerun()


def _same_value(a, b):
    if pd.isna(a) and pd.isna(b):
        return True
    if pd.isna(a) or pd.isna(b):
        return False
    try:
        return abs(float(a) - float(b)) < 1e-9
    except (TypeError, ValueError):
        return str(a).strip() == str(b).strip()


def editor_changes(original, edited, id_col, cols):
    """편집표(st.data_editor)에서 사용자가 실제로 바꾼 것만 골라낸다.

    예전에는 저장할 때 표의 모든 줄을 화면 값으로 덮어쓰고, 표에서 빠진 줄은 모두 삭제했다.
    그래서 화면이 오래된 상태(다른 곳에서 바꾼 직후)거나 다른 농장의 편집 상태가 섞이면
    고치지 않은 줄이 되돌아가거나 통째로 지워졌다. 이제는
      - changed: 기존 줄 중 cols 값이 원래와 달라진 줄 (삭제 체크한 줄 제외)
      - deleted: '삭제'에 체크한 기존 줄
      - added  : 표 아래에 새로 추가한 줄 (num_rows="add" 인 표만)
    만 돌려준다."""
    existing = edited[edited.index.isin(original.index)]
    added = edited[~edited.index.isin(original.index)]
    delete_mask = existing["삭제"].fillna(False).astype(bool)
    deleted = existing[delete_mask & existing[id_col].notna()]
    kept = existing[~delete_mask]
    changed_idx = [
        idx for idx, row in kept.iterrows()
        if any(not _same_value(row[c], original.at[idx, c]) for c in cols)
    ]
    return kept.loc[changed_idx], deleted, added


@st.dialog("⚠️ 삭제 확인")
def confirm_delete_dialog(message, preview_df, on_confirm):
    """삭제가 들어간 저장은 지워질 줄을 보여 주고 한 번 더 확인받은 뒤 실행한다."""
    st.warning(message)
    farm_dataframe(preview_df, width="stretch", hide_index=True)
    col_cancel, col_ok = st.columns(2)
    if col_cancel.button("취소", width="stretch"):
        st.rerun()
    if col_ok.button("삭제하고 저장", type="primary", width="stretch"):
        on_confirm()
        st.rerun()


def save_with_delete_confirm(label, changed, deleted, added, preview_cols, amount_col, on_save):
    """바뀐 것이 없으면 알리고, 삭제가 있으면 확인 팝업을, 없으면 바로 저장한다."""
    if changed.empty and deleted.empty and added.empty:
        st.info("바뀐 내용이 없습니다.")
        return
    if deleted.empty:
        on_save()
        st.rerun()
    msg = f"{label} {len(deleted)}건을 삭제합니다."
    if amount_col:
        msg += f" (금액 합계 {pd.to_numeric(deleted[amount_col], errors='coerce').fillna(0).sum():,.0f}원)"
    if not changed.empty or not added.empty:
        msg += f" 수정 {len(changed)}건, 추가 {len(added)}건도 함께 저장됩니다."
    confirm_delete_dialog(msg, deleted[preview_cols], on_save)


def show_table_total(count, amount_label, amount):
    """편집표(st.data_editor) 바로 아래에 합계 줄을 붙인다. 표 안에 합계 행을 넣으면
    저장 시 실제 데이터로 들어가 버리므로 표 밖에 따로 그린다."""
    st.markdown(
        f"<div style='display:flex;justify-content:space-between;padding:10px 16px;"
        f"margin:-4px 0 12px;background:#EAF2ED;border:1px solid #D5E0D9;border-radius:10px;"
        f"font-weight:800'><span>합계 ({count:,}건)</span>"
        f"<span>{amount_label} {amount:,.0f} 원</span></div>",
        unsafe_allow_html=True,
    )
import json

# ========== 농장 설정 ==========
FARMS_JSON_PATH = os.path.join(DB_DIR, "farms.json")

def load_farms():
    data = None
    if os.path.exists(FARMS_JSON_PATH):
        try:
            with open(FARMS_JSON_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            pass
            
    if not data:
        data = {
            "구미선산농장": {
                "db_filename": "erp_sunsan.db",
                "color": "#4F46E5",
            },
            "구미고아농장": {
                "db_filename": "erp_goa.db",
                "color": "#059669",
            }
        }
        with open(FARMS_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            
    for k, v in data.items():
        v["db_file"] = os.path.join(DB_DIR, v.get("db_filename", f"erp_{k}.db"))
    return data

def save_farms(farms_dict):
    to_save = {}
    for k, v in farms_dict.items():
        to_save[k] = {
            "db_filename": v.get("db_filename", f"erp_{k}.db"),
            "color": v.get("color", "#4F46E5"),
            "buildings_count": v.get("buildings_count", 6),
            "pens_count": v.get("pens_count", 20)
        }
    with open(FARMS_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(to_save, f, ensure_ascii=False, indent=2)

FARM_CONFIG = load_farms()

def _migrate_schema_pg(db_file):
    # 프로세스당 농장별로 한 번만 실제로 점검한다 (db_adapter 가 기억해 둔다).
    # 아래 표·컬럼·트리거가 이미 다 있으면 DDL 을 실행하지 않는다 (잠금을 잡지 않게).
    db_adapter.ensure_schema(db_file, PG_DDL, extra_sql=[
        "ALTER TABLE testgroup_master ADD COLUMN IF NOT EXISTS location_mapping TEXT",
        "ALTER TABLE purchase ADD COLUMN IF NOT EXISTS unit TEXT",
        "ALTER TABLE item_master ADD COLUMN IF NOT EXISTS unit TEXT",
    ], required_tables=DB_TABLES,
       required_columns=[("testgroup_master", "location_mapping"), ("purchase", "unit"), ("item_master", "unit")],
       required_triggers=["trg_after_insert_purchase"])

# ========== DB 연결 관리 ==========
# Streamlit은 버튼/폼을 누를 때마다 스크립트를 처음부터 다시 실행한다.
# 이때 st.rerun() / st.stop() 은 예외로 동작하므로 스크립트 맨 아래의 conn.close() 가
# 실행되지 않고, 커밋되지 않은 연결이 잠금(RESERVED lock)을 쥔 채 남는다.
# 그 상태에서 다음 쓰기를 시도하면 "database is locked" 가 발생한다.
# 그래서 열어 둔 연결을 모두 등록해 두고, 다음 실행 시작 시점에 한 번에 정리한다.
def _conn_registry():
    if "_open_db_conns" not in st.session_state:
        st.session_state["_open_db_conns"] = []
    return st.session_state["_open_db_conns"]


def db_connect(db_path):
    # check_same_thread=False: 다음 실행(다른 스레드)에서 정리할 수 있도록 허용
    conn = sqlite3.connect(db_path, timeout=30, check_same_thread=False)
    if not USE_PG:  # SQLite 전용 설정 (Postgres 는 참조 무결성이 항상 켜져 있다)
        conn.execute("PRAGMA busy_timeout = 30000")
        conn.execute("PRAGMA journal_mode = WAL")  # 읽는 중에도 쓰기가 막히지 않는다
        conn.execute("PRAGMA foreign_keys = ON")   # 스키마에 선언된 참조 무결성을 실제로 적용
    _conn_registry().append(conn)
    return conn


def db_exists(db_file):
    """농장 DB가 만들어져 있는지. Supabase 모드에서는 파일이 없으므로 스키마에 표가 있는지 본다."""
    if USE_PG:
        return db_adapter.database_exists(db_file)
    return os.path.exists(db_file)


def close_stale_connections():
    """이전 실행에서 닫히지 않고 남은 연결을 롤백 후 모두 닫는다."""
    registry = _conn_registry()
    while registry:
        stale = registry.pop()
        for step in (stale.rollback, stale.close):
            try:
                step()
            except Exception:
                pass


def migrate_schema(db_file):
    """이미 만들어져 있는 DB에 나중에 추가된 표(예: cattle_item_usage_log)를 채워 넣는다.

    SQLITE_DDL 은 전부 'CREATE TABLE/TRIGGER IF NOT EXISTS' 라서 몇 번을 다시
    실행해도 안전하다. 이걸 안 하면, 기존 DB에서 새 기능을 처음 쓸 때
    pandas.to_sql() 이 PRIMARY KEY/UNIQUE 제약 없이 즉석에서 표를 만들어버려
    새로 만든 DB와 스키마가 미묘하게 달라진다.
    """
    if USE_PG:
        _migrate_schema_pg(db_file)
        return
    if not os.path.exists(db_file):
        return
    conn = sqlite3.connect(db_file, timeout=30)
    try:
        conn.executescript(SQLITE_DDL)
        # CREATE TABLE IF NOT EXISTS 는 이미 있는 표에는 손을 대지 않으므로,
        # 기존 표에 나중에 추가된 컬럼은 여기서 하나씩 채워 넣는다.
        for table, column, coldef in (
            ("testgroup_master", "location_mapping", "TEXT"),
            ("purchase", "unit", "TEXT"),
            ("item_master", "unit", "TEXT"),
        ):
            try:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {coldef}")
            except sqlite3.OperationalError:
                pass
        _migrate_cattle_default_check(conn)
        _migrate_item_category_check(conn)
        _migrate_purchase_trigger(conn)
        _repair_dangling_cattle_refs(conn)
        _repair_numpy_int_blobs(conn)
        conn.commit()
    finally:
        conn.close()


def _migrate_cattle_default_check(conn):
    """예전에 만들어진 DB는 cattle.feed_type/roughage_grade의 CHECK 제약에
    '표준' 값이 빠진 채로 남아 있다 (예: ('제한형','증량형')만 허용, 기본값 '증량형').
    엑셀 일괄등록은 값이 비어 있으면 '표준'을 기본값으로 넣는데, 이 CHECK 제약
    때문에 매번 "CHECK constraint failed"로 등록이 거부되고, 이 오류가 예전
    코드에서는 '이미 등록되어 건너뜀'으로 잘못 표시됐었다.
    CHECK 제약은 ALTER TABLE로 고칠 수 없으므로, 표를 새로 만들어 데이터를 옮긴다."""
    row = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='cattle'"
    ).fetchone()
    if not row or not row[0]:
        return
    if "IN ('표준', '제한형', '증량형')" in row[0] and "IN ('표준', '고급', '저급')" in row[0]:
        return  # 이미 최신 스키마

    # legacy_alter_table=ON 이 없으면 RENAME 시 다른 표(disease_record 등)의
    # REFERENCES cattle(...) 가 임시 표 이름으로 따라 바뀌어, 임시 표를 지운 뒤
    # 참조가 끊긴 표만 남는다.
    conn.execute("PRAGMA legacy_alter_table = ON")
    conn.execute("ALTER TABLE cattle RENAME TO cattle_old_migration")
    conn.execute("""
        CREATE TABLE cattle (
            cattle_id TEXT PRIMARY KEY,
            kpn TEXT,
            birth_date DATE,
            test_group_code TEXT REFERENCES testgroup_master(test_group_code),
            status TEXT CHECK (status IN ('사육', '출하', '폐사')) NOT NULL DEFAULT '사육',
            admission_date DATE,
            closure_date DATE,
            market_name TEXT,
            building TEXT,
            pen_number INTEGER,
            feed_type TEXT CHECK (feed_type IN ('표준', '제한형', '증량형')) DEFAULT '표준',
            roughage_grade TEXT CHECK (roughage_grade IN ('표준', '고급', '저급')) DEFAULT '표준',
            castration_date DATE,
            calf_price NUMERIC(12, 2) NOT NULL DEFAULT 0,
            commission_fee NUMERIC(12, 2) NOT NULL DEFAULT 0,
            transport_fee NUMERIC(12, 2) NOT NULL DEFAULT 0,
            initial_cost NUMERIC(12, 2) NOT NULL DEFAULT 0,
            insurance_value NUMERIC(12, 2) DEFAULT 0,
            insurance_premium NUMERIC(12, 2) DEFAULT 0,
            accident_date DATE,
            insurance_claim NUMERIC(12, 2) DEFAULT 0,
            memo TEXT
        )
    """)
    columns = [r[1] for r in conn.execute("PRAGMA table_info(cattle_old_migration)").fetchall()]
    col_list = ", ".join(columns)
    conn.execute(f"INSERT INTO cattle ({col_list}) SELECT {col_list} FROM cattle_old_migration")
    conn.execute("DROP TABLE cattle_old_migration")
    conn.execute("PRAGMA legacy_alter_table = OFF")


NUMERIC_COLUMNS = {
    "cattle": ["pen_number", "calf_price", "commission_fee", "transport_fee", "initial_cost",
               "insurance_value", "insurance_premium", "insurance_claim"],
    "item_master": ["current_stock", "moving_avg_price"],
    "purchase": ["quantity", "total_amount"],
    "monthly_usage": ["total_usage", "applied_price", "calculated_amount"],
    "monthly_fixedcost": ["total_billed_amount"],
    "cattle_cost_log": ["allocated_variable_cost", "allocated_fixed_cost"],
    "cattle_item_usage_log": ["allocated_usage", "allocated_amount"],
}


def _repair_numpy_int_blobs(conn):
    """pandas 에서 읽은 numpy.int64 를 그대로 INSERT 하면 sqlite3 가 숫자가 아닌 8바이트 BLOB 으로
    저장한다 (예: 적용단가 583 → b'G\\x02\\x00...'). 이런 값은 합계 계산에서 빠지고 화면 표에서도
    오류를 내므로, 원래 정수로 되돌린다."""
    for table, columns in NUMERIC_COLUMNS.items():
        for col in columns:
            try:
                rows = conn.execute(
                    f"SELECT rowid, {col} FROM {table} WHERE typeof({col}) = 'blob' AND length({col}) = 8"
                ).fetchall()
            except sqlite3.OperationalError:
                continue
            for rowid, raw in rows:
                conn.execute(
                    f"UPDATE {table} SET {col} = ? WHERE rowid = ?",
                    (int.from_bytes(raw, "little", signed=True), rowid),
                )


def _table_ddl(table):
    """SQLITE_DDL 에서 해당 표의 CREATE TABLE 문만 잘라 낸다."""
    marker = f"CREATE TABLE IF NOT EXISTS {table} ("
    start = SQLITE_DDL.find(marker)
    if start < 0:
        return None
    end = SQLITE_DDL.find(");", start)
    return SQLITE_DDL[start:end + 1]


def _migrate_purchase_trigger(conn):
    """예전 매입 트리거는 이동평균단가를 정수끼리 나눠(SQLite 는 정수/정수 = 정수) 소수점이 잘렸다.
    예: 첫 매입 1,620kg / 652,050원 → 402.5원이어야 하는데 402원으로 저장되어,
    매입분을 다 써도 산출총액(651,240원)이 매입금액보다 810원 적게 나왔다.
    트리거는 IF NOT EXISTS 로는 바뀌지 않으므로 지우고 다시 만들고, 이미 잘린 단가도 매입 내역으로 다시 계산한다."""
    row = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type='trigger' AND name='trg_after_insert_purchase'"
    ).fetchone()
    if row and row[0] and "* 1.0 /" in row[0]:
        return  # 이미 고쳐진 트리거
    conn.execute("DROP TRIGGER IF EXISTS trg_after_insert_purchase")
    start = SQLITE_DDL.index("CREATE TRIGGER IF NOT EXISTS trg_after_insert_purchase")
    end = SQLITE_DDL.index("END;", start) + len("END;")
    conn.execute(SQLITE_DDL[start:end])
    # 현재재고는 매입 때만 늘어나므로 이동평균단가 = 매입금액 합계 / 매입수량 합계 (매입 수정 시 재계산과 같은 식)
    conn.execute("""
        UPDATE item_master SET moving_avg_price = COALESCE((
            SELECT ROUND(SUM(total_amount) * 1.0 / SUM(quantity), 2)
            FROM purchase p WHERE p.item_code = item_master.item_code AND quantity > 0
        ), moving_avg_price)
    """)

def _migrate_item_category_check(conn):
    """예전 DB의 item_master.category CHECK 제약은 ('사료','조사료','약품') 만 허용해,
    화면에서 '기타저장품' 을 고르면 품목 등록이 "CHECK constraint failed" 로 거부된다.
    cattle 과 마찬가지로 표를 새로 만들어 데이터를 옮긴다."""
    row = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='item_master'"
    ).fetchone()
    if not row or not row[0] or "'기타저장품'" in row[0]:
        return  # 표가 없거나 이미 최신 스키마

    # purchase/monthly_usage/cattle_item_usage_log 의 REFERENCES item_master(...) 와
    # 매입 트리거 본문이 임시 표 이름으로 따라 바뀌지 않도록 legacy 모드로 RENAME 한다.
    conn.execute("PRAGMA legacy_alter_table = ON")
    conn.execute("ALTER TABLE item_master RENAME TO item_master_old_migration")
    conn.execute(_table_ddl("item_master"))
    columns = [r[1] for r in conn.execute("PRAGMA table_info(item_master_old_migration)").fetchall()]
    col_list = ", ".join(columns)
    conn.execute(f"INSERT INTO item_master ({col_list}) SELECT {col_list} FROM item_master_old_migration")
    conn.execute("DROP TABLE item_master_old_migration")
    conn.execute("PRAGMA legacy_alter_table = OFF")

def _repair_dangling_cattle_refs(conn):
    """cattle 표를 다시 만드는 과정에서 다른 표의 REFERENCES cattle(...) 이
    사라진 임시 표 이름을 가리킨 채 남은 DB를 되돌린다.
    참조가 끊긴 표는 foreign_keys=ON 상태에서 INSERT 하면 'no such table' 로 실패한다."""
    broken = [
        r[0] for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' "
            "AND name <> 'cattle_old_migration' AND sql LIKE '%cattle_old_migration%'"
        ).fetchall()
    ]
    if not broken:
        return

    conn.execute("PRAGMA legacy_alter_table = ON")
    for table in broken:
        ddl = _table_ddl(table)
        if not ddl:
            continue
        tmp = f"{table}_fkfix"
        conn.execute(f"ALTER TABLE {table} RENAME TO {tmp}")
        conn.execute(ddl)
        columns = [r[1] for r in conn.execute(f"PRAGMA table_info({tmp})").fetchall()]
        col_list = ", ".join(columns)
        conn.execute(f"INSERT INTO {table} ({col_list}) SELECT {col_list} FROM {tmp}")
        conn.execute(f"DROP TABLE {tmp}")
    conn.execute("PRAGMA legacy_alter_table = OFF")


# ========== 백업 ==========
def backup_db(db_file, reason="auto", dest_dir=None):
    """SQLite 스냅샷 백업. 단순 파일 복사와 달리 쓰기 도중에도 안전하다.

    dest_dir 을 지정하면 그 폴더에 저장하고, 오래된 백업 정리는 하지 않는다.
    """
    if not db_exists(db_file):
        return None
    name = os.path.splitext(os.path.basename(db_file))[0]
    out_dir = dest_dir or BACKUP_DIR
    os.makedirs(out_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  # 생성 날짜·시간을 파일명에 넣는다
    out = os.path.join(out_dir, "%s_%s_%s.db" % (name, stamp, reason))
    if USE_PG:
        # Supabase 데이터를 SQLite 파일로 내보낸다. 이 파일은 '백업 파일로 복원'에 그대로 쓸 수 있다.
        db_adapter.export_to_sqlite(db_file, out, SQLITE_DDL, DB_TABLES,
                                    trigger_names=["trg_after_insert_purchase"])
    else:
        conn = db_connect(db_file)
        try:
            conn.execute("VACUUM INTO ?", (out,))
        finally:
            conn.close()
    if dest_dir is None:
        _prune_backups(name)
    return out


def _prune_backups(name):
    """농장별로 최근 BACKUP_KEEP 개만 남기고 오래된 백업을 지운다."""
    files = sorted(glob.glob(os.path.join(BACKUP_DIR, name + "_*.db")), key=os.path.getmtime)
    for old in files[:-BACKUP_KEEP]:
        try:
            os.remove(old)
        except OSError:
            pass


def daily_backup(db_file):
    """그날 첫 실행 때 한 번만 자동 백업한다."""
    name = os.path.splitext(os.path.basename(db_file))[0]
    today = datetime.now().strftime("%Y-%m-%d")
    if glob.glob(os.path.join(BACKUP_DIR, "%s_%s_*_auto.db" % (name, today))):
        return
    try:
        backup_db(db_file, "auto")
    except Exception as e:
        st.sidebar.warning("자동 백업에 실패했습니다: %s" % e)


def restore_db(db_file, uploaded_bytes):
    """업로드한 백업 파일로 DB를 교체한다. 교체 전 현재 DB를 먼저 백업한다."""
    tmp_path = os.path.join(tempfile.gettempdir(), "erp_restore_%s.db" % datetime.now().strftime("%Y%m%d%H%M%S"))
    with open(tmp_path, "wb") as fh:
        fh.write(uploaded_bytes)
    try:
        probe = _sqlite3.connect(tmp_path)  # 업로드 파일은 Supabase 모드에서도 SQLite 파일이다
        if probe.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            probe.close()
            return False, "파일이 손상되었습니다. 다른 백업 파일을 사용하세요."
        tables = {r[0] for r in probe.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        probe.close()
        missing = {"cattle", "testgroup_master", "item_master"} - tables
        if missing:
            return False, "이 시스템의 DB 파일이 아닙니다. (없는 표: %s)" % ", ".join(sorted(missing))

        if db_exists(db_file):
            backup_db(db_file, "before-restore")
        close_stale_connections()
        if USE_PG:
            db_adapter.import_from_sqlite(db_file, tmp_path, PG_DDL, DB_TABLES,
                                          trigger_tables=["purchase"])
            return True, "복원이 완료되었습니다."
        for suffix in ("-wal", "-shm"):
            leftover = db_file + suffix
            if os.path.exists(leftover):
                os.remove(leftover)
        shutil.copyfile(tmp_path, db_file)
        return True, "복원이 완료되었습니다."
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


# ========== 접속 보호 (외부 공개 시 필수) ==========
def require_password():
    """환경변수 ERP_PASSWORD 가 설정된 경우에만 비밀번호를 요구한다 (로컬 사용 시에는 통과)."""
    expected = os.environ.get("ERP_PASSWORD")
    if not expected:
        try:
            expected = st.secrets.get("ERP_PASSWORD")
        except Exception:
            expected = None
    if not expected or st.session_state.get("_authed"):
        return
    st.markdown("## 🔒 대구축협 시험농장 관리 시스템")
    pw = st.text_input("접속 비밀번호", type="password")
    if pw:
        if pw == expected:
            st.session_state["_authed"] = True
            st.rerun()
        st.error("비밀번호가 올바르지 않습니다.")
    st.stop()


def init_db(farm_name):
    cfg = FARM_CONFIG[farm_name]
    db_file = cfg["db_file"]

    # 파일을 지우기 전에 열려 있는 연결을 먼저 끊어야 잠금이 풀린다.
    close_stale_connections()

    if USE_PG:
        db_adapter.reset_database(db_file, PG_DDL)
        return

    if os.path.exists(db_file):
        try:
            os.remove(db_file)
        except OSError as e:
            st.error(
                f"'{db_file}' 파일을 삭제할 수 없습니다. 다른 프로그램(DB 뷰어, 구글 드라이브 동기화 등)이 "
                f"파일을 사용 중인지 확인한 뒤 다시 시도하세요. ({e})"
            )
            st.stop()

    conn = db_connect(db_file)
    conn.executescript(SQLITE_DDL)
    conn.commit()
    conn.close()

def settlement_cattle(conn, settlement_month):
    """정산월에 배분받을 개체와 각자의 사육일수(rearing_days)를 구한다.
    정산 실행과 정산 전 미리보기가 같은 기준으로 두수를 세도록 한 곳에 모아 둔다."""
    import calendar
    from datetime import datetime, timedelta
    # status = '사육' 로만 걸러내면 이번 정산월 도중에 출하/폐사한 개체가 통째로 빠지고,
    # 그 개체가 실제로 소비한 사료·고정비 몫이 남은 개체에게 그대로 전가된다.
    # 그래서 상태와 무관하게 "이번 정산월에 하루라도 사육 이력이 겹치는 개체"를 모두 포함시키고,
    # 실제 겹치는 일수는 아래 calc_days() 에서 admission_date/closure_date 로 정확히 계산한다.
    query = """
        SELECT cattle_id, test_group_code, admission_date, closure_date
        FROM cattle
        WHERE (admission_date IS NULL OR strftime('%Y-%m', admission_date) <= ?)
        AND (closure_date IS NULL OR strftime('%Y-%m', closure_date) >= ?)
    """
    cattle_df = pd.read_sql(query, conn, params=(settlement_month, settlement_month))
    if cattle_df.empty:
        return cattle_df.assign(rearing_days=pd.Series(dtype=int))

    year, month = map(int, settlement_month.split('-'))
    last_day = calendar.monthrange(year, month)[1]
    month_start = datetime(year, month, 1)
    month_end = datetime(year, month, last_day)

    def calc_days(row):
        adm = row['admission_date']
        # 입식 당일은 절식하므로 사료·고정비 배분에서 빼고, 입식 다음 날부터 사육일수로 센다.
        # (예: 4/30 입식 개체는 4월 사육일수 0일 → 4월 정산 대상에서 제외, 5월부터 배분)
        if pd.notna(adm) and str(adm).strip():
            adm_d = datetime.strptime(str(adm)[:10], '%Y-%m-%d') + timedelta(days=1)
        else:
            adm_d = month_start
        start = max(month_start, adm_d)

        # 이번 달 중 출하/폐사(closure_date)했다면 그 날짜까지만 사육일수로 센다.
        end = month_end
        clo = row['closure_date']
        if pd.notna(clo) and str(clo).strip():
            clo_d = datetime.strptime(str(clo)[:10], '%Y-%m-%d')
            end = min(month_end, clo_d)

        if start > end: return 0
        return (end - start).days + 1

    cattle_df['rearing_days'] = cattle_df.apply(calc_days, axis=1)
    return cattle_df[cattle_df['rearing_days'] > 0]


def distribute_monthly_costs(db_file, settlement_month, replace=False):
    conn = db_connect(db_file)
    if replace:
        # 재정산: 이 달에 이미 적재된 배분 내역을 지우고 새로 계산한다.
        # 아래 적재까지 한 트랜잭션이라, 도중에 실패하면 지운 내역도 함께 되돌아간다.
        conn.execute("DELETE FROM cattle_item_usage_log WHERE settlement_month = ?", (settlement_month,))
        conn.execute("DELETE FROM cattle_cost_log WHERE settlement_month = ?", (settlement_month,))

    cattle_df = settlement_cattle(conn, settlement_month)
    if cattle_df.empty:
        conn.close()
        return False, "해당 월에 실제 사육일수가 있는 개체가 없습니다."

    total_farm_days = cattle_df['rearing_days'].sum()
    group_days = cattle_df.groupby('test_group_code')['rearing_days'].sum().to_dict()

    fixed_cost_df = pd.read_sql("SELECT SUM(total_billed_amount) as total_fixed_cost FROM monthly_fixedcost WHERE settlement_month = ?", conn, params=(settlement_month,))
    total_fixed_cost = float(fixed_cost_df['total_fixed_cost'].iloc[0]) if pd.notna(fixed_cost_df['total_fixed_cost'].iloc[0]) else 0.0

    usage_df = pd.read_sql("SELECT test_group_code, SUM(calculated_amount) as total_variable_cost FROM monthly_usage WHERE settlement_month = ? GROUP BY test_group_code", conn, params=(settlement_month,))
    group_vcost = dict(zip(usage_df['test_group_code'], usage_df['total_variable_cost']))

    item_usage_df = pd.read_sql("SELECT test_group_code, item_code, SUM(total_usage) as total_qty, SUM(calculated_amount) as total_amt FROM monthly_usage WHERE settlement_month = ? GROUP BY test_group_code, item_code", conn, params=(settlement_month,))

    log_data = []
    item_log_data = []
    for _, row in cattle_df.iterrows():
        t_group = row['test_group_code']
        days = row['rearing_days']
        
        f_cost = (total_fixed_cost * days / total_farm_days) if total_farm_days > 0 else 0
        g_days = group_days.get(t_group, 0)
        v_total = float(group_vcost.get(t_group, 0))
        v_cost = (v_total * days / g_days) if g_days > 0 else 0
        
        log_data.append({
            'cattle_id': row['cattle_id'],
            'settlement_month': settlement_month,
            'allocated_variable_cost': round(v_cost, 2),
            'allocated_fixed_cost': round(f_cost, 2)
        })

        group_items = item_usage_df[item_usage_df['test_group_code'] == t_group]
        for _, itm in group_items.iterrows():
            item_code = itm['item_code']
            item_qty = float(itm['total_qty'])
            item_amt = float(itm['total_amt'])
            
            alloc_qty = (item_qty * days / g_days) if g_days > 0 else 0
            alloc_amt = (item_amt * days / g_days) if g_days > 0 else 0
            
            item_log_data.append({
                'cattle_id': row['cattle_id'],
                'settlement_month': settlement_month,
                'item_code': item_code,
                'allocated_usage': round(alloc_qty, 2),
                'allocated_amount': round(alloc_amt, 2)
            })
        
    log_df = pd.DataFrame(log_data)
    item_log_df = pd.DataFrame(item_log_data)
    
    try:
        log_df.to_sql('cattle_cost_log', conn, if_exists='append', index=False, method='multi')
        if not item_log_df.empty:
            item_log_df.to_sql('cattle_item_usage_log', conn, if_exists='append', index=False, method='multi')
        conn.close()
        return True, f"'{settlement_month}' 개체별 일할계산(사육일수 비례) 정산이 완료되었습니다. (총 {len(log_df)}마리 적용)"
    except Exception as e:
        conn.rollback()
        conn.close()
        return False, f"오류 발생 (이미 정산된 연월일 수 있습니다): {e}"


# ========== 결산 리포트 ==========
# 시험군 배지 색상 팔레트. 시험군 수가 팔레트보다 많으면 순서대로 다시 돌려쓴다.
_REPORT_GROUP_PALETTE = [
    ("bg-blue-100", "text-blue-700"),
    ("bg-amber-100", "text-amber-700"),
    ("bg-emerald-100", "text-emerald-700"),
    ("bg-rose-100", "text-rose-700"),
    ("bg-violet-100", "text-violet-700"),
    ("bg-cyan-100", "text-cyan-700"),
]


def list_settled_months(db_file):
    """cattle_cost_log 에 이미 정산 기록이 있는 연월 목록 (최신순)."""
    if not db_exists(db_file):
        return []
    conn = db_connect(db_file)
    try:
        df = pd.read_sql(
            "SELECT DISTINCT settlement_month FROM cattle_cost_log ORDER BY settlement_month DESC",
            conn,
        )
    except sqlite3.Error:
        return []
    finally:
        conn.close()
    return df['settlement_month'].tolist()


def generate_settlement_report(db_file, farm_name, settlement_month):
    """정산월 결산 리포트를 인쇄 가능한 HTML 문서(문자열)로 만든다.

    반환값: (성공여부, HTML 문자열 또는 오류 메시지)
    """
    conn = db_connect(db_file)
    try:
        log_df = pd.read_sql(
            """
            SELECT l.cattle_id, c.test_group_code, t.test_name,
                   l.allocated_variable_cost, l.allocated_fixed_cost
            FROM cattle_cost_log l
            LEFT JOIN cattle c ON l.cattle_id = c.cattle_id
            LEFT JOIN testgroup_master t ON c.test_group_code = t.test_group_code
            WHERE l.settlement_month = ?
            ORDER BY c.test_group_code, l.cattle_id
            """,
            conn,
            params=(settlement_month,),
        )
        
        item_log_df = pd.read_sql(
            """
            SELECT u.cattle_id, u.item_code, m.item_name, u.allocated_usage, u.allocated_amount
            FROM cattle_item_usage_log u
            LEFT JOIN item_master m ON u.item_code = m.item_code
            WHERE u.settlement_month = ?
            """,
            conn,
            params=(settlement_month,)
        )
    finally:
        conn.close()

    if log_df.empty:
        return False, "해당 연월에 정산된 내역이 없습니다. 먼저 '월말 정산' 탭에서 정산을 실행하세요."

    head_count = len(log_df)
    total_variable = float(log_df['allocated_variable_cost'].fillna(0).sum())
    total_fixed = float(log_df['allocated_fixed_cost'].fillna(0).sum())

    group_codes = sorted({g for g in log_df['test_group_code'].tolist() if g})
    color_map = {g: _REPORT_GROUP_PALETTE[i % len(_REPORT_GROUP_PALETTE)] for i, g in enumerate(group_codes)}

    def esc(value):
        return html.escape(str(value)) if pd.notna(value) else ""

    item_usage_dict = {}
    for cattle_id, group in item_log_df.groupby('cattle_id'):
        items_html = []
        for _, row in group.iterrows():
            item_name = row['item_name'] if pd.notna(row['item_name']) else row['item_code']
            qty = row['allocated_usage']
            amt = row['allocated_amount']
            items_html.append(f"<span class='inline-block bg-slate-100 text-slate-600 rounded px-2 py-1 text-xs mr-1 mb-1'>{esc(item_name)}: {qty:,.2f}kg ({amt:,.0f}원)</span>")
        item_usage_dict[cattle_id] = "".join(items_html)

    row_html_parts = []
    for _, row in log_df.iterrows():
        gcode = row['test_group_code']
        gname = row['test_name'] if pd.notna(row['test_name']) else (gcode or "미지정")
        bg, fg = color_map.get(gcode, ("bg-slate-100", "text-slate-700"))
        v_cost = float(row['allocated_variable_cost'] or 0)
        f_cost = float(row['allocated_fixed_cost'] or 0)
        
        items_str = item_usage_dict.get(row['cattle_id'], "<span class='text-xs text-slate-400'>내역 없음</span>")

        row_html_parts.append(f"""
        <tr class="hover:bg-slate-50 transition-colors border-b border-slate-100">
            <td class="py-4 px-5 text-sm font-bold text-slate-700 align-top">{esc(row['cattle_id'])}</td>
            <td class="py-4 px-5 text-sm text-center align-top">
                <span class="px-3 py-1 rounded-full text-xs font-bold {bg} {fg}">{esc(gname)}</span>
            </td>
            <td class="py-4 px-5 align-top">
                <div class="text-sm text-right font-medium text-slate-600 mb-2">{v_cost:,.0f} 원</div>
                <div class="text-right flex flex-wrap justify-end gap-1">{items_str}</div>
            </td>
            <td class="py-4 px-5 text-sm text-right font-medium text-slate-600 align-top">{f_cost:,.0f} 원</td>
            <td class="py-4 px-5 text-sm text-right font-bold text-indigo-600 bg-indigo-50/30 align-top">{v_cost + f_cost:,.0f} 원</td>
        </tr>""")

    generated_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    html_doc = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>{esc(settlement_month)} 시험농장 결산 리포트</title>
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://fonts.googleapis.com/css2?family=Pretendard:wght@400;500;700;800&display=swap" rel="stylesheet">
<style>
    body {{ font-family: 'Pretendard', sans-serif; }}
    @media print {{
        body {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; background-color: white !important; }}
        .no-print {{ display: none; }}
    }}
</style>
</head>
<body class="bg-slate-50 p-4 md:p-8 text-slate-800">
    <div class="max-w-[210mm] mx-auto bg-white p-10 md:p-12 shadow-2xl rounded-xl border border-slate-100">

        <div class="border-b-4 border-slate-900 pb-6 mb-10 flex justify-between items-end">
            <div>
                <p class="text-indigo-600 font-bold tracking-wider text-sm mb-1">{esc(farm_name)}</p>
                <h1 class="text-4xl font-extrabold text-slate-900">월간 원가 및 사양 결산서</h1>
            </div>
            <div class="text-right">
                <p class="text-2xl font-bold text-slate-700 bg-slate-100 px-4 py-1 rounded-md">{esc(settlement_month)}</p>
                <p class="text-slate-400 text-sm mt-2">출력일자: {generated_at}</p>
            </div>
        </div>

        <section class="mb-12">
            <h2 class="text-xl font-bold text-slate-800 mb-5 flex items-center">
                <span class="w-2 h-6 bg-indigo-600 rounded mr-3"></span>
                원가 배부 총괄 요약
            </h2>
            <div class="grid grid-cols-3 gap-6">
                <div class="bg-gradient-to-br from-slate-50 to-slate-100 p-6 rounded-xl border border-slate-200 shadow-sm">
                    <p class="text-sm text-slate-500 font-bold mb-1">당월 정산 두수</p>
                    <p class="text-3xl font-black text-slate-800">{head_count} <span class="text-lg font-medium text-slate-500">두</span></p>
                </div>
                <div class="bg-gradient-to-br from-indigo-50 to-indigo-100 p-6 rounded-xl border border-indigo-200 shadow-sm">
                    <p class="text-sm text-indigo-600 font-bold mb-1">당월 총 변동비 (사료/약품)</p>
                    <p class="text-3xl font-black text-indigo-900">{total_variable:,.0f} <span class="text-lg font-medium text-indigo-600">원</span></p>
                </div>
                <div class="bg-gradient-to-br from-emerald-50 to-emerald-100 p-6 rounded-xl border border-emerald-200 shadow-sm">
                    <p class="text-sm text-emerald-600 font-bold mb-1">당월 총 고정비 (인건비/전기 등)</p>
                    <p class="text-3xl font-black text-emerald-900">{total_fixed:,.0f} <span class="text-lg font-medium text-emerald-600">원</span></p>
                </div>
            </div>
        </section>

        <section>
            <h2 class="text-xl font-bold text-slate-800 mb-5 flex items-center">
                <span class="w-2 h-6 bg-emerald-500 rounded mr-3"></span>
                개체별 당월 원가 배부 명세서 (사육일수 비례 배분)
            </h2>
            <div class="overflow-hidden rounded-xl border border-slate-200 shadow-sm">
                <table class="min-w-full bg-white">
                    <thead class="bg-slate-900 text-white">
                        <tr>
                            <th class="py-4 px-5 text-left font-semibold text-sm">개체번호</th>
                            <th class="py-4 px-5 text-center font-semibold text-sm">시험군</th>
                            <th class="py-4 px-5 text-right font-semibold text-sm">배부 변동비 (금액 및 품목상세)</th>
                            <th class="py-4 px-5 text-right font-semibold text-sm">배부 고정비 (원)</th>
                            <th class="py-4 px-5 text-right font-semibold text-sm text-indigo-300">당월 총원가 (원)</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100">
                        {''.join(row_html_parts)}
                    </tbody>
                </table>
            </div>
        </section>

        <div class="mt-16 text-center text-sm font-medium text-slate-400 border-t border-slate-100 pt-6">
            본 문서는 '대구축협 시험농장 관리 시스템'에 의해 자동 생성되었습니다.
        </div>

        <div class="mt-8 text-center no-print">
            <button onclick="window.print()" class="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-3 px-8 rounded-full shadow-lg transition-transform transform hover:-translate-y-1">
                PDF 인쇄 및 저장
            </button>
        </div>
    </div>
</body>
</html>"""
    return True, html_doc


OVERALL_REPORT_CSS = """
:root {
  --bg: #f1f5f9; --card: #ffffff; --ink: #0f172a; --muted: #64748b; --line: #e2e8f0;
  --brand: #4F46E5; --green: #059669; --red: #DC2626; --blue: #2563EB; --amber: #D97706;
}
* { box-sizing: border-box; }
body {
  margin: 0; background: var(--bg); color: var(--ink);
  font-family: 'Pretendard', 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif;
  line-height: 1.5; -webkit-font-smoothing: antialiased;
}
.container { max-width: 1200px; margin: 0 auto; padding: 32px 24px 48px; }

.report-header {
  display: flex; align-items: center; justify-content: space-between; gap: 16px; flex-wrap: wrap;
  background: linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%); color: #fff;
  border-radius: 20px; padding: 24px 28px; box-shadow: 0 10px 30px rgba(67, 56, 202, 0.25);
}
.brand { display: flex; align-items: center; gap: 14px; }
.brand img { height: 48px; background: #fff; border-radius: 12px; padding: 4px; }
.brand h1 { margin: 0; font-size: 26px; font-weight: 800; letter-spacing: -0.02em; }
.date-badge {
  background: rgba(255, 255, 255, 0.15); border: 1px solid rgba(255, 255, 255, 0.3);
  padding: 8px 16px; border-radius: 999px; font-weight: 600; font-size: 14px; white-space: nowrap;
}

.section { margin-top: 36px; }
.section-title { display: flex; align-items: center; gap: 10px; font-size: 19px; font-weight: 800; margin: 0 0 14px; }
.section-title .no {
  display: inline-flex; align-items: center; justify-content: center; width: 28px; height: 28px;
  border-radius: 8px; background: var(--brand); color: #fff; font-size: 14px;
}

.kpi-grid { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 16px; }
.kpi {
  position: relative; overflow: hidden; background: var(--card); border: 1px solid var(--line);
  border-radius: 18px; padding: 20px;
  box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04), 0 8px 24px rgba(15, 23, 42, 0.07);
}
.kpi::before { content: ''; position: absolute; left: 0; right: 0; top: 0; height: 4px; background: var(--accent); }
.kpi.total { --accent: var(--brand); }
.kpi.breeding { --accent: var(--green); }
.kpi.dead { --accent: var(--red); }
.kpi.shipped { --accent: var(--blue); }
.kpi.cost { --accent: var(--amber); }
.kpi-label { font-size: 13px; font-weight: 600; color: var(--muted); }
.kpi-value {
  margin-top: 12px; font-size: 30px; font-weight: 800; letter-spacing: -0.02em;
  color: var(--accent); font-variant-numeric: tabular-nums;
}
.kpi-value small { font-size: 15px; font-weight: 700; color: var(--muted); margin-left: 3px; }
.kpi-sub { margin-top: 4px; font-size: 12px; color: var(--muted); }

.card {
  background: var(--card); border: 1px solid var(--line); border-radius: 18px; padding: 20px;
  box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04), 0 8px 24px rgba(15, 23, 42, 0.06); margin-bottom: 16px;
}
.card-head { display: flex; align-items: baseline; justify-content: space-between; gap: 8px; flex-wrap: wrap; margin-bottom: 12px; }
.card-head h3 { margin: 0; font-size: 15px; font-weight: 700; }
.card-head span { font-size: 12px; color: var(--muted); }
.chart-box { position: relative; height: 320px; }
.donut-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 16px; }
.donut-item .chart-box { height: 300px; }
.donut-item p { margin: 8px 0 0; text-align: center; font-weight: 700; font-size: 14px; }
.chart-fallback { display: flex; align-items: center; justify-content: center; height: 100%; color: var(--muted); font-size: 13px; }

.table-wrap { max-height: 460px; overflow: auto; background: var(--card); border: 1px solid var(--line); border-radius: 14px; }
.data-table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 14px; }
.data-table th {
  position: sticky; top: 0; z-index: 2; background: #1e293b; color: #fff; font-weight: 600;
  padding: 11px 14px; text-align: right; white-space: nowrap;
}
.data-table td {
  padding: 10px 14px; border-bottom: 1px solid var(--line); text-align: right;
  font-variant-numeric: tabular-nums; white-space: nowrap;
}
.data-table th.left, .data-table td.left { text-align: left; }
.data-table tbody tr:nth-child(even) td { background: #f8fafc; }
.data-table tbody tr:hover td { background: #e0e7ff; }
.data-table tfoot td {
  position: sticky; bottom: 0; z-index: 1; background: #f1f5f9; font-weight: 800;
  border-top: 2px solid #cbd5e1; border-bottom: none;
}
.dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 8px; vertical-align: middle; }
.farm-sub { display: flex; align-items: center; font-size: 16px; font-weight: 700; margin: 22px 0 10px; }
.empty { background: var(--card); border: 1px dashed #cbd5e1; border-radius: 14px; padding: 24px; text-align: center; color: var(--muted); }

.footer { margin-top: 40px; padding-top: 16px; border-top: 1px solid var(--line); text-align: center; color: var(--muted); font-size: 12px; }
.print-wrap { margin-top: 20px; text-align: center; }
.print-btn {
  background: var(--brand); color: #fff; border: none; border-radius: 999px; padding: 12px 32px;
  font-size: 15px; font-weight: 700; cursor: pointer; box-shadow: 0 8px 20px rgba(79, 70, 229, 0.3);
  font-family: inherit;
}
.print-btn:hover { background: #4338ca; }

@media (max-width: 1024px) {
  .kpi-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); }
}
@media (max-width: 640px) {
  .container { padding: 16px 16px 32px; }
  .report-header { padding: 20px; border-radius: 16px; }
  .brand h1 { font-size: 20px; }
  .brand img { height: 38px; }
  .kpi-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
  .kpi { padding: 16px; }
  .kpi-value { font-size: 24px; }
  .chart-box, .donut-item .chart-box { height: 260px; }
  .data-table { font-size: 13px; }
  .data-table th, .data-table td { padding: 9px 10px; }
}
@media (max-width: 380px) {
  .kpi-grid { grid-template-columns: minmax(0, 1fr); }
}
@media print {
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body { background: #fff; }
  .container { max-width: none; padding: 0; }
  .no-print { display: none !important; }
  .kpi-grid { grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 10px; }
  .kpi, .card { box-shadow: none; break-inside: avoid; }
  .table-wrap { max-height: none; overflow: visible; }
  .data-table th, .data-table tfoot td { position: static; }
  .data-table tbody tr:hover td { background: inherit; }
  .farm-sub { break-after: avoid; }
}
"""

OVERALL_REPORT_JS = """
(function () {
  var data = JSON.parse(document.getElementById('report-data').textContent);
  if (typeof Chart === 'undefined') {
    document.querySelectorAll('.chart-box').forEach(function (el) {
      el.innerHTML = '<div class="chart-fallback">차트를 불러오지 못했습니다 (인터넷 연결 필요)</div>';
    });
    return;
  }
  var nf = new Intl.NumberFormat('ko-KR');
  var FONT = "'Pretendard', 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif";
  Chart.defaults.font.family = FONT;
  Chart.defaults.color = '#475569';
  Chart.defaults.animation.duration = 600;

  var barValueLabels = {
    id: 'barValueLabels',
    afterDatasetsDraw: function (chart) {
      var c = chart.ctx;
      c.save();
      c.textAlign = 'center'; c.textBaseline = 'bottom';
      c.font = '700 13px ' + FONT;
      chart.data.datasets.forEach(function (ds, di) {
        var meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;
        c.fillStyle = di === 0 ? '#4338ca' : '#b45309';
        meta.data.forEach(function (bar, i) {
          c.fillText(nf.format(ds.data[i]) + ds.unit, bar.x, bar.y - 6);
        });
      });
      c.restore();
    }
  };

  var barEl = document.getElementById('farmBarChart');
  if (barEl && data.farms.length) {
    new Chart(barEl, {
      type: 'bar',
      data: {
        labels: data.farms.map(function (f) { return f.name; }),
        datasets: [
          { label: '전체 입식 (두)', unit: '두', data: data.farms.map(function (f) { return f.total; }),
            backgroundColor: 'rgba(79, 70, 229, 0.85)', borderRadius: 8, maxBarThickness: 64, yAxisID: 'y' },
          { label: '총 구입비용 (만원)', unit: '만원', data: data.farms.map(function (f) { return f.cost_man; }),
            backgroundColor: 'rgba(217, 119, 6, 0.85)', borderRadius: 8, maxBarThickness: 64, yAxisID: 'y1' }
        ]
      },
      plugins: [barValueLabels],
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'rectRounded' } },
          tooltip: { callbacks: { label: function (ctx) { return ' ' + ctx.dataset.label + ': ' + nf.format(ctx.parsed.y); } } }
        },
        scales: {
          x: { grid: { display: false }, ticks: { font: { weight: '700' } } },
          y: { beginAtZero: true, grace: '15%', position: 'left', title: { display: true, text: '입식 두수 (두)' },
               ticks: { callback: function (v) { return nf.format(v); } } },
          y1: { beginAtZero: true, grace: '15%', position: 'right', grid: { drawOnChartArea: false },
                title: { display: true, text: '총 구입비용 (만원)' },
                ticks: { callback: function (v) { return nf.format(v); } } }
        }
      }
    });
  }

  var palette = ['#4F46E5', '#059669', '#D97706', '#2563EB', '#DC2626', '#7C3AED',
                 '#0891B2', '#DB2777', '#65A30D', '#EA580C', '#475569', '#0D9488'];

  function centerText(total) {
    return {
      id: 'centerText',
      afterDraw: function (chart) {
        var meta = chart.getDatasetMeta(0);
        if (!meta || !meta.data.length) return;
        var x = meta.data[0].x, y = meta.data[0].y, c = chart.ctx;
        c.save();
        c.textAlign = 'center'; c.textBaseline = 'middle';
        c.fillStyle = '#0f172a'; c.font = '800 22px ' + FONT;
        c.fillText(nf.format(total) + '두', x, y - 9);
        c.fillStyle = '#64748b'; c.font = '600 11px ' + FONT;
        c.fillText('총 구입', x, y + 13);
        c.restore();
      }
    };
  }

  // 같은 우시장은 어느 농장 도넛에서든 같은 색으로 보이게 한다.
  var marketColor = {};
  data.markets.forEach(function (m) {
    m.items.forEach(function (x) {
      if (!(x.market in marketColor)) {
        marketColor[x.market] = palette[Object.keys(marketColor).length % palette.length];
      }
    });
  });

  data.markets.forEach(function (m, i) {
    var el = document.getElementById('donut-' + i);
    if (!el) return;
    var total = m.items.reduce(function (s, x) { return s + x.count; }, 0);
    new Chart(el, {
      type: 'doughnut',
      data: {
        labels: m.items.map(function (x) { return x.market; }),
        datasets: [{
          data: m.items.map(function (x) { return x.count; }),
          backgroundColor: m.items.map(function (x) { return marketColor[x.market]; }),
          borderColor: '#ffffff', borderWidth: 2, hoverOffset: 8
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '62%',
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 12, usePointStyle: true } },
          tooltip: { callbacks: { label: function (ctx) {
            return ' ' + ctx.label + ': ' + nf.format(ctx.parsed) + '두 (' + (ctx.parsed / total * 100).toFixed(1) + '%)';
          } } }
        }
      },
      plugins: [centerText(total)]
    });
  });

  window.addEventListener('beforeprint', function () {
    Object.values(Chart.instances).forEach(function (c) { c.resize(); });
  });
})();
"""


# ========== 월말 팔공 입력 자료 (인쇄용) ==========
ACCOUNTING_SHEET_CSS = """
@page { size: A4 portrait; margin: 12mm 10mm; }
* { box-sizing: border-box; }
body { font-family: 'Pretendard', 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; color: #111; margin: 0;
       background: #f3f4f6; font-size: 12px; }
.sheet { max-width: 900px; margin: 16px auto; background: #fff; padding: 22px 26px; box-shadow: 0 2px 10px rgba(0,0,0,.08); }
.head { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; border-bottom: 2px solid #111;
        padding-bottom: 10px; margin-bottom: 14px; }
.head h1 { font-size: 20px; margin: 0 0 4px; }
.head .meta { color: #444; line-height: 1.6; }
table.sign { border-collapse: collapse; }
table.sign th, table.sign td { border: 1px solid #333; width: 64px; text-align: center; padding: 3px; font-size: 11px; }
table.sign td { height: 46px; }
h2 { font-size: 14px; margin: 18px 0 6px; padding-left: 8px; border-left: 4px solid #2E6B57; }
.keep { break-inside: avoid-page; page-break-inside: avoid; }
h3 { font-size: 13px; margin: 10px 0 4px; }
.warn { color: #B42318; font-weight: 700; font-size: 11px; margin: 2px 0 6px; }
.note { color: #555; font-size: 11px; margin: 0 0 6px; }
table.t { width: 100%; border-collapse: collapse; margin-bottom: 4px; }
table.t th, table.t td { border: 1px solid #9ca3af; padding: 4px 6px; }
table.t th { background: #eef2f0; font-weight: 700; text-align: center; white-space: nowrap; }
table.t td.n { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
table.t td.c { text-align: center; white-space: nowrap; }
table.t tr.sub td { background: #f7f7f7; font-weight: 700; }
table.t tr.total td { background: #e5ece8; font-weight: 800; }
table.t td.chk { text-align: center; width: 34px; color: #666; }
.summary { display: grid; grid-template-columns: repeat(4, 1fr); gap: 6px; }
.summary div { border: 1px solid #9ca3af; padding: 6px 8px; }
.summary b { display: block; font-size: 11px; color: #444; font-weight: 600; }
.summary span { font-size: 15px; font-weight: 800; font-variant-numeric: tabular-nums; }
.empty { color: #666; padding: 6px 0; }
.toolbar { max-width: 900px; margin: 12px auto 0; text-align: right; }
.toolbar button { font-size: 14px; padding: 8px 18px; border: 0; border-radius: 8px; background: #2E6B57; color: #fff; cursor: pointer; }
thead { display: table-header-group; }
tr { page-break-inside: avoid; }
@media print {
  body { background: #fff; }
  .sheet { box-shadow: none; margin: 0; max-width: none; padding: 0; }
  .toolbar { display: none; }
  h2 { page-break-after: avoid; }
}
"""


# 회사 회계 프로그램 전표 (계정 대응표). 다른 전표가 생기면 여기에 같은 형식으로 추가한다.
#   amount 의 값: "fixed_total" = 그 달 고정비 전체 합계, "usage_total" = 그 달 시험군별 사용량(산출액) 전체 합계,
#                 ["인건비", ...] = 고정비 내역 중 해당 지출 항목들의 합계
#   basis: 인쇄물의 '근거' 칸에 적는 설명
ACCOUNTING_JOURNALS = [
    {
        "name": "사양관리비 원가배부",
        "lines": [
            ("차", "21610203", "동물<소>-사양관리비원", "fixed_total", "농장 고정비 합계"),
            ("대", "21610401", "동물<사양관리비>임금", ["인건비", "시험사양수고비"], "인건비, 시험사양수고비"),
            ("대", "21610409", "동물<사양관리비>기타", ["전기세", "CCTV사용료", "우수등급장려금"], "전기세, CCTV사용료, 우수등급장려금"),
            ("대", "21610406", "동물<사양관리비>-가축공제", ["가축보험료"], "가축보험료"),
        ],
    },
    {
        "name": "기타저장품 원가배부",
        "lines": [
            ("차", "21610204", "동물<소>기타저장품원", "usage_total", "시험군별 사용량 (시험군 상관없이 전체)"),
            ("대", "21492108", "기타저장품<생축>", "usage_total", "시험군별 사용량 (시험군 상관없이 전체)"),
        ],
    },
]


def build_journal_html(fixed, usage_total, esc, won, chk):
    """ACCOUNTING_JOURNALS 대응표대로 그 달 전표(차변·대변) 표를 만든다."""
    fixed_by_item = fixed.groupby("expense_item")["total_billed_amount"].sum()
    fixed_total = float(fixed["total_billed_amount"].sum())
    mapped = {item for j in ACCOUNTING_JOURNALS for *_, amount, _ in j["lines"] if isinstance(amount, list) for item in amount}
    parts = []
    for j in ACCOUNTING_JOURNALS:
        rows, dr_sum, cr_sum = [], 0.0, 0.0
        for side, code, name, amount, basis in j["lines"]:
            if amount == "fixed_total":
                value = fixed_total
            elif amount == "usage_total":
                value = float(usage_total)
            else:
                value = float(sum(fixed_by_item.get(item, 0) for item in amount))
                hit = [f"{item} {won(fixed_by_item[item])}" for item in amount if item in fixed_by_item]
                basis = " + ".join(hit) if hit else f"{basis} (이 달 없음)"
            if side == "차":
                dr_sum += value
            else:
                cr_sum += value
            dr, cr = (won(value), "") if side == "차" else ("", won(value))
            rows.append(f'<tr><td class="c">{esc(side)}</td><td class="c">{esc(code)}</td><td>{esc(name)}</td>'
                        f'<td class="n">{dr}</td><td class="n">{cr}</td><td>{esc(basis)}</td>{chk}</tr>')
        ok = abs(dr_sum - cr_sum) < 0.5
        rows.append(f'<tr class="total"><td colspan="3">합계{"" if ok else " — 차변·대변 불일치"}</td>'
                    f'<td class="n">{won(dr_sum)}</td><td class="n">{won(cr_sum)}</td><td></td><td></td></tr>')
        note = ""
        if not ok:
            unmapped = [f"{item} {won(v)}" for item, v in fixed_by_item.items() if item not in mapped]
            note = ('<p class="warn">※ 차변과 대변이 ' + won(abs(dr_sum - cr_sum)) + '원 다릅니다.'
                    + (f' 계정이 정해지지 않은 고정비 항목: {esc(", ".join(unmapped))}' if unmapped else '') + '</p>')
        parts.append(
            f'<div class="keep"><h3>{esc(j["name"])}</h3><table class="t"><thead><tr><th>차/대</th><th>계정코드</th><th>계정명</th>'
            f'<th>차변(원)</th><th>대변(원)</th><th>근거</th><th>입력<br>확인</th></tr></thead><tbody>'
            + "".join(rows) + f"</tbody></table>{note}</div>")
    return "".join(parts)


def generate_accounting_sheet(db_file, farm_name, month):
    """한 달 치 자료를 회사 회계 프로그램에 옮겨 적기 좋게 A4 인쇄용 HTML 로 만든다 (읽기 전용).

    1 요약 / 2 매입(입고) 내역 / 3 품목 수불부(기초·입고·출고·기말) / 4 고정비 / 5 시험군별 원가 배분 /
    6 개체 증감(입식·출하·폐사) / 7 월말 사육 개체 장부가. 줄마다 '입력확인' 칸(□)을 둔다.
    반환값: (성공여부, HTML 또는 오류 메시지)
    """
    esc = lambda v: html.escape("" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v))
    won = lambda v: f"{float(v):,.0f}" if pd.notna(v) else "-"
    qty = lambda v: (f"{float(v):,.2f}".rstrip("0").rstrip(".") if pd.notna(v) else "-")
    chk = '<td class="chk">□</td>'

    conn = db_connect(db_file)
    try:
        purchases = pd.read_sql("""
            SELECT p.purchase_date, i.category, i.item_name, p.item_code, p.quantity, COALESCE(p.unit, i.unit) AS unit, p.total_amount
            FROM purchase p JOIN item_master i ON i.item_code = p.item_code
            WHERE substr(p.purchase_date, 1, 7) = ?
            ORDER BY i.category, p.purchase_date, i.item_name
        """, conn, params=(month,))
        # 품목 수불: 기초 = 이 달 이전 매입 - 이 달 이전 사용, 입고 = 이 달 매입, 출고 = 이 달 사용
        ledger = pd.read_sql("""
            SELECT i.item_code, i.item_name, i.category, i.unit,
              COALESCE((SELECT SUM(quantity) FROM purchase p WHERE p.item_code = i.item_code AND substr(p.purchase_date, 1, 7) < ?), 0)
            - COALESCE((SELECT SUM(total_usage) FROM monthly_usage u WHERE u.item_code = i.item_code AND u.settlement_month < ?), 0) AS open_qty,
              COALESCE((SELECT SUM(total_amount) FROM purchase p WHERE p.item_code = i.item_code AND substr(p.purchase_date, 1, 7) < ?), 0)
            - COALESCE((SELECT SUM(calculated_amount) FROM monthly_usage u WHERE u.item_code = i.item_code AND u.settlement_month < ?), 0) AS open_amt,
              COALESCE((SELECT SUM(quantity) FROM purchase p WHERE p.item_code = i.item_code AND substr(p.purchase_date, 1, 7) = ?), 0) AS in_qty,
              COALESCE((SELECT SUM(total_amount) FROM purchase p WHERE p.item_code = i.item_code AND substr(p.purchase_date, 1, 7) = ?), 0) AS in_amt,
              COALESCE((SELECT SUM(total_usage) FROM monthly_usage u WHERE u.item_code = i.item_code AND u.settlement_month = ?), 0) AS out_qty,
              COALESCE((SELECT SUM(calculated_amount) FROM monthly_usage u WHERE u.item_code = i.item_code AND u.settlement_month = ?), 0) AS out_amt
            FROM item_master i ORDER BY i.category, i.item_name
        """, conn, params=(month,) * 8)
        fixed = pd.read_sql(
            "SELECT expense_item, total_billed_amount FROM monthly_fixedcost WHERE settlement_month = ? ORDER BY fixed_cost_id",
            conn, params=(month,))
        alloc = pd.read_sql("""
            SELECT COALESCE(t.test_name, '(시험군 없음)') AS grp, COUNT(*) AS heads,
                   SUM(l.allocated_variable_cost) AS var_cost, SUM(l.allocated_fixed_cost) AS fix_cost
            FROM cattle_cost_log l
            LEFT JOIN cattle c ON c.cattle_id = l.cattle_id
            LEFT JOIN testgroup_master t ON t.test_group_code = c.test_group_code
            WHERE l.settlement_month = ?
            GROUP BY 1 ORDER BY 1
        """, conn, params=(month,))
        admitted = pd.read_sql("""
            SELECT c.cattle_id, c.admission_date, t.test_name, c.market_name, c.calf_price, c.commission_fee, c.transport_fee, c.initial_cost
            FROM cattle c LEFT JOIN testgroup_master t ON t.test_group_code = c.test_group_code
            WHERE substr(c.admission_date, 1, 7) = ? ORDER BY c.admission_date, c.cattle_id
        """, conn, params=(month,))
        closed = pd.read_sql("""
            SELECT c.cattle_id, c.status, c.closure_date, t.test_name, c.initial_cost,
                   COALESCE((SELECT SUM(h.allocated_variable_cost + h.allocated_fixed_cost) FROM cattle_cost_log h
                             WHERE h.cattle_id = c.cattle_id AND h.settlement_month <= ?), 0) AS raised_cost
            FROM cattle c LEFT JOIN testgroup_master t ON t.test_group_code = c.test_group_code
            WHERE c.status IN ('출하', '폐사') AND substr(c.closure_date, 1, 7) = ?
            ORDER BY c.status, c.closure_date, c.cattle_id
        """, conn, params=(month, month))
        # 월말 사육 개체: 이 달 말까지 입식했고, 이 달 말까지 출하·폐사하지 않은 개체
        on_hand = pd.read_sql("""
            SELECT COALESCE(t.test_name, '(시험군 없음)') AS grp, COUNT(*) AS heads, SUM(c.initial_cost) AS buy_cost,
                   SUM(COALESCE((SELECT SUM(h.allocated_variable_cost + h.allocated_fixed_cost) FROM cattle_cost_log h
                                  WHERE h.cattle_id = c.cattle_id AND h.settlement_month <= ?), 0)) AS raised_cost
            FROM cattle c LEFT JOIN testgroup_master t ON t.test_group_code = c.test_group_code
            WHERE substr(c.admission_date, 1, 7) <= ?
              AND (c.closure_date IS NULL OR c.closure_date = '' OR substr(c.closure_date, 1, 7) > ?)
            GROUP BY 1 ORDER BY 1
        """, conn, params=(month, month, month))
    except Exception as e:
        return False, f"자료를 읽는 중 오류가 발생했습니다: {e}"
    finally:
        conn.close()

    ledger["close_qty"] = ledger["open_qty"] + ledger["in_qty"] - ledger["out_qty"]
    ledger["close_amt"] = ledger["open_amt"] + ledger["in_amt"] - ledger["out_amt"]
    active = ledger[(ledger[["open_qty", "in_qty", "out_qty", "close_qty"]].abs() > 1e-9).any(axis=1)]
    alloc["total"] = alloc["var_cost"] + alloc["fix_cost"]
    on_hand["book"] = on_hand["buy_cost"] + on_hand["raised_cost"]
    closed["book"] = closed["initial_cost"] + closed["raised_cost"]
    settled = not alloc.empty

    # ---- 1. 요약 ----
    summary_items = [
        ("매입(입고) 합계", f"{won(purchases['total_amount'].sum())}원"),
        ("사용(출고) = 변동비", f"{won(active['out_amt'].sum())}원"),
        ("고정비 합계", f"{won(fixed['total_billed_amount'].sum())}원"),
        ("원가 배분 합계", f"{won(alloc['total'].sum())}원" if settled else "미정산"),
        ("입식", f"{len(admitted):,}두 · {won(admitted['initial_cost'].sum())}원"),
        ("출하 / 폐사", f"{int((closed['status'] == '출하').sum())}두 / {int((closed['status'] == '폐사').sum())}두"),
        ("월말 사육두수", f"{int(on_hand['heads'].sum()):,}두"),
        ("월말 사육 장부가", f"{won(on_hand['book'].sum())}원"),
    ]
    summary_html = '<div class="summary">' + "".join(f"<div><b>{esc(k)}</b><span>{esc(v)}</span></div>" for k, v in summary_items) + "</div>"
    if not settled:
        summary_html += f'<p class="note">※ {esc(month)} 은 아직 월말 정산 전입니다. 원가 배분·장부가의 사육비는 정산 후 반영됩니다.</p>'

    # ---- 2. 매입 내역 (분류별 소계) ----
    if purchases.empty:
        purchase_html = '<div class="empty">이 달 매입 내역이 없습니다.</div>'
    else:
        rows = []
        for cat, g in purchases.groupby("category", sort=False):
            for _, r in g.iterrows():
                price = r["total_amount"] / r["quantity"] if r["quantity"] else None
                rows.append(f'<tr><td class="c">{esc(r["purchase_date"])}</td><td class="c">{esc(cat)}</td><td>{esc(r["item_name"])}</td>'
                            f'<td class="n">{qty(r["quantity"])}</td><td class="c">{esc(r["unit"])}</td>'
                            f'<td class="n">{qty(price)}</td><td class="n">{won(r["total_amount"])}</td>{chk}</tr>')
            rows.append(f'<tr class="sub"><td colspan="6">{esc(cat)} 소계 ({len(g)}건)</td><td class="n">{won(g["total_amount"].sum())}</td><td></td></tr>')
        rows.append(f'<tr class="total"><td colspan="6">매입 합계 ({len(purchases)}건)</td><td class="n">{won(purchases["total_amount"].sum())}</td><td></td></tr>')
        purchase_html = ('<table class="t"><thead><tr><th>매입일자</th><th>분류</th><th>품목</th><th>수량</th><th>단위</th>'
                         '<th>단가</th><th>금액(원)</th><th>입력<br>확인</th></tr></thead><tbody>' + "".join(rows) + "</tbody></table>")

    # ---- 3. 품목 수불부 ----
    if active.empty:
        ledger_html = '<div class="empty">이 달 수불 내역이 없습니다.</div>'
    else:
        rows = []
        for cat, g in active.groupby("category", sort=False):
            for _, r in g.iterrows():
                rows.append(
                    f'<tr><td class="c">{esc(cat)}</td><td>{esc(r["item_name"])}</td><td class="c">{esc(r["unit"])}</td>'
                    f'<td class="n">{qty(r["open_qty"])}</td><td class="n">{won(r["open_amt"])}</td>'
                    f'<td class="n">{qty(r["in_qty"])}</td><td class="n">{won(r["in_amt"])}</td>'
                    f'<td class="n">{qty(r["out_qty"])}</td><td class="n">{won(r["out_amt"])}</td>'
                    f'<td class="n">{qty(r["close_qty"])}</td><td class="n">{won(r["close_amt"])}</td>{chk}</tr>')
            rows.append(f'<tr class="sub"><td colspan="3">{esc(cat)} 소계</td><td></td><td class="n">{won(g["open_amt"].sum())}</td>'
                        f'<td></td><td class="n">{won(g["in_amt"].sum())}</td><td></td><td class="n">{won(g["out_amt"].sum())}</td>'
                        f'<td></td><td class="n">{won(g["close_amt"].sum())}</td><td></td></tr>')
        rows.append(f'<tr class="total"><td colspan="3">합계</td><td></td><td class="n">{won(active["open_amt"].sum())}</td>'
                    f'<td></td><td class="n">{won(active["in_amt"].sum())}</td><td></td><td class="n">{won(active["out_amt"].sum())}</td>'
                    f'<td></td><td class="n">{won(active["close_amt"].sum())}</td><td></td></tr>')
        ledger_html = ('<table class="t"><thead><tr><th rowspan="2">분류</th><th rowspan="2">품목</th><th rowspan="2">단위</th>'
                       '<th colspan="2">기초</th><th colspan="2">입고(매입)</th><th colspan="2">출고(사용)</th><th colspan="2">기말</th>'
                       '<th rowspan="2">입력<br>확인</th></tr><tr><th>수량</th><th>금액</th><th>수량</th><th>금액</th>'
                       '<th>수량</th><th>금액</th><th>수량</th><th>금액</th></tr></thead><tbody>' + "".join(rows) + "</tbody></table>")

    # ---- 4. 고정비 ----
    if fixed.empty:
        fixed_html = '<div class="empty">이 달 고정비 내역이 없습니다.</div>'
    else:
        rows = [f'<tr><td>{esc(r["expense_item"])}</td><td class="n">{won(r["total_billed_amount"])}</td>{chk}</tr>' for _, r in fixed.iterrows()]
        rows.append(f'<tr class="total"><td>고정비 합계 ({len(fixed)}건)</td><td class="n">{won(fixed["total_billed_amount"].sum())}</td><td></td></tr>')
        fixed_html = '<table class="t"><thead><tr><th>지출 항목</th><th>금액(원)</th><th>입력<br>확인</th></tr></thead><tbody>' + "".join(rows) + "</tbody></table>"

    # ---- 5. 시험군별 원가 배분 ----
    if not settled:
        alloc_html = f'<div class="empty">{esc(month)} 은 아직 정산하지 않았습니다.</div>'
    else:
        rows = [f'<tr><td>{esc(r["grp"])}</td><td class="n">{int(r["heads"]):,}</td><td class="n">{won(r["var_cost"])}</td>'
                f'<td class="n">{won(r["fix_cost"])}</td><td class="n">{won(r["total"])}</td>{chk}</tr>' for _, r in alloc.iterrows()]
        rows.append(f'<tr class="total"><td>합계</td><td class="n">{int(alloc["heads"].sum()):,}</td><td class="n">{won(alloc["var_cost"].sum())}</td>'
                    f'<td class="n">{won(alloc["fix_cost"].sum())}</td><td class="n">{won(alloc["total"].sum())}</td><td></td></tr>')
        alloc_html = ('<table class="t"><thead><tr><th>시험군</th><th>두수</th><th>변동비(원)</th><th>고정비(원)</th><th>배분 합계(원)</th>'
                      '<th>입력<br>확인</th></tr></thead><tbody>' + "".join(rows) + "</tbody></table>")

    # ---- 6. 개체 증감 ----
    if admitted.empty:
        admit_html = '<div class="empty">이 달 입식한 개체가 없습니다.</div>'
        admit_sum_html = admit_html
    else:
        # 회계 입력은 거래(입식일·우시장) 단위로 하므로 먼저 거래별 합계를 보여 주고, 개체별 명세는 아래에 둔다.
        deals = (admitted.assign(market_name=admitted["market_name"].fillna("(우시장 없음)"))
                 .groupby(["admission_date", "market_name"], sort=True)
                 .agg(heads=("cattle_id", "count"), calf=("calf_price", "sum"), comm=("commission_fee", "sum"),
                      trans=("transport_fee", "sum"), total=("initial_cost", "sum")).reset_index())
        rows = [f'<tr><td class="c">{esc(r["admission_date"])}</td><td>{esc(r["market_name"])}</td><td class="n">{int(r["heads"]):,}</td>'
                f'<td class="n">{won(r["calf"])}</td><td class="n">{won(r["comm"])}</td><td class="n">{won(r["trans"])}</td>'
                f'<td class="n">{won(r["total"])}</td>{chk}</tr>' for _, r in deals.iterrows()]
        rows.append(f'<tr class="total"><td colspan="2">합계 ({len(deals)}건)</td><td class="n">{len(admitted):,}</td>'
                    f'<td class="n">{won(admitted["calf_price"].sum())}</td><td class="n">{won(admitted["commission_fee"].sum())}</td>'
                    f'<td class="n">{won(admitted["transport_fee"].sum())}</td><td class="n">{won(admitted["initial_cost"].sum())}</td><td></td></tr>')
        admit_sum_html = ('<table class="t"><thead><tr><th>입식일</th><th>우시장</th><th>두수</th><th>구입금액</th><th>수수료</th>'
                          '<th>운송료</th><th>구입비용합계</th><th>입력<br>확인</th></tr></thead><tbody>' + "".join(rows) + "</tbody></table>")
        rows = [f'<tr><td class="c">{esc(r["admission_date"])}</td><td class="c">{esc(r["cattle_id"])}</td><td>{esc(r["test_name"])}</td>'
                f'<td>{esc(r["market_name"])}</td><td class="n">{won(r["calf_price"])}</td><td class="n">{won(r["commission_fee"])}</td>'
                f'<td class="n">{won(r["transport_fee"])}</td><td class="n">{won(r["initial_cost"])}</td>{chk}</tr>' for _, r in admitted.iterrows()]
        rows.append(f'<tr class="total"><td colspan="4">입식 합계 ({len(admitted)}두)</td><td class="n">{won(admitted["calf_price"].sum())}</td>'
                    f'<td class="n">{won(admitted["commission_fee"].sum())}</td><td class="n">{won(admitted["transport_fee"].sum())}</td>'
                    f'<td class="n">{won(admitted["initial_cost"].sum())}</td><td></td></tr>')
        admit_html = ('<table class="t"><thead><tr><th>입식일</th><th>이표번호</th><th>시험군</th><th>우시장</th><th>구입금액</th>'
                      '<th>수수료</th><th>운송료</th><th>구입비용합계</th><th>입력<br>확인</th></tr></thead><tbody>' + "".join(rows) + "</tbody></table>")
    if closed.empty:
        closed_html = '<div class="empty">이 달 출하·폐사한 개체가 없습니다.</div>'
    else:
        rows = [f'<tr><td class="c">{esc(r["closure_date"])}</td><td class="c">{esc(r["status"])}</td><td class="c">{esc(r["cattle_id"])}</td>'
                f'<td>{esc(r["test_name"])}</td><td class="n">{won(r["initial_cost"])}</td><td class="n">{won(r["raised_cost"])}</td>'
                f'<td class="n">{won(r["book"])}</td>{chk}</tr>' for _, r in closed.iterrows()]
        rows.append(f'<tr class="total"><td colspan="4">합계 ({len(closed)}두)</td><td class="n">{won(closed["initial_cost"].sum())}</td>'
                    f'<td class="n">{won(closed["raised_cost"].sum())}</td><td class="n">{won(closed["book"].sum())}</td><td></td></tr>')
        closed_html = ('<table class="t"><thead><tr><th>종결일</th><th>구분</th><th>이표번호</th><th>시험군</th><th>구입원가</th>'
                       '<th>누적 사육비</th><th>누적 원가</th><th>입력<br>확인</th></tr></thead><tbody>' + "".join(rows) + "</tbody></table>")

    # ---- 7. 월말 사육 개체 장부가 ----
    if on_hand.empty:
        onhand_html = '<div class="empty">월말 사육 중인 개체가 없습니다.</div>'
    else:
        rows = [f'<tr><td>{esc(r["grp"])}</td><td class="n">{int(r["heads"]):,}</td><td class="n">{won(r["buy_cost"])}</td>'
                f'<td class="n">{won(r["raised_cost"])}</td><td class="n">{won(r["book"])}</td>{chk}</tr>' for _, r in on_hand.iterrows()]
        rows.append(f'<tr class="total"><td>합계</td><td class="n">{int(on_hand["heads"].sum()):,}</td><td class="n">{won(on_hand["buy_cost"].sum())}</td>'
                    f'<td class="n">{won(on_hand["raised_cost"].sum())}</td><td class="n">{won(on_hand["book"].sum())}</td><td></td></tr>')
        onhand_html = ('<table class="t"><thead><tr><th>시험군</th><th>두수</th><th>구입원가 합계</th><th>누적 사육비</th>'
                       '<th>장부가(원)</th><th>입력<br>확인</th></tr></thead><tbody>' + "".join(rows) + "</tbody></table>")

    alloc_note = ""
    if settled:
        diff_var = alloc["var_cost"].sum() - active["out_amt"].sum()
        diff_fix = alloc["fix_cost"].sum() - fixed["total_billed_amount"].sum()
        parts = [f"{name} {round(v):+,}원" for name, v in (("변동비", diff_var), ("고정비", diff_fix)) if abs(v) >= 0.5]
        if parts:
            alloc_note = (f'<p class="note">※ 등록 금액과의 차이: {", ".join(parts)} — '
                          '개체별로 나눌 때 원 미만을 반올림해서 생기는 차이입니다.</p>')

    journal_html = build_journal_html(fixed, active["out_amt"].sum(), esc, won, chk)

    made_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    body = f"""
<div class="toolbar"><button onclick="window.print()">🖨️ 인쇄 / PDF 저장</button></div>
<div class="sheet">
  <div class="head">
    <div>
      <h1>월말 팔공 입력 자료</h1>
      <div class="meta">농장: <b>{esc(farm_name)}</b> &nbsp;|&nbsp; 대상 월: <b>{esc(month)}</b><br>출력일시: {esc(made_at)}</div>
    </div>
    <table class="sign"><tr><th>작성</th><th>검토</th><th>승인</th></tr><tr><td></td><td></td><td></td></tr></table>
  </div>
  <h2>1. 요약</h2>{summary_html}
  <h2>2. 회계 전표 (원가배부)</h2><p class="note">회계 프로그램 전표 입력용 · 차변 합계와 대변 합계가 다르면 빨간 글씨로 표시됩니다</p>{journal_html}
  <h2>3. 매입(입고) 내역</h2><p class="note">매입일자가 {esc(month)} 인 매입 · 분류별 소계</p>{purchase_html}
  <h2>4. 품목 수불부</h2><p class="note">기초 = 전월까지 매입 − 전월까지 사용 · 출고 = 이 달 월말 비용 등록(시험군별 사용량) · 금액은 산출액 기준</p>{ledger_html}
  <div class="keep"><h2>5. 고정비</h2>{fixed_html}</div>
  <div class="keep"><h2>6. 시험군별 원가 배분 (월말 정산 결과)</h2>{alloc_html}{alloc_note}</div>
  <div class="keep"><h2>7-1. 입식 (입식일 · 우시장별 합계)</h2>{admit_sum_html}</div>
  <div class="keep"><h2>7-2. 출하 · 폐사 개체</h2><p class="note">누적 원가 = 구입원가 + 이 달까지 배분된 사육비</p>{closed_html}</div>
  <div class="keep"><h2>8. 월말 사육 개체 장부가</h2><p class="note">{esc(month)} 말 기준 사육 중인 개체 · 장부가 = 구입원가 + 이 달까지 배분된 사육비</p>{onhand_html}</div>
  <h2>[첨부] 입식 개체 명세</h2><p class="note">7-1 합계의 개체별 내역 (증빙용)</p>{admit_html}
</div>"""
    page = (f'<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
            f'<title>월말 팔공 입력 자료 {esc(farm_name)} {esc(month)}</title><style>{ACCOUNTING_SHEET_CSS}</style></head><body>{body}</body></html>')
    return True, page


# ========== 재고조사표 (월말 실사용, 인쇄용) ==========
# 회사에서 쓰던 엑셀 양식(재고조사표)과 같은 칸 구성:
#   A 전 재고조사일 기준 재고량 / B 그 뒤 매입량 / C = A + B 원가배부 전 장부상 재고량 /
#   D 기준일 실 재고량 / E = C - D 급여량(원가배부량)
# 첫 줄은 '한우위탁우'(두수): A 전월말 사육두수, B 이 달 입식, D 이 달 말 사육두수, E 출하·폐사.
STOCK_COUNT_CATTLE_LABEL = "한우위탁우"


def _month_bounds(month):
    """'2026-09' → ('2026-09-01', '2026-09-30', '2026-08-31')"""
    import calendar
    y, m = int(month[:4]), int(month[5:7])
    last = calendar.monthrange(y, m)[1]
    py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
    prev_last = calendar.monthrange(py, pm)[1]
    return f"{month}-01", f"{month}-{last:02d}", f"{py:04d}-{pm:02d}-{prev_last:02d}"


def stock_count_rows(db_file, month):
    """재고조사표 줄 목록. D(실 재고량)는 장부상 값(C - 이 달 등록 사용량)으로 미리 채워 둔다.
    반환: DataFrame [코드, 상품명, 규격, A, B, C, 등록사용량, D]"""
    start, end, prev_end = _month_bounds(month)
    conn = db_connect(db_file)
    try:
        alive = "(closure_date IS NULL OR closure_date = '' OR closure_date > ?)"
        head_a = conn.execute(f"SELECT COUNT(*) FROM cattle WHERE COALESCE(admission_date, '') <= ? AND {alive}",
                              (prev_end, prev_end)).fetchone()[0]
        head_b = conn.execute("SELECT COUNT(*) FROM cattle WHERE admission_date >= ? AND admission_date <= ?",
                              (start, end)).fetchone()[0]
        head_d = conn.execute(f"SELECT COUNT(*) FROM cattle WHERE COALESCE(admission_date, '') <= ? AND {alive}",
                              (end, end)).fetchone()[0]
        items = pd.read_sql("""
            SELECT i.item_code, i.item_name, i.unit,
              COALESCE((SELECT SUM(quantity) FROM purchase p WHERE p.item_code = i.item_code AND p.purchase_date < ?), 0)
            - COALESCE((SELECT SUM(total_usage) FROM monthly_usage u WHERE u.item_code = i.item_code AND u.settlement_month < ?), 0) AS a,
              COALESCE((SELECT SUM(quantity) FROM purchase p WHERE p.item_code = i.item_code AND p.purchase_date >= ? AND p.purchase_date <= ?), 0) AS b,
              COALESCE((SELECT SUM(total_usage) FROM monthly_usage u WHERE u.item_code = i.item_code AND u.settlement_month = ?), 0) AS used
            FROM item_master i
        """, conn, params=(start, month, start, end, month))
    finally:
        conn.close()

    rows = [{"코드": "_cattle", "상품명": STOCK_COUNT_CATTLE_LABEL, "규격": "두",
             "A": float(head_a), "B": float(head_b), "C": float(head_a + head_b),
             "등록사용량": float(head_a + head_b - head_d), "D": float(head_d)}]
    for c in ("a", "b", "used"):
        items[c] = pd.to_numeric(items[c], errors="coerce").fillna(0).astype(float)
    items = items[(items["a"].abs() > 0.005) | (items["b"].abs() > 0.005) | (items["used"].abs() > 0.005)].copy()
    # 등록한 순서(품목코드 번호순)대로 — 회사 양식과 같은 순서
    items["_n"] = items["item_code"].map(lambda c: int("".join(filter(str.isdigit, str(c))) or 0))
    for r in items.sort_values(["_n", "item_code"]).itertuples(index=False):
        c = r.a + r.b
        rows.append({"코드": r.item_code, "상품명": r.item_name, "규격": r.unit or "",
                     "A": r.a, "B": r.b, "C": c, "등록사용량": r.used, "D": c - r.used})
    return pd.DataFrame(rows)


def _stock_qty(v):
    """0 은 '-', 정수는 콤마, 소수는 둘째 자리까지."""
    if v is None or pd.isna(v) or abs(float(v)) < 0.005:
        return "-"
    return f"{float(v):,.2f}".rstrip("0").rstrip(".")


STOCK_COUNT_CSS = """
@page { size: A4 portrait; margin: 12mm 10mm; }
* { box-sizing: border-box; }
body { font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; color: #111; margin: 0; background: #f3f4f6; font-size: 13px; }
.sheet { max-width: 900px; margin: 16px auto; background: #fff; padding: 24px 26px; box-shadow: 0 2px 10px rgba(0,0,0,.08); }
h1 { text-align: center; font-size: 24px; letter-spacing: 1px; margin: 0 0 6px; font-weight: 700; }
.date { text-align: center; margin: 0 0 14px; }
.sign { text-align: right; line-height: 2; margin-bottom: 6px; font-size: 14px; }
table { width: 100%; border-collapse: collapse; }
th, td { border: 1px solid #111; padding: 5px 6px; }
th { background: #f2f2f2; font-weight: 600; text-align: center; line-height: 1.35; }
td.c { text-align: center; }
td.n { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
thead { display: table-header-group; }
tr { page-break-inside: avoid; }
.toolbar { max-width: 900px; margin: 12px auto 0; text-align: right; }
.toolbar button { font-size: 14px; padding: 8px 18px; border: 0; border-radius: 8px; background: #2E6B57; color: #fff; cursor: pointer; }
@media print {
  body { background: #fff; }
  .sheet { box-shadow: none; margin: 0; max-width: none; padding: 0; }
  .toolbar { display: none; }
}
"""


def _stock_count_titles(month, farm_label):
    _, end, _ = _month_bounds(month)
    y, m, d = end.split("-")
    return (f"{y}년 {m}월 재고조사표({farm_label})", f"{y}.{m}.{d}.", f"{m}월 {d}일 기준")


def generate_stock_count_sheet(rows, month, farm_label, examiner, witness):
    """재고조사표 인쇄용 HTML. rows 는 stock_count_rows() 결과에 D(실 재고량)를 확정한 것."""
    esc = lambda v: html.escape("" if v is None else str(v))
    title, base_date, day_label = _stock_count_titles(month, farm_label)
    body_rows = []
    for i, r in enumerate(rows.itertuples(index=False), start=1):
        body_rows.append(
            f'<tr><td class="c">{i}</td><td>{esc(r.상품명)}</td><td class="c">{esc(r.규격)}</td>'
            f'<td class="n">{_stock_qty(r.A)}</td><td class="n">{_stock_qty(r.B)}</td><td class="n">{_stock_qty(r.C)}</td>'
            f'<td class="n">{_stock_qty(r.D)}</td><td class="n">{_stock_qty(r.C - r.D)}</td><td></td></tr>')
    return f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>{esc(title)}</title>
<style>{STOCK_COUNT_CSS}</style></head><body>
<div class="toolbar"><button onclick="window.print()">🖨️ 인쇄 / PDF 저장</button></div>
<div class="sheet">
<h1>{esc(title)}</h1>
<p class="date">(기준일 : {base_date})</p>
<div class="sign">조사자 : {esc(examiner)} &nbsp;(인)<br>입회자 : {esc(witness)} &nbsp;(인)</div>
<table><thead><tr>
<th style="width:5%">순번</th><th>상품명</th><th style="width:6%">규격</th>
<th style="width:12%">전 재고조사일<br>기준 재고량<br>(A)</th>
<th style="width:12%">전 재고조사일<br>이후 매입량<br>(B)</th>
<th style="width:14%">{esc(day_label)}<br>원가배부 전<br>장부상 재고량<br>(C = A + B)</th>
<th style="width:12%">{esc(day_label)}<br>실 재고량<br>(D)</th>
<th style="width:12%">급여량<br>(원가배부량)<br>(E = C − D)</th>
<th style="width:7%">비고</th>
</tr></thead><tbody>{''.join(body_rows)}</tbody></table>
</div></body></html>"""


def stock_count_excel(rows, month, farm_label, examiner, witness):
    """재고조사표 엑셀(.xlsx) — 회사 양식과 같은 칸 배치."""
    title, base_date, day_label = _stock_count_titles(month, farm_label)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        wb = xw.book
        ws = wb.add_worksheet("재고조사표")
        f_title = wb.add_format({"bold": True, "font_size": 18, "align": "center", "valign": "vcenter"})
        f_center = wb.add_format({"align": "center"})
        f_right = wb.add_format({"align": "right"})
        f_head = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "text_wrap": True,
                                "border": 1, "bg_color": "#F2F2F2"})
        f_txt = wb.add_format({"border": 1})
        f_ctr = wb.add_format({"border": 1, "align": "center"})
        f_num = wb.add_format({"border": 1, "num_format": '#,##0.##;-#,##0.##;"-"'})
        ws.set_column(0, 0, 6); ws.set_column(1, 1, 20); ws.set_column(2, 2, 7)
        ws.set_column(3, 7, 15); ws.set_column(8, 8, 8)
        ws.merge_range(0, 0, 0, 8, title, f_title)
        ws.set_row(0, 30)
        ws.merge_range(1, 0, 1, 8, f"(기준일 : {base_date})", f_center)
        ws.merge_range(2, 0, 2, 8, f"조사자 : {examiner}  (인)", f_right)
        ws.merge_range(3, 0, 3, 8, f"입회자 : {witness}  (인)", f_right)
        heads = ["순번", "상품명", "규격", "전 재고조사일\n기준 재고량\n(A)", "전 재고조사일\n이후 매입량\n(B)",
                 f"{day_label}\n원가배부 전\n장부상 재고량\n(C = A + B)", f"{day_label}\n실 재고량\n(D)",
                 "급여량\n(원가배부량)\n(E = C - D)", "비고"]
        ws.set_row(4, 66)
        for c, h in enumerate(heads):
            ws.write(4, c, h, f_head)
        for i, r in enumerate(rows.itertuples(index=False), start=1):
            row = 4 + i
            ws.write(row, 0, i, f_ctr)
            ws.write(row, 1, r.상품명, f_txt)
            ws.write(row, 2, r.규격, f_ctr)
            for c, v in zip(range(3, 8), (r.A, r.B, r.C, r.D, r.C - r.D)):
                ws.write_number(row, c, round(float(v), 2), f_num)
            ws.write(row, 8, "", f_txt)
        ws.fit_to_pages(1, 0)
        ws.set_paper(9)  # A4
    return buf.getvalue()


# ========== 원가배부 내역 (시험군별 사용량 · 재고 대사, 인쇄용) ==========
# 회사에서 쓰던 엑셀 양식(원가배부 내역)과 같은 칸 구성:
#   시험군별 사용량 [A][B].. / F 전 재고조사일 이후 사용량(시험군 합계) / G 전 재고조사일 기준 재고량 /
#   H 이후 매입량 / I = G + H 장부상 재고량 / J 실 재고량 / K = I - J 원가배부량 / L = K - F 차액분
def cost_allocation_rows(db_file, month):
    """반환: (시험군 목록 [(코드, 이름)], DataFrame[코드, 품목, 규격, g0.., F, G, H, I])
    시험군 사용량 = 월말 비용 등록(monthly_usage) 그 달 값."""
    start, end, _ = _month_bounds(month)
    conn = db_connect(db_file)
    try:
        usage = pd.read_sql(
            "SELECT item_code, test_group_code, SUM(total_usage) AS q FROM monthly_usage "
            "WHERE settlement_month = ? GROUP BY item_code, test_group_code", conn, params=(month,))
        groups = pd.read_sql("SELECT test_group_code, test_name FROM testgroup_master ORDER BY test_group_code", conn)
        items = pd.read_sql("""
            SELECT i.item_code, i.item_name, i.unit,
              COALESCE((SELECT SUM(quantity) FROM purchase p WHERE p.item_code = i.item_code AND p.purchase_date < ?), 0)
            - COALESCE((SELECT SUM(total_usage) FROM monthly_usage u WHERE u.item_code = i.item_code AND u.settlement_month < ?), 0) AS g,
              COALESCE((SELECT SUM(quantity) FROM purchase p WHERE p.item_code = i.item_code AND p.purchase_date >= ? AND p.purchase_date <= ?), 0) AS h
            FROM item_master i
        """, conn, params=(start, month, start, end))
    finally:
        conn.close()

    # 그 달 사용량이 있는 시험군만 칸으로 둔다 (없으면 등록된 시험군 전체).
    used_codes = set(usage["test_group_code"].dropna())
    grp = [(r.test_group_code, r.test_name) for r in groups.itertuples(index=False)
           if not used_codes or r.test_group_code in used_codes]
    usage["q"] = pd.to_numeric(usage["q"], errors="coerce").fillna(0).astype(float)
    by = {(r.item_code, r.test_group_code): r.q for r in usage.itertuples(index=False)}

    rows = []
    for c in ("g", "h"):
        items[c] = pd.to_numeric(items[c], errors="coerce").fillna(0).astype(float)
    items["_n"] = items["item_code"].map(lambda c: int("".join(filter(str.isdigit, str(c))) or 0))
    for r in items.sort_values(["_n", "item_code"]).itertuples(index=False):
        per = [by.get((r.item_code, code), 0.0) for code, _ in grp]
        f = float(sum(by.get((r.item_code, code), 0.0) for code in used_codes)) if used_codes else 0.0
        if max(abs(r.g), abs(r.h), abs(f)) < 0.005:
            continue
        row = {"코드": r.item_code, "품목": r.item_name, "규격": r.unit or ""}
        row.update({f"g{i}": v for i, v in enumerate(per)})
        row.update({"F": f, "G": r.g, "H": r.h, "I": r.g + r.h})
        rows.append(row)
    cols = ["코드", "품목", "규격", *[f"g{i}" for i in range(len(grp))], "F", "G", "H", "I"]
    return grp, pd.DataFrame(rows, columns=cols)


def _cost_alloc_headers(groups, month):
    """시험군 칸 글자(A, B, ..)와 나머지 칸 글자. 시험군이 5개 이하면 회사 양식처럼 F~L."""
    import string
    letters = string.ascii_uppercase
    g_letters = [letters[i] for i in range(len(groups))]
    base = max(len(groups), 5)
    F, G, H, I, J, K, L = (letters[base + i] for i in range(7))
    f_formula = "+".join(g_letters) if len(g_letters) <= 3 else f"{g_letters[0]}~{g_letters[-1]}"
    _, end, _ = _month_bounds(month)
    y, m, d = end.split("-")
    day = f"{m}월 {d}일"
    heads = [f"{name}<br>[{l}]" for (_, name), l in zip(groups, g_letters)] + [
        f"전 재고조사일<br>이후 사용량<br>({F}={f_formula or '0'})",
        f"전 재고조사일<br>기준 재고량<br>({G})",
        f"전 재고조사일<br>이후 매입량<br>({H})",
        f"{day} 기준<br>원가배부전<br>장부상 재고량<br>({I}={G}+{H})",
        f"{day}<br>기준<br>실 재고량<br>({J})",
        f"원가배부량<br>{K}=({I}-{J})",
        f"차액분<br>{L}=({K}-{F})",
    ]
    return heads, f"{y}.{m}.{d}", f"{y}년 {m}월"


def _cost_alloc_values(rows, n_groups, actual):
    """줄마다 [시험군들.., F, G, H, I, J, K, L]. actual: {품목코드: 실 재고량} (없으면 I - F)."""
    out = []
    for r in rows.to_dict("records"):
        j = actual.get(r["코드"]) if actual else None
        j = r["I"] - r["F"] if j is None or pd.isna(j) else float(j)
        k = r["I"] - j
        out.append([r[f"g{i}"] for i in range(n_groups)] + [r["F"], r["G"], r["H"], r["I"], j, k, k - r["F"]])
    return out


def generate_cost_allocation_sheet(groups, rows, month, farm_label, actual, diff_note, writers):
    esc = lambda v: html.escape("" if v is None else str(v))
    heads, base_date, ym = _cost_alloc_headers(groups, month)
    title = f"{ym} {farm_label} 원가배부 내역"
    values = _cost_alloc_values(rows, len(groups), actual)
    body = []
    for r, vals in zip(rows.to_dict("records"), values):
        body.append(f'<tr><td>{esc(r["품목"])}</td><td class="c">{esc(r["규격"])}</td>'
                    + "".join(f'<td class="n">{_stock_qty(v)}</td>' for v in vals) + "</tr>")
    writer_text = ", ".join(f"{esc(w.strip())}(인)" for w in str(writers).split(",") if w.strip())
    css = STOCK_COUNT_CSS.replace("size: A4 portrait", "size: A4 landscape").replace("max-width: 900px", "max-width: 1100px")
    return f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>{esc(title)}</title>
<style>{css}
.foot {{ padding: 7px 2px; }}
th {{ font-size: 12px; }}
</style></head><body>
<div class="toolbar"><button onclick="window.print()">🖨️ 인쇄 / PDF 저장</button></div>
<div class="sheet">
<h1>{esc(title)}</h1>
<p class="date">(재고조사일 : {base_date})</p>
<table><thead><tr><th style="width:14%">구 분</th><th style="width:5%">규격</th>{''.join(f'<th>{h}</th>' for h in heads)}</tr></thead>
<tbody>{''.join(body)}</tbody></table>
<div class="foot">▣ 차액분 : {esc(diff_note)}</div>
<div class="foot">▣ 작성자 : {writer_text}</div>
</div></body></html>"""


def cost_allocation_excel(groups, rows, month, farm_label, actual, diff_note, writers):
    heads, base_date, ym = _cost_alloc_headers(groups, month)
    heads = [h.replace("<br>", "\n") for h in heads]
    title = f"{ym} {farm_label} 원가배부 내역"
    values = _cost_alloc_values(rows, len(groups), actual)
    ncol = 2 + len(heads)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        wb = xw.book
        ws = wb.add_worksheet("원가배부내역")
        f_title = wb.add_format({"bold": True, "font_size": 18, "align": "center", "valign": "vcenter"})
        f_center = wb.add_format({"align": "center"})
        f_head = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "text_wrap": True,
                                "border": 1, "bg_color": "#F2F2F2"})
        f_txt = wb.add_format({"border": 1})
        f_ctr = wb.add_format({"border": 1, "align": "center"})
        f_num = wb.add_format({"border": 1, "num_format": '#,##0.##;-#,##0.##;"-"'})
        ws.set_column(0, 0, 16); ws.set_column(1, 1, 6); ws.set_column(2, ncol - 1, 13)
        ws.merge_range(0, 0, 0, ncol - 1, title, f_title)
        ws.set_row(0, 30)
        ws.merge_range(1, 0, 1, ncol - 1, f"(재고조사일 : {base_date})", f_center)
        ws.set_row(2, 80)
        for c, h in enumerate(["구 분", "규격", *heads]):
            ws.write(2, c, h, f_head)
        for i, (r, vals) in enumerate(zip(rows.to_dict("records"), values), start=3):
            ws.write(i, 0, r["품목"], f_txt)
            ws.write(i, 1, r["규격"], f_ctr)
            for c, v in enumerate(vals, start=2):
                ws.write_number(i, c, round(float(v), 2), f_num)
        last = 3 + len(values)
        writer_text = ", ".join(f"{w.strip()}(인)" for w in str(writers).split(",") if w.strip())
        f_plain = wb.add_format({"valign": "vcenter"})
        ws.merge_range(last, 0, last, ncol - 1, f"▣ 차액분 : {diff_note}", f_plain)
        ws.merge_range(last + 1, 0, last + 1, ncol - 1, f"▣ 작성자 : {writer_text}", f_plain)
        ws.set_landscape(); ws.set_paper(9); ws.fit_to_pages(1, 0)
    return buf.getvalue()


def list_accounting_months(db_file):
    """매입·사용량·고정비·정산·입식·종결 중 하나라도 있는 연월 (최신순)."""
    conn = db_connect(db_file)
    try:
        rows = conn.execute("""
            SELECT substr(purchase_date, 1, 7) FROM purchase
            UNION SELECT settlement_month FROM monthly_usage
            UNION SELECT settlement_month FROM monthly_fixedcost
            UNION SELECT settlement_month FROM cattle_cost_log
            UNION SELECT substr(admission_date, 1, 7) FROM cattle
            UNION SELECT substr(closure_date, 1, 7) FROM cattle
        """).fetchall()
    finally:
        conn.close()
    return sorted({r[0] for r in rows if r[0] and re.fullmatch(r"\d{4}-\d{2}", str(r[0]))}, reverse=True)


def generate_overall_report_html(df_all, farm_order, farm_colors):
    """전체 현황 통합 보고서(요약 카드 + Chart.js 차트 + 표)를 단독 실행 가능한 HTML로 만든다."""
    esc = html.escape

    def num(n):
        return f"{int(n):,}"

    def man(won):
        return int(won) // 10000

    def avg_won(series):
        return series.mean(skipna=True) if series.notna().any() else 0

    df = df_all.copy()
    df['초기원가'] = pd.to_numeric(df['초기원가'], errors='coerce')
    present = set(df['농장명'])
    farms = [f for f in farm_order if f in present]

    total_cnt = len(df)
    breeding_cnt = int((df['상태'] == '사육').sum())
    dead_cnt = int((df['상태'] == '폐사').sum())
    shipped_cnt = int((df['상태'] == '출하').sum())
    total_cost = df['초기원가'].sum(skipna=True)

    def pct(part):
        return f"{part / total_cnt * 100:.1f}%" if total_cnt else "0.0%"

    kpis = [
        ("total", "전체 누적 입식", num(total_cnt), "두", "등록된 전체 개체"),
        ("breeding", "현재 사육중", num(breeding_cnt), "두", f"전체의 {pct(breeding_cnt)}"),
        ("dead", "누적 폐사", num(dead_cnt), "두", f"폐사율 {pct(dead_cnt)}"),
        ("shipped", "누적 출하", num(shipped_cnt), "두", f"출하율 {pct(shipped_cnt)}"),
        ("cost", "총 구입비용", num(man(total_cost)), "만원", f"두당 평균 {num(man(avg_won(df['초기원가'])))}만원"),
    ]
    kpi_html = "".join(
        f'<div class="kpi {cls}"><div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}<small>{unit}</small></div><div class="kpi-sub">{sub}</div></div>'
        for cls, label, value, unit, sub in kpis
    )

    # 2. 농장별 요약
    def mort_rate(dead, total):
        return f"{dead / total * 100:.1f}%" if total else "-"

    farm_rows, chart_farms = [], []
    for f in farms:
        d = df[df['농장명'] == f]
        cost = d['초기원가'].sum(skipna=True)
        chart_farms.append({"name": f, "total": len(d), "cost_man": man(cost)})
        farm_rows.append(
            f'<tr><td class="left"><span class="dot" style="background:{esc(farm_colors.get(f, "#4F46E5"))}"></span>{esc(f)}</td>'
            f'<td>{num(len(d))}</td><td>{num((d["상태"] == "사육").sum())}</td>'
            f'<td>{num((d["상태"] == "출하").sum())}</td><td>{num((d["상태"] == "폐사").sum())}</td>'
            f'<td>{mort_rate(int((d["상태"] == "폐사").sum()), len(d))}</td>'
            f'<td>{num(man(cost))}</td><td>{num(man(avg_won(d["초기원가"])))}</td></tr>'
        )
    farm_table = (
        '<div class="table-wrap"><table class="data-table"><thead><tr>'
        '<th class="left">농장명</th><th>전체 입식 (두)</th><th>현재 사육중 (두)</th>'
        '<th>누적 출하 (두)</th><th>누적 폐사 (두)</th><th>폐사율</th><th>총 구입비용 (만원)</th><th>두당 평균 (만원)</th>'
        '</tr></thead><tbody>' + "".join(farm_rows) + '</tbody>'
        f'<tfoot><tr><td class="left">합계</td><td>{num(total_cnt)}</td><td>{num(breeding_cnt)}</td>'
        f'<td>{num(shipped_cnt)}</td><td>{num(dead_cnt)}</td><td>{mort_rate(dead_cnt, total_cnt)}</td><td>{num(man(total_cost))}</td>'
        f'<td>{num(man(avg_won(df["초기원가"])))}</td></tr></tfoot></table></div>'
    )

    # 3. 농장별 우시장 구입 현황
    dm = df[df['우시장'].notna() & (df['우시장'].astype(str).str.strip() != '')].copy()
    dm['우시장'] = dm['우시장'].astype(str).str.strip()
    chart_markets, donut_items, market_sections = [], [], []
    idx = 0
    for f in farms:
        d = dm[dm['농장명'] == f]
        if d.empty:
            continue
        idx += 1
        grp = (
            d.groupby('우시장')
            .agg(cnt=('개체번호', 'count'), cost=('초기원가', lambda x: x.sum(skipna=True)),
                 avg=('초기원가', avg_won))
            .reset_index()
            .sort_values(['cnt', '우시장'], ascending=[False, True])
        )
        farm_cnt = int(grp['cnt'].sum())
        chart_markets.append({
            "farm": f,
            "items": [{"market": r['우시장'], "count": int(r['cnt'])} for _, r in grp.iterrows()],
        })
        donut_items.append(
            f'<div class="donut-item"><div class="chart-box"><canvas id="donut-{idx - 1}"></canvas></div>'
            f'<p>{idx}) {esc(f)}</p></div>'
        )
        rows = "".join(
            f'<tr><td class="left">{esc(r["우시장"])}</td><td>{num(r["cnt"])}</td>'
            f'<td>{r["cnt"] / farm_cnt * 100:.1f}%</td><td>{num(man(r["cost"]))}</td>'
            f'<td>{num(man(r["avg"]))}</td></tr>'
            for _, r in grp.iterrows()
        )
        market_sections.append(
            f'<div class="farm-sub"><span class="dot" style="background:{esc(farm_colors.get(f, "#4F46E5"))}"></span>'
            f'{idx}) {esc(f)}</div>'
            '<div class="table-wrap"><table class="data-table"><thead><tr>'
            '<th class="left">우시장</th><th>구입 마릿수 (두)</th><th>비중</th>'
            '<th>총 구입비용 (만원)</th><th>두당 평균 (만원)</th></tr></thead>'
            f'<tbody>{rows}</tbody>'
            f'<tfoot><tr><td class="left">소계</td><td>{num(farm_cnt)}</td><td>100.0%</td>'
            f'<td>{num(man(d["초기원가"].sum(skipna=True)))}</td><td>{num(man(avg_won(d["초기원가"])))}</td></tr></tfoot>'
            '</table></div>'
        )

    if market_sections:
        market_html = (
            '<div class="card"><div class="card-head"><h3>농장별 우시장 구입 마릿수 비중</h3>'
            '<span>도넛 위에 마우스를 올리면 마릿수와 비중이 표시됩니다</span></div>'
            f'<div class="donut-grid">{"".join(donut_items)}</div></div>'
            + "".join(market_sections)
        )
    else:
        market_html = '<div class="empty">등록된 우시장 구입 이력이 없습니다.</div>'

    chart_data = json.dumps({"farms": chart_farms, "markets": chart_markets}, ensure_ascii=False)
    # <script> 안에 넣는 JSON 이므로 '</script>' 로 조기 종료되지 않게 '<' 를 이스케이프한다.
    chart_data = chart_data.replace("<", "\\u003c")

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>대구축협 시험농장 현황 보고</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>{OVERALL_REPORT_CSS}</style>
</head>
<body>
<div class="container">
  <header class="report-header">
    <div class="brand">
      <img src="{MEDAL_ICON_DATA_URI}" alt="심볼">
      <h1>대구축협 시험농장 현황 보고</h1>
    </div>
    <div class="date-badge">기준일: {datetime.now().strftime('%Y년 %m월 %d일')}</div>
  </header>

  <section class="section">
    <h2 class="section-title"><span class="no">1</span>전체 요약 현황</h2>
    <div class="kpi-grid">{kpi_html}</div>
  </section>

  <section class="section">
    <h2 class="section-title"><span class="no">2</span>농장별 요약 현황</h2>
    <div class="card">
      <div class="card-head"><h3>농장별 전체 입식 두수 · 총 구입비용</h3><span>왼쪽 축: 두수 / 오른쪽 축: 만원</span></div>
      <div class="chart-box"><canvas id="farmBarChart"></canvas></div>
    </div>
    {farm_table}
  </section>

  <section class="section">
    <h2 class="section-title"><span class="no">3</span>농장별 우시장 구입 현황</h2>
    {market_html}
  </section>

  <div class="footer">본 문서는 '대구축협 시험농장 관리 시스템'에 의해 자동 생성되었습니다.</div>
  <div class="print-wrap no-print"><button class="print-btn" onclick="window.print()">🖨️ PDF 인쇄 및 저장</button></div>
</div>
<script id="report-data" type="application/json">{chart_data}</script>
<script>{OVERALL_REPORT_JS}</script>
</body>
</html>"""


# ========== UI 메인 ==========

# 이전 실행(rerun)에서 닫히지 않은 연결부터 정리한다. -> "database is locked" 방지
close_stale_connections()
show_pending_toasts()

# 외부에서 접속 가능한 환경이면 비밀번호를 먼저 확인한다.
require_password()

# 사이드바: 농장 선택
st.sidebar.markdown(
    f"""
    <div style="display:flex; align-items:center; justify-content:center; gap:10px; padding: 4px 0 8px 0;">
        <img src="{MEDAL_ICON_DATA_URI}" alt="심볼" style="height: 40px;">
        <span style="font-size: 2.5rem;">🐂</span>
        <h2 style="margin:0; font-size:1.1rem; text-align:left; white-space:nowrap; word-break:keep-all;">대구축협 시험농장<br>관리 시스템</h2>
    </div>
    """,
    unsafe_allow_html=True,
)
st.sidebar.markdown("---")

farm_names = ["시험농장 전체 현황"] + list(FARM_CONFIG.keys())
selected_farm = st.sidebar.radio(
    "메뉴 및 농장 선택",
    farm_names,
    index=0,
    help="전체 현황 대시보드를 보거나, 정산을 수행할 농장을 선택하세요."
)

# 농장마다 품목·재고·매입 데이터는 따로(스키마별로) 저장되지만, 화면의 입력 상태(입력표에 넣던 값,
# 편집표의 저장 안 한 수정, 정산연월 등)는 세션 하나에 같은 키로 남는다. 그대로 두면 선산에서 입력하던
# 매입 줄이나 품목표 수정이 고아 화면에 따라와 고아 데이터에 저장될 수 있으므로, 농장을 바꾸면 비운다.
_SESSION_KEYS_KEPT_ACROSS_FARMS = {"_authed", "_open_db_conns", "_pending_toasts", "_active_farm"}
if st.session_state.get("_active_farm") != selected_farm:
    if "_active_farm" in st.session_state:
        for _k in [k for k in st.session_state.keys() if k not in _SESSION_KEYS_KEPT_ACROSS_FARMS]:
            del st.session_state[_k]
    st.session_state["_active_farm"] = selected_farm

with st.sidebar.expander("➕ 새 농장 추가"):
    with st.form("add_farm_form", clear_on_submit=True):
        new_farm_name = st.text_input("새 농장 이름")
        new_farm_color = st.color_picker("테마 색상", "#3B82F6")
        new_farm_b = st.number_input("동 개수 (예: 6)", min_value=1, max_value=20, value=6)
        new_farm_p = st.number_input("동별 우방 개수 (예: 20)", min_value=1, max_value=100, value=20)
        if st.form_submit_button("농장 추가", type="primary", width="stretch"):
            if new_farm_name:
                if new_farm_name in FARM_CONFIG:
                    st.error("이미 존재하는 농장입니다.")
                else:
                    import re
                    safe_name = re.sub(r'[\\/*?:"<>|]', "", new_farm_name)
                    new_db_filename = f"erp_{safe_name}.db"
                    FARM_CONFIG[new_farm_name] = {
                        "db_file": new_db_filename,
                        "color": new_farm_color,
                        "buildings_count": new_farm_b,
                        "pens_count": new_farm_p
                    }
                    save_farms(FARM_CONFIG)
                    notify(f"'{new_farm_name}' 농장이 추가되었습니다!", icon="✅")
                    st.rerun()
            else:
                st.warning("농장 이름을 입력하세요.")


def cattle_reset_preview(db_file):
    """개체 전체 삭제로 함께 지워질 데이터 건수 (cattle 및 cattle을 참조하는 표)."""
    tables = [
        ("개체", "cattle"),
        ("질병·처방 기록", "disease_record"),
        ("개체별 원가 내역", "cattle_cost_log"),
        ("개체별 품목 사용 내역", "cattle_item_usage_log"),
    ]
    rows = []
    if not db_exists(db_file):
        return rows
    conn = db_connect(db_file)
    try:
        for label, table in tables:
            try:
                count = conn.execute("SELECT COUNT(*) FROM %s" % table).fetchone()[0]
            except sqlite3.Error:
                continue
            if count:
                rows.append((label, count))
    finally:
        conn.close()
    return rows


@st.dialog("⚠️ 등록된 개체를 전부 삭제하시겠습니까?")
def cattle_reset_dialog():
    farm_list = list(FARM_CONFIG.keys())
    default_idx = farm_list.index(selected_farm) if selected_farm in farm_list else 0
    target_farm = st.selectbox("삭제할 농장 선택", farm_list, index=default_idx, key="cattle_reset_farm_sel")
    db_file = FARM_CONFIG[target_farm]["db_file"]

    st.error(f"**{target_farm}**에 등록된 모든 개체(입식 내역)가 삭제됩니다. 이 작업은 되돌릴 수 없습니다.")
    rows = cattle_reset_preview(db_file)
    if rows:
        st.markdown("**지워지는 데이터**")
        st.markdown("\n".join("- %s **%s건**" % (label, format(count, ",")) for label, count in rows))
    else:
        st.caption("현재 등록된 개체가 없습니다.")
    st.caption("시험군·품목·매입 내역 등은 그대로 남습니다. 실행 직전 자동으로 백업본을 만들기 때문에, 사이드바의 '백업 파일로 복원'으로 되돌릴 수 있습니다.")

    col_cancel, col_run = st.columns(2)
    if col_cancel.button("취소", width="stretch", key="cattle_reset_cancel"):
        st.rerun()
    if col_run.button(f"네, {target_farm} 개체를 삭제합니다", type="primary", width="stretch", key="cattle_reset_run"):
        backup_db(db_file, "before-cattle-reset")
        wc = db_connect(db_file)
        wc.execute("DELETE FROM cattle_item_usage_log")
        wc.execute("DELETE FROM cattle_cost_log")
        wc.execute("DELETE FROM disease_record")
        wc.execute("DELETE FROM cattle")
        wc.commit(); wc.close()
        st.session_state["cattle_reset_done"] = True
        st.session_state["cattle_reset_done_farm"] = target_farm
        st.rerun()


with st.sidebar.expander("🐂 등록된 개체 전체 삭제"):
    st.caption("선택한 농장에 등록된 개체(입식 내역)만 삭제합니다. 시험군·품목 등은 유지됩니다. 실행 직전 자동으로 백업본을 만듭니다.")
    if st.button("개체 전체 삭제 실행", width="stretch", key="open_cattle_reset"):
        st.session_state["show_cattle_reset_dialog"] = True

if st.session_state.pop("show_cattle_reset_dialog", False):
    cattle_reset_dialog()
if st.session_state.pop("cattle_reset_done", False):
    done_farm = st.session_state.pop("cattle_reset_done_farm", "")
    st.sidebar.success(f"'{done_farm}'에 등록된 개체가 모두 삭제되었습니다. (직전 상태는 백업에 보관)")

st.sidebar.markdown("---")


if selected_farm == "시험농장 전체 현황":
    st.markdown(
        f"""
        <div class="farm-header" style="display:flex; align-items:center; gap:14px; margin-bottom:0.4rem;">
            <img class="farm-title-logo" src="{MEDAL_ICON_DATA_URI}" alt="심볼" style="height: 44px;">
            <div style="min-width:0;">
                <h1 class="farm-title" style="margin:0; padding:0;"><span class="title-long">대구축협 시험농장 관리 시스템</span><span class="title-short">시험농장 현황</span></h1>
                <div style="color:#5F6F66; font-size:0.95rem; margin-top:2px;">
                    등록된 모든 농장의 개체 현황을 한눈에 보는 통합 대시보드
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.subheader("🌐 농장 통합 대시보드")
    
    all_cattle_dfs = []
    
    for farm_nm, farm_cfg_info in FARM_CONFIG.items():
        f_db = farm_cfg_info["db_file"]
        if db_exists(f_db):
            try:
                f_conn = sqlite3.connect(f_db)
                df_f = pd.read_sql("SELECT cattle_id as 개체번호, status as 상태, admission_date as 입식일, castration_date as 거세일, closure_date as 종결일, initial_cost as 초기원가, market_name as 우시장 FROM cattle", f_conn)
                df_f['농장명'] = farm_nm
                all_cattle_dfs.append(df_f)
                f_conn.close()
            except Exception as e:
                pass
                
    if all_cattle_dfs:
        df_all = pd.concat(all_cattle_dfs, ignore_index=True)
        
        total_admission = len(df_all)
        current_breeding = len(df_all[df_all['상태'] == '사육'])
        dead_cattle = len(df_all[df_all['상태'] == '폐사'])
        shipped_cattle = len(df_all[df_all['상태'] == '출하'])
        total_initial_cost = int(df_all['초기원가'].sum(skipna=True)) if '초기원가' in df_all.columns else 0
        
        def _share(part):
            return f"{part / total_admission * 100:.1f}%" if total_admission else "0.0%"

        m1, m2, m3, m4, m5 = st.container(key="kpi_all").columns(5)
        m1.metric("전체 누적 입식", f"{total_admission:,}두",
                  delta=f"{len(all_cattle_dfs)}개 농장 합계", delta_color="off", delta_arrow="off")
        m2.metric("현재 사육중", f"{current_breeding:,}두",
                  delta=f"전체의 {_share(current_breeding)}", delta_color="off", delta_arrow="off")
        m3.metric("누적 폐사", f"{dead_cattle:,}두",
                  delta=f"폐사율 {_share(dead_cattle)}", delta_color="inverse" if dead_cattle else "off",
                  delta_arrow="off")
        m4.metric("누적 출하", f"{shipped_cattle:,}두",
                  delta=f"출하율 {_share(shipped_cattle)}", delta_color="off", delta_arrow="off")
        m5.metric("총 구입비용", f"{total_initial_cost // 10000:,}만원",
                  delta="송아지 구입비 합계", delta_color="off", delta_arrow="off")

        st.write("")

        count_col = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="center")
        money_col = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right", step=1)

        st.markdown("##### 🏢 농장별 요약 현황")
        farm_summary = df_all.groupby('농장명').agg(
            전체입식=('개체번호', 'count'),
            사육중=('상태', lambda x: int((x == '사육').sum())),
            출하=('상태', lambda x: int((x == '출하').sum())),
            폐사=('상태', lambda x: int((x == '폐사').sum())),
            총구입비용_만원=('초기원가', lambda x: int(x.sum(skipna=True)) // 10000),
            평균구입금액_만원=('초기원가', lambda x: int(x.mean(skipna=True)) // 10000 if not x.isna().all() else 0)
        ).reset_index()
        # 폐사율 = 누적 폐사 / 전체 입식
        farm_summary.insert(
            farm_summary.columns.get_loc('폐사') + 1, '폐사율',
            (farm_summary['폐사'] / farm_summary['전체입식'].where(farm_summary['전체입식'] > 0) * 100).fillna(0),
        )
        farm_dataframe(
            farm_summary,
            width="stretch", hide_index=True,
            column_config={
                "농장명": st.column_config.TextColumn("농장명"),
                "전체입식": count_col("전체 입식 (두)"),
                "사육중": count_col("현재 사육중 (두)"),
                "출하": count_col("누적 출하 (두)"),
                "폐사": count_col("누적 폐사 (두)"),
                "폐사율": st.column_config.NumberColumn("폐사율 (%)", format="%.1f%%", alignment="right"),
                "총구입비용_만원": money_col("총 구입비용 (만원)"),
                "평균구입금액_만원": money_col("두당 평균 (만원)"),
            },
        )
        st.write("")

        st.markdown("##### 🏪 농장별 우시장 구입 현황")
        df_market = df_all[df_all['우시장'].notna() & (df_all['우시장'].astype(str).str.strip() != '')]
        if not df_market.empty:
            market_summary = df_market.groupby(['농장명', '우시장']).agg(
                구입마릿수=('개체번호', 'count'),
                총구입비용_만원=('초기원가', lambda x: int(x.sum(skipna=True)) // 10000),
                평균구입비용_만원=('초기원가', lambda x: int(x.mean(skipna=True)) // 10000 if not x.isna().all() else 0)
            ).reset_index()
            farm_dataframe(
                market_summary,
                width="stretch", hide_index=True,
                column_config={
                    "구입마릿수": count_col("구입 마릿수 (두)"),
                    "총구입비용_만원": money_col("총 구입비용 (만원)"),
                    "평균구입비용_만원": money_col("두당 평균 (만원)"),
                },
            )
        else:
            st.info("등록된 우시장 구입 이력이 없습니다.")

        st.write("")
                
        with st.expander("통합 데이터 상세 표"):
            df_detail = df_all.copy()
            df_detail.insert(0, '순번', range(1, len(df_detail) + 1))
            for col in ['입식일', '거세일', '종결일']:
                df_detail[col] = pd.to_datetime(df_detail[col], errors='coerce')
            df_detail['상태'] = df_detail['상태'].map({'사육': '🟢 사육', '출하': '🔵 출하', '폐사': '🔴 폐사'}).fillna(df_detail['상태'])
            farm_dataframe(
                df_detail,
                width="stretch",
                hide_index=True,
                column_config={
                    "순번": st.column_config.NumberColumn(width=60),
                    "입식일": st.column_config.DateColumn(format="YYYY-MM-DD"),
                    "거세일": st.column_config.DateColumn(format="YYYY-MM-DD"),
                    "종결일": st.column_config.DateColumn(format="YYYY-MM-DD"),
                    "초기원가": money_col("구입비용 (원)"),
                },
            )
            
        st.markdown("---")
        st.subheader("📑 통합 보고서 생성")
        st.markdown("현재 전체 현황 대시보드의 요약 수치 및 농장별 데이터 표를 기반으로 인쇄 가능한 HTML 보고서를 생성합니다.")
        if st.button("📄 보고서 생성", type="primary"):
            st.session_state["overall_report_html"] = generate_overall_report_html(
                df_all,
                list(FARM_CONFIG.keys()),
                {k: v.get("color", "#4F46E5") for k, v in FARM_CONFIG.items()},
            )
            
        overall_html = st.session_state.get("overall_report_html")
        if overall_html:
            st.success("통합 보고서가 생성되었습니다.")
            st.download_button(
                "⬇️ HTML 파일로 내려받기",
                overall_html.encode("utf-8"),
                file_name=f"통합보고서_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
                mime="text/html",
                width="stretch",
            )
            st.markdown("###### 미리보기")
            st.components.v1.html(overall_html, height=1400, scrolling=True)

    else:
        st.info("데이터가 있는 농장이 없습니다.")
    st.stop()

farm_cfg = FARM_CONFIG[selected_farm]
DB_FILE = farm_cfg["db_file"]
farm_color = farm_cfg["color"]
farm_b_cnt = farm_cfg.get("buildings_count", 6)
farm_p_cnt = farm_cfg.get("pens_count", 20)


# DB 자동 생성
if not db_exists(DB_FILE):
    init_db(selected_farm)

# 기존 DB에 나중에 추가된 표가 빠져 있으면 채워 넣는다 (신규 생성 직후에도 실행되지만
# 전부 IF NOT EXISTS 라서 안전하다).
migrate_schema(DB_FILE)

# 그날 첫 접속이면 자동 백업
daily_backup(DB_FILE)

st.sidebar.markdown("---")
st.sidebar.markdown("##### 💾 데이터 백업")
with st.sidebar.expander("백업 만들기 / 내려받기"):
    st.caption("저장 폴더: %s" % MANUAL_BACKUP_DIR)
    if st.button("지금 백업 만들기", width="stretch"):
        try:
            saved = backup_db(DB_FILE, "manual", dest_dir=MANUAL_BACKUP_DIR)
            st.session_state["last_backup"] = saved
            st.success("%s 백업 저장 완료 · %s" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), os.path.basename(saved)))
        except Exception as e:
            st.session_state["last_backup"] = None
            st.error("백업에 실패했습니다: %s" % e)
    last_backup = st.session_state.get("last_backup")
    if last_backup and os.path.exists(last_backup):
        with open(last_backup, "rb") as fh:
            st.download_button(
                "내려받기: " + os.path.basename(last_backup),
                fh.read(),
                file_name=os.path.basename(last_backup),
                mime="application/octet-stream",
                width="stretch",
            )
    kept = glob.glob(os.path.join(BACKUP_DIR, os.path.splitext(os.path.basename(DB_FILE))[0] + "_*.db"))
    st.caption("보관 중인 백업 %d개 · 매일 첫 접속 시 자동 백업 (최근 %d개 유지)" % (len(kept), BACKUP_KEEP))

with st.sidebar.expander("백업 파일로 복원"):
    st.caption("내려받아 둔 .db 백업 파일로 현재 데이터를 되돌립니다. 복원 직전 현재 상태도 자동 백업됩니다.")
    restore_file = st.file_uploader("백업 파일 선택", type=["db"], key="restore_uploader")
    if restore_file is not None and st.button("이 파일로 덮어쓰기", width="stretch"):
        ok, msg = restore_db(DB_FILE, restore_file.getvalue())
        if ok:
            notify(msg, icon="✅")
            st.rerun()
        else:
            st.error(msg)

# 리셋 확인용 비밀번호. 환경변수(ERP_RESET_PASSWORD) 나 secrets 로 덮어쓸 수 있다.
RESET_PASSWORD_DEFAULT = "1234"


def reset_password():
    """리셋 확인 비밀번호를 찾는다. 전용 값(ERP_RESET_PASSWORD) → 접속 비밀번호(ERP_PASSWORD) → 기본값."""
    for key in ("ERP_RESET_PASSWORD", "ERP_PASSWORD"):
        value = os.environ.get(key)
        if not value:
            try:
                value = st.secrets.get(key)
            except Exception:
                value = None
        if value:
            return value
    return RESET_PASSWORD_DEFAULT


def reset_preview(db_file):
    """리셋으로 지워질 데이터 건수. 표시용이므로 읽기에 실패한 항목은 조용히 넘어간다."""
    tables = [
        ("개체", "cattle"),
        ("질병·처방 기록", "disease_record"),
        ("품목 매입", "purchase"),
        ("월별 사용량", "monthly_usage"),
        ("월별 고정비", "monthly_fixedcost"),
        ("개체별 원가 내역", "cattle_cost_log"),
    ]
    rows = []
    if not db_exists(db_file):
        return rows
    conn = db_connect(db_file)
    try:
        for label, table in tables:
            try:
                count = conn.execute("SELECT COUNT(*) FROM %s" % table).fetchone()[0]
            except sqlite3.Error:
                continue
            if count:
                rows.append((label, count))
    finally:
        conn.close()
    return rows


@st.dialog("⚠️ 정말 초기화하시겠습니까?")
def reset_dialog(farm_name, db_file):
    st.error(f"**{farm_name}**의 모든 데이터가 삭제됩니다. 이 작업은 되돌릴 수 없습니다.")
    rows = reset_preview(db_file)
    if rows:
        st.markdown("**지워지는 데이터**")
        st.markdown("\n".join("- %s **%s건**" % (label, format(count, ",")) for label, count in rows))
    else:
        st.caption("현재 입력된 데이터가 없습니다.")
    st.caption("실행 직전 자동으로 백업본을 만들기 때문에, 사이드바의 '백업 파일로 복원'으로 되돌릴 수 있습니다.")

    pw = st.text_input("계속하려면 관리자 비밀번호를 입력하세요", type="password", key="reset_pw")
    col_cancel, col_run = st.columns(2)
    if col_cancel.button("취소", width="stretch"):
        st.session_state.pop("reset_pw", None)
        st.rerun()
    # disabled 를 쓰면 비밀번호를 입력한 직후 첫 클릭이 먹히지 않으므로, 눌렀을 때 검사한다.
    if col_run.button("삭제하고 초기화", type="primary", width="stretch"):
        if not pw:
            st.warning("비밀번호를 입력하세요.")
        elif pw != reset_password():
            st.error("비밀번호가 올바르지 않습니다. 초기화하지 않았습니다.")
        else:
            backup_db(db_file, "before-reset")
            init_db(farm_name)
            st.session_state.pop("reset_pw", None)
            st.session_state["reset_done"] = True
            st.rerun()


st.sidebar.markdown("---")
with st.sidebar.expander("⚠️ 초기 상태로 리셋"):
    st.caption(f"**{selected_farm}**의 데이터가 **전부 삭제**됩니다. 실행 직전 자동으로 백업본을 만듭니다.")
    if st.button("리셋 실행", width="stretch"):
        st.session_state.pop("reset_pw", None)
        st.session_state["show_reset_dialog"] = True

if st.session_state.pop("show_reset_dialog", False):
    reset_dialog(selected_farm, DB_FILE)
if st.session_state.pop("reset_done", False):
    st.sidebar.success("데이터베이스가 리셋되었습니다. (직전 상태는 백업에 보관)")

st.sidebar.markdown("---")
st.sidebar.caption("저장 위치: %s" % ("Supabase (PostgreSQL)" if USE_PG else DB_DIR))
if USE_PG and st.sidebar.button(
    "🔄 최신 데이터 다시 불러오기", width="stretch",
    help="화면을 빠르게 하려고 조회 결과를 잠시(최대 10분) 기억해 둡니다. 이 화면에서 등록·수정한 내용은 바로 반영되지만, "
         "다른 PC나 Supabase 에서 직접 바꾼 내용을 바로 보려면 누르세요.",
):
    db_adapter.clear_cache()
    st.rerun()

# 타이틀 (선택된 농장 표시)
st.markdown(
    f"""
    <div class="farm-header" style="display:flex; align-items:center; justify-content:space-between; gap:16px; flex-wrap:wrap; margin-bottom:0.4rem;">
        <div style="display:flex; align-items:center; gap:14px; min-width:0;">
            <img class="farm-title-logo" src="{MEDAL_ICON_DATA_URI}" alt="심볼" style="height: 44px;">
            <div style="min-width:0;">
                <h1 class="farm-title" style="margin:0; padding:0;"><span class="title-long">대구축협 시험농장 관리 시스템</span><span class="title-short">시험농장 관리</span></h1>
                <div style="color:#5F6F66; font-size:0.95rem; margin-top:2px;">
                    개체 입식 · 사육 현황 · 품목 매입 · 월말 원가 정산(사육일수 비례 배분)
                </div>
            </div>
        </div>
        <div style="background:{html.escape(farm_color)}; color:white; padding:8px 20px; border-radius:999px;
                    font-weight:800; font-size:1rem; box-shadow:0 4px 12px rgba(0,0,0,0.12);">
            {html.escape(selected_farm)}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

conn = db_connect(DB_FILE)

# KPI Metrics 표시
status_counts = dict(conn.execute("SELECT status, COUNT(*) FROM cattle GROUP BY status").fetchall())
total_admitted = sum(status_counts.values())
current_raising = status_counts.get('사육', 0)
dead_count = status_counts.get('폐사', 0)
shipped_count = status_counts.get('출하', 0)

this_month = datetime.now().strftime('%Y-%m')
month_admit_cnt, month_calf_cost = conn.execute(
    "SELECT COUNT(*), COALESCE(SUM(initial_cost), 0) FROM cattle WHERE substr(admission_date, 1, 7) = ?",
    (this_month,),
).fetchone()
month_purchase_amt = conn.execute(
    "SELECT COALESCE(SUM(total_amount), 0) FROM purchase WHERE substr(purchase_date, 1, 7) = ?",
    (this_month,),
).fetchone()[0]

# 평균 개월령 계산 (사육 중인 개체 대상)
active_cattle_dates = pd.read_sql("SELECT birth_date FROM cattle WHERE status='사육' AND birth_date IS NOT NULL AND birth_date != ''", conn)
avg_months_str = "-"
if not active_cattle_dates.empty:
    import calendar
    today_dt = pd.to_datetime('today')
    total_m, total_d, valid_cnt = 0, 0, 0
    for bdate in active_cattle_dates['birth_date']:
        try:
            start = pd.to_datetime(bdate)
            end = today_dt
            months = (end.year - start.year) * 12 + (end.month - start.month)
            if end.day >= start.day:
                days = end.day - start.day
            else:
                months -= 1
                prev_month = 12 if end.month == 1 else end.month - 1
                prev_year = end.year - 1 if end.month == 1 else end.year
                _, days_in_prev_month = calendar.monthrange(prev_year, prev_month)
                days = days_in_prev_month - start.day + end.day
            
            total_m += (months + 1)
            total_d += days
            valid_cnt += 1
        except: pass
        
    if valid_cnt > 0:
        total_days_all = (total_m * 30.436875) + total_d
        avg_days = total_days_all / valid_cnt
        avg_m = int(avg_days // 30.436875)
        avg_d = int(avg_days % 30.436875)
        avg_months_str = f"{avg_m}개월 {avg_d}일"

def _rate(part):
    return f"{part / total_admitted * 100:.1f}%" if total_admitted else "0.0%"

kpi_cols = st.container(key="kpi_main").columns(6)
kpi_cols[0].metric("현재 사육 두수", f"{current_raising:,}두",
                   delta=f"전체 입식 {total_admitted:,}두", delta_color="off", delta_arrow="off")
kpi_cols[1].metric("누적 출하", f"{shipped_count:,}두",
                   delta=f"출하율 {_rate(shipped_count)}", delta_color="off", delta_arrow="off")
kpi_cols[2].metric("누적 폐사", f"{dead_count:,}두",
                   delta=f"폐사율 {_rate(dead_count)}", delta_color="inverse" if dead_count else "off",
                   delta_arrow="off")
kpi_cols[3].metric(f"{datetime.now().month}월 송아지 구입비", f"{int(month_calf_cost) // 10000:,}만원",
                   delta=f"이달 {month_admit_cnt:,}두 입식", delta_color="off", delta_arrow="off")
kpi_cols[4].metric(f"{datetime.now().month}월 품목 매입액", f"{int(month_purchase_amt) // 10000:,}만원",
                   delta="사료·약품 등 입고", delta_color="off", delta_arrow="off")
kpi_cols[5].metric("평균 개월령", avg_months_str,
                   delta="사육중 개체 기준", delta_color="off", delta_arrow="off")
st.write("")

tab_cattle, tab1, tab0, tab2, tab_settle, tab_report, tab_slaughter = st.tabs(["🐂 개체 관리", "📊 사육 및 재고 현황", "📦 품목·매입 관리", "💰 월말 등록", "🚀 월말 정산", "🧾 결산 리포트", "🥩 도축 성적"])

# ===== 개체 관리 탭 =====
with tab_cattle:
    sub_tab1, sub_tab2, sub_tab3 = st.tabs(["🐂 입식 등록", "📋 상태 변경 / 질병 기록", "📊 전체 현황"])
    
    with sub_tab1:
        col_reg1, col_reg2 = st.columns(2)
        
        with col_reg1:
            st.subheader("🧪 시험군 등록")
            with st.form("add_group_form", clear_on_submit=True):
                new_group_name = st.text_input("시험군 명칭", placeholder="예: 대조군, 처리군A 등")
                new_group_start = st.date_input("시작일")
                
                st.markdown("###### 📍 자동 할당 조건 1")
                st.caption("개체 일괄 등록 시 아래 동/우방 조건에 맞는 개체를 이 시험군으로 자동 배정합니다.")
                buildings_list = [f"{i}동" for i in range(1, farm_b_cnt + 1)]
                pens_list = [str(i) for i in range(1, farm_p_cnt + 1)]
                
                col_c1_1, col_c1_2 = st.columns(2)
                with col_c1_1: new_group_b1 = st.multiselect("대상 동 (조건 1)", options=buildings_list)
                with col_c1_2: new_group_p1 = st.multiselect("대상 우방 (조건 1)", options=pens_list)
                
                st.markdown("###### 📍 자동 할당 조건 2 (필요 시 추가)")
                col_c2_1, col_c2_2 = st.columns(2)
                with col_c2_1: new_group_b2 = st.multiselect("대상 동 (조건 2)", options=buildings_list)
                with col_c2_2: new_group_p2 = st.multiselect("대상 우방 (조건 2)", options=pens_list)

                submitted_group = st.form_submit_button("시험군 등록", type="primary", width="stretch")

                if submitted_group:
                    if new_group_name:
                        try:
                            wc = db_connect(DB_FILE)
                            conditions = []
                            if new_group_b1 or new_group_p1: conditions.append({"buildings": new_group_b1, "pens": new_group_p1})
                            if new_group_b2 or new_group_p2: conditions.append({"buildings": new_group_b2, "pens": new_group_p2})
                            loc_map = json.dumps(conditions, ensure_ascii=False) if conditions else None
                            # 코드를 입력받지 않고 명칭을 코드로 동일하게 사용
                            wc.execute("INSERT INTO testgroup_master (test_group_code, test_name, start_date, location_mapping) VALUES (?, ?, ?, ?)", (new_group_name, new_group_name, new_group_start.isoformat(), loc_map))
                            wc.commit(); wc.close()
                            notify(f"시험군 '{new_group_name}' 등록 완료", icon="✅")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            wc.rollback(); wc.close()
                            st.error("이미 존재하는 시험군 명칭입니다.")
                    else:
                        st.warning("시험군 명칭을 입력하세요.")
            st.caption("기존 개체를 시험군에 배정하거나 옮기려면 '상태 변경 / 질병 기록' 탭의 '개체 위치 및 시험군 이동'을 이용하세요.")

            df_groups = pd.read_sql("SELECT test_group_code as 시험군코드, test_name as 시험명칭, start_date as 시작일, end_date as 종료일, location_mapping as 자동할당조건 FROM testgroup_master", conn)
            
            def format_loc(x):
                if pd.isna(x) or not x: return ""
                try:
                    d = json.loads(x)
                    d_list = d if isinstance(d, list) else [d]
                    res = []
                    for cond in d_list:
                        b = ",".join(cond.get('buildings', []))
                        
                        p_list = cond.get('pens', [])
                        nums = []
                        for p_str in p_list:
                            try: nums.append(int(p_str))
                            except: pass
                            
                        if nums and len(nums) == len(p_list):
                            nums = sorted(list(set(nums)))
                            ranges = []
                            start = end = nums[0]
                            for i in range(1, len(nums)):
                                if nums[i] == end + 1:
                                    end = nums[i]
                                else:
                                    ranges.append(f"{start}-{end}" if start != end else str(start))
                                    start = end = nums[i]
                            ranges.append(f"{start}-{end}" if start != end else str(start))
                            p = ",".join(ranges)
                        else:
                            p = ",".join(p_list)
                            
                        if b and p: res.append(f"{b}({p}번방)")
                        elif b: res.append(b)
                        elif p: res.append(f"{p}번방")
                    return " / ".join(res)
                except: return ""
                
            with st.expander(f"📋 등록된 시험군 ({len(df_groups)}개) · 수정 / 삭제"):
                if '자동할당조건' in df_groups.columns:
                    df_groups['자동할당조건'] = df_groups['자동할당조건'].apply(format_loc)
            
                # 시험군코드는 UI 화면 테이블에서 숨김 처리
                farm_dataframe(df_groups[['시험명칭', '시작일', '종료일', '자동할당조건']], width="stretch", hide_index=True)
            
                if not df_groups.empty:
                    st.markdown("##### 시험군 수정 / 삭제")
                
                    # 시험명칭만 깔끔하게 표시
                    group_opts = {r['시험명칭']: r['시험군코드'] for _, r in df_groups.iterrows()}
                    edit_target = st.selectbox("대상 시험군", list(group_opts.keys()), key="edit_group_sel")
                
                    target_code = group_opts[edit_target]
                    current_row = df_groups[df_groups['시험군코드'] == target_code].iloc[0]
                    current_name = current_row['시험명칭']
                    current_start = pd.to_datetime(current_row['시작일']).date()

                    with st.form("edit_group_form", clear_on_submit=False):
                        edit_name = st.text_input("새 시험명칭", value=current_name)
                        edit_start = st.date_input("새 시작일", value=current_start)

                        col_b1, col_b2 = st.columns(2)
                        with col_b1:
                            submitted_edit = st.form_submit_button("수정", type="primary", width="stretch")
                        with col_b2:
                            submitted_delete = st.form_submit_button("삭제", type="secondary", width="stretch")

                        if submitted_edit:
                            if edit_name:
                                wc = db_connect(DB_FILE)
                                wc.execute("UPDATE testgroup_master SET test_name = ?, start_date = ? WHERE test_group_code = ?", (edit_name, edit_start.isoformat(), target_code))
                                wc.commit(); wc.close()
                                notify("시험군 수정 완료", icon="✅")
                                st.rerun()
                            else:
                                st.warning("새 시험명칭을 입력하세요.")

                        if submitted_delete:
                            wc = db_connect(DB_FILE)
                            # 개체가 있는지 확인하여 무결성 오류 방지
                            cattle_cnt = pd.read_sql("SELECT COUNT(*) as cnt FROM cattle WHERE test_group_code=?", wc, params=(target_code,)).iloc[0]['cnt']
                            if cattle_cnt > 0:
                                st.error(f"이 시험군에 등록된 개체가 {cattle_cnt}마리 있어 삭제할 수 없습니다. 개체를 먼저 삭제/이동하세요.")
                                wc.close()
                            else:
                                wc.execute("DELETE FROM testgroup_master WHERE test_group_code = ?", (target_code,))
                                wc.commit(); wc.close()
                                notify(f"시험군 '{target_code}' 삭제 완료", icon="✅")
                                st.rerun()

        with col_reg2:
            st.subheader("🐂 개체 입식 등록")
            st.caption("표에 여러 마리를 한꺼번에 입력해 등록하거나, 엑셀 파일을 통해 일괄 등록할 수 있습니다.")
            
            groups_for_cattle = pd.read_sql("SELECT test_group_code, test_name FROM testgroup_master", conn)
            if groups_for_cattle.empty:
                st.info("먼저 시험군을 등록해 주세요.")
            else:
                cattle_group_opts = {r['test_name']: r['test_group_code'] for _, r in groups_for_cattle.iterrows()}
                
                # 등록 직후 rerun 으로 화면이 새로 그려져도 결과가 보이도록 세션에 담아 두고 여기서 보여 준다.
                bulk_msg = st.session_state.pop("bulk_upload_msg", None)
                with st.expander("📁 엑셀로 일괄 등록", expanded=bool(bulk_msg)):
                    if bulk_msg:
                        (st.success if bulk_msg[0] == "ok" else st.warning)(bulk_msg[1])
                    st.markdown("**1. 일괄 등록용 엑셀 양식 다운로드**")
                    st.info("💡 **팁:** 엑셀의 '동', '우방 (칸번호)' 열에 위치를 입력하면 업로드 시 설정한 시험군으로 자동 할당됩니다.")
                    df_template = pd.DataFrame(columns=[
                        '이표번호 (개체번호)', 'KPN', '생년월일', '입식일 (구입일)',
                        '우시장', '동', '우방 (칸번호)', 
                        '송아지 구입금액', '수수료', '운송료', 
                        '거세일자', '가축보험 가입금액', '가축보험 보험료', '비고'
                    ])
                    df_template.loc[0] = [
                        '410002123456789 (예시)', 'KPN1234', '2023-01-01', '2023-06-01',
                        '안성우시장', '1동', '5', 
                        3000000, 50000, 100000, 
                        '', '', '', '작성 전 이 줄은 지우고 입력하세요'
                    ]
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_template.to_excel(writer, index=False, sheet_name='입식양식')
                    st.download_button(
                        label="엑셀 양식 다운로드",
                        data=output.getvalue(),
                        file_name="개체입식_일괄등록_양식.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.markdown("---")
                    st.markdown("**2. 작성한 엑셀 파일 업로드 및 시험군 할당**")
                    cattle_group_opts = {"✨ 동/우방 조건으로 자동 할당": "AUTO", **cattle_group_opts}
                    bulk_group_label = st.selectbox("업로드할 개체들의 **소속 시험군** 선택", list(cattle_group_opts.keys()), key="bulk_group_sel")
                    uploaded_file = st.file_uploader("작성된 엑셀 파일을 선택하세요.", type=["xlsx", "xls"])
                    
                    if uploaded_file is not None:
                        try:
                            df_upload = pd.read_excel(uploaded_file)
                            # 컬럼명 공백 차이를 먼저 정리한다 (예전 양식 호환)
                            df_upload.columns = df_upload.columns.str.strip()

                            # 이표번호 컬럼 찾기 (공백 유무 확인)
                            id_col = next(
                                (c for c in ("이표번호 (개체번호)", "이표번호(개체번호)") if c in df_upload.columns), None
                            )
                            if id_col is None:
                                st.error("엑셀 파일에 '이표번호 (개체번호)' 항목이 없습니다. 올바른 양식을 사용해 주세요. (현재 파일 항목: %s)" % ", ".join(df_upload.columns))
                                st.stop()

                            # 이미 등록된 개체는 건너뛰고, 파일에서 새로 추가된 개체만 등록한다.
                            existing_ids = set(
                                pd.read_sql("SELECT cattle_id FROM cattle", conn)["cattle_id"].astype(str).str.strip()
                            )
                            file_ids = []
                            for val in df_upload[id_col]:
                                text = clean_excel_text(val)
                                if text:
                                    file_ids.append(text)
                            unique_ids = list(dict.fromkeys(file_ids))
                            new_ids = [v for v in unique_ids if v not in existing_ids]
                            st.info(
                                "파일 %d건 · 새로 등록될 개체 **%d건** · 이미 등록되어 건너뛸 개체 %d건"
                                % (len(file_ids), len(new_ids), len(unique_ids) - len(new_ids))
                            )
                            if new_ids:
                                with st.expander("새로 등록될 이표번호 %d건 보기" % len(new_ids)):
                                    st.write(", ".join(new_ids))
                            fill_missing_dates = st.checkbox(
                                "이미 등록된 개체의 비어 있는 생년월일·입식일은 파일 값으로 채우기",
                                value=False,
                                key="bulk_fill_dates",
                            )

                            do_bulk_upload = st.button("개체 일괄등록", type="primary", width="stretch")

                            if do_bulk_upload:
                                wc = db_connect(DB_FILE)
                                success_cnt, skip_cnt, fail_cnt, filled_cnt = 0, 0, 0, 0
                                fail_reasons = {}
                                seen_ids = set()
                                selected_bulk_group_code = cattle_group_opts[bulk_group_label]
                                
                                auto_locations = {}
                                if selected_bulk_group_code == "AUTO":
                                    try:
                                        loc_df = pd.read_sql("SELECT test_group_code, location_mapping FROM testgroup_master WHERE location_mapping IS NOT NULL", wc)
                                        for _, r in loc_df.iterrows():
                                            try: auto_locations[r['test_group_code']] = json.loads(r['location_mapping'])
                                            except: pass
                                    except: pass

                                for _, row in df_upload.iterrows():
                                    try:
                                        # 빈칸 및 과거 양식 호환 처리를 위한 헬퍼 함수
                                        def get_col_val(cols):
                                            if isinstance(cols, str): cols = [cols]
                                            for c in cols:
                                                if c in df_upload.columns:
                                                    return row.get(c)
                                            return None

                                        def get_str(cols):
                                            return clean_excel_text(get_col_val(cols))

                                        def get_num(cols, default=0):
                                            val = get_col_val(cols)
                                            if pd.isna(val) or val is None: return default
                                            v = str(val).strip()
                                            if not v or v.lower() == 'nan' or v == 'None': return default
                                            try:
                                                return float(val)
                                            except:
                                                return default
                                        
                                        def get_date(cols):
                                            val = get_col_val(cols)
                                            if pd.isna(val) or val is None: return None
                                            if hasattr(val, 'strftime'): return val.strftime('%Y-%m-%d')
                                            v = str(val).strip()
                                            if not v or v.lower() == 'nan' or v == 'None': return None
                                            return v[:10]
                                        
                                        cid = get_str(['이표번호 (개체번호)', '이표번호(개체번호)'])
                                        if not cid: continue
                                        
                                        calf_p = get_num('송아지 구입금액', 0)
                                        comm_f = get_num('수수료', 0)
                                        trans_f = get_num('운송료', 0)
                                        init_c = calf_p + comm_f + trans_f
                                        
                                        birth_d = get_date(['생년월일', '생년월일(YYYY-MM-DD)'])
                                        admin_d = get_date(['입식일 (구입일)', '입식일(YYYY-MM-DD)'])
                                        castr_d = get_date(['거세일자', '거세', '거세일(YYYY-MM-DD)'])

                                        # 파일에 새로 추가된 개체만 등록한다. 이미 등록됐거나 파일 안에서 중복된 이표번호는 건너뛴다.
                                        if cid in existing_ids or cid in seen_ids:
                                            if fill_missing_dates and cid in existing_ids and (birth_d or admin_d):
                                                cur = wc.execute(
                                                    "SELECT birth_date, admission_date FROM cattle WHERE cattle_id = ?", (cid,)
                                                ).fetchone()
                                                sets, params = [], []
                                                if cur and birth_d and not (cur[0] or "").strip():
                                                    sets.append("birth_date = ?")
                                                    params.append(birth_d)
                                                if cur and admin_d and not (cur[1] or "").strip():
                                                    sets.append("admission_date = ?")
                                                    params.append(admin_d)
                                                if sets:
                                                    params.append(cid)
                                                    wc.execute("UPDATE cattle SET %s WHERE cattle_id = ?" % ", ".join(sets), params)
                                                    filled_cnt += 1
                                            skip_cnt += 1
                                            seen_ids.add(cid)
                                            continue
                                        seen_ids.add(cid)
                                        
                                        status_val = get_str('상태') or '사육'
                                        kpn_val = get_str('KPN')
                                        market_val = get_str('우시장')
                                        building_val = get_str('동')
                                        pen_val = get_num(['우방 (칸번호)', '우방(칸번호)'], None)
                                        feed_val = get_str(['사료구분', '사료구분(표준/증량형/제한형)', '사료구분(증량형/제한형)']) or '표준'
                                        roughage_val = get_str(['조사료등급', '조사료등급(표준/고급/저급)', '조사료등급(고급/저급)']) or '표준'
                                        memo_val = get_str('비고')
                                        ins_v = get_num(['가축보험 가입금액', '가축보험 가입금액(만원)'], None)
                                        ins_p = get_num('가축보험 보험료', None)
                                        
                                        final_group_code = selected_bulk_group_code
                                        if final_group_code == "AUTO":
                                            final_group_code = None
                                            pen_str = str(int(pen_val)) if pen_val is not None else ""
                                            for gcode, locs in auto_locations.items():
                                                if isinstance(locs, dict): locs = [locs]
                                                matched = False
                                                for cond in locs:
                                                    b_list = cond.get('buildings', [])
                                                    p_list = cond.get('pens', [])
                                                    if not b_list and not p_list: continue
                                                    
                                                    b_match = True
                                                    if b_list:
                                                        b_match = False
                                                        b_val_clean = str(building_val).replace("동", "").strip() if building_val else ""
                                                        for b in b_list:
                                                            if str(b).replace("동", "").strip() == b_val_clean:
                                                                b_match = True
                                                                break
                                                                
                                                    p_match = (pen_str in p_list) if p_list else True
                                                    if b_match and p_match:
                                                        matched = True
                                                        break
                                                if matched:
                                                    final_group_code = gcode
                                                    break
                                        
                                        wc.execute(
                                            "INSERT INTO cattle (cattle_id, kpn, birth_date, test_group_code, status, admission_date, market_name, building, pen_number, feed_type, roughage_grade, castration_date, calf_price, commission_fee, transport_fee, initial_cost, insurance_value, insurance_premium, memo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                            (
                                                cid, kpn_val, birth_d,
                                                final_group_code, status_val, admin_d,
                                                market_val, building_val, pen_val,
                                                feed_val, roughage_val, castr_d,
                                                calf_p, comm_f, trans_f, init_c,
                                                ins_v, ins_p, memo_val
                                            )
                                        )
                                        success_cnt += 1
                                    except sqlite3.IntegrityError as e:
                                        if "UNIQUE" in str(e):
                                            # 위에서 걸러지지 않은 중복(동시 입력 등)도 오류가 아닌 '건너뜀'으로 처리
                                            skip_cnt += 1
                                        else:
                                            fail_reasons.setdefault(str(e), []).append(cid)
                                            fail_cnt += 1
                                    except Exception as e:
                                        fail_reasons.setdefault(str(e), []).append(cid)
                                        fail_cnt += 1
                                wc.commit(); wc.close()
                                parts = ["신규 %d건 등록" % success_cnt]
                                if skip_cnt:
                                    parts.append("이미 등록되어 건너뜀 %d건" % skip_cnt)
                                if filled_cnt:
                                    parts.append("기존 개체 빈 날짜 보완 %d건" % filled_cnt)
                                if fail_cnt:
                                    parts.append("오류 %d건" % fail_cnt)
                                summary = " · ".join(parts)
                                if not (success_cnt or filled_cnt):
                                    summary = "새로 추가된 개체가 없습니다. (%s)" % summary
                                # st.rerun() 이 화면을 지워 버리므로, 오류 사유도 요약에 같이 담아 둔다.
                                for reason, ids in list(fail_reasons.items())[:3]:
                                    summary += "\n\n- %s → %d건 (예: %s)" % (
                                        reason, len(ids), ", ".join(ids[:3])
                                    )
                                st.session_state["bulk_upload_msg"] = (
                                    "ok" if (success_cnt or filled_cnt) and not fail_cnt else "warn",
                                    summary,
                                )
                                st.rerun()
                        except Exception as e:
                            st.error(f"엑셀 파일 읽기 오류: {e}")
                
                st.markdown("---")
                st.markdown("**✍️ 직접 입력 (여러 마리 한꺼번에 등록)**")
                # 같은 날 같은 시장에서 들여온 개체를 한꺼번에 넣으므로, 시험군·입식일·우시장은 한 번만 고르고
                # 표에 줄마다 개체별 내용을 넣어 한 번에 등록한다 (품목 매입 등록과 같은 방식).
                manual_group_opts = {k: v for k, v in cattle_group_opts.items() if v != "AUTO"}
                m1, m2, m3 = st.columns(3)
                with m1:
                    new_cattle_group = st.selectbox("기본 시험군", list(manual_group_opts.keys()), key="manual_cattle_group")
                with m2:
                    new_admission_date = st.date_input("기본 입식일 (구입일)", key="manual_admission_date",
                                                       help="표에서 입식일을 비워 둔 줄에 적용됩니다. 줄마다 따로 입력하면 그 값이 우선합니다.")
                with m3:
                    new_market = st.text_input("기본 우시장", placeholder="예: 순정축협(정읍)", key="manual_market",
                                               help="표에서 우시장을 비워 둔 줄에 적용됩니다. 줄마다 따로 입력하면 그 값이 우선합니다.")

                cattle_entry_ver = st.session_state.setdefault("_cattle_entry_ver", 0)
                # 주의: 줄 추가가 되는 st.data_editor 는 넘기는 표가 바뀌면 입력한 내용을 지우므로,
                # 기본 표는 세션에 한 번만 만들어 두고 등록 후에만 키 버전을 올려 비운다.
                if "_cattle_entry_base" not in st.session_state:
                    n_blank = 8
                    text_cols = ["이표번호", "시험군", "KPN", "생년월일", "입식일", "우시장", "동", "사료구분", "조사료등급", "거세일"]
                    num_cols = ["칸", "송아지 구입금액", "수수료", "운송료", "보험 가입금액", "보험료"]
                    blank = {c: pd.Series([None] * n_blank, dtype="object") for c in text_cols}
                    blank.update({c: pd.Series([None] * n_blank, dtype="float") for c in num_cols})
                    st.session_state["_cattle_entry_base"] = pd.DataFrame(blank)[
                        ["이표번호", "시험군", "KPN", "생년월일", "입식일", "우시장", "동", "칸", "사료구분", "조사료등급", "거세일",
                         "송아지 구입금액", "수수료", "운송료", "보험 가입금액", "보험료"]
                    ]
                cattle_edited = farm_data_editor(
                    st.session_state["_cattle_entry_base"],
                    placeholder="",
                    width="stretch",
                    hide_index=True,
                    num_rows="dynamic",
                    key=f"cattle_entry_editor_{cattle_entry_ver}",
                    column_config={
                        "이표번호": st.column_config.TextColumn("이표번호 (개체번호)", width="medium"),
                        "시험군": st.column_config.SelectboxColumn("시험군 (비우면 기본값)", options=list(manual_group_opts.keys()), width="medium"),
                        "KPN": st.column_config.TextColumn("KPN", width="small"),
                        "생년월일": st.column_config.DateColumn("생년월일", format="YYYY-MM-DD"),
                        "입식일": st.column_config.DateColumn("입식일 (비우면 기본값)", format="YYYY-MM-DD"),
                        "우시장": st.column_config.TextColumn("우시장 (비우면 기본값)", width="medium"),
                        "동": st.column_config.SelectboxColumn("동", options=[f"{i}동" for i in range(1, 7)], width="small"),
                        "칸": st.column_config.NumberColumn("칸번호", min_value=1, max_value=20, step=1, format="%d", width="small"),
                        "사료구분": st.column_config.SelectboxColumn("사료구분", options=["표준", "증량형", "제한형"], default="표준", width="small"),
                        "조사료등급": st.column_config.SelectboxColumn("조사료등급", options=["표준", "고급", "저급"], default="표준", width="small"),
                        "거세일": st.column_config.DateColumn("거세일", format="YYYY-MM-DD"),
                        "송아지 구입금액": st.column_config.NumberColumn("송아지 구입금액 (원)", min_value=0, format="localized", alignment="right"),
                        "수수료": st.column_config.NumberColumn("수수료 (원)", min_value=0, default=30000, format="localized", alignment="right"),
                        "운송료": st.column_config.NumberColumn("운송료 (원)", min_value=0, default=0, format="localized", alignment="right"),
                        "보험 가입금액": st.column_config.NumberColumn("보험 가입금액 (만원)", min_value=0, format="localized", alignment="right"),
                        "보험료": st.column_config.NumberColumn("보험료 (원)", min_value=0, format="localized", alignment="right"),
                    },
                )
                st.caption("이표번호를 입력한 줄만 등록됩니다. 줄이 모자라면 표 아래 빈 칸을 눌러 추가하세요. (새 줄의 수수료는 30,000원, 사료구분·조사료등급은 '표준'이 기본값)")

                def _cell(v):
                    return None if pd.isna(v) else v

                def _cell_date(v):
                    v = _cell(v)
                    if v is None:
                        return None
                    return v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else str(v)[:10]

                def _cell_num(v, default=0):
                    v = _cell(v)
                    return default if v is None else float(v)

                entry_existing = set(
                    pd.read_sql("SELECT cattle_id FROM cattle", conn)["cattle_id"].astype(str).str.strip()
                )
                cattle_rows, cattle_problems, entry_seen = [], [], set()
                for r in cattle_edited.to_dict("records"):
                    cid = clean_excel_text(_cell(r["이표번호"]))
                    if not cid:
                        if any(pd.notna(v) for k, v in r.items() if k != "이표번호"):
                            cattle_problems.append("이표번호 없이 다른 값만 입력된 줄이 있습니다.")
                        continue
                    if cid in entry_existing:
                        cattle_problems.append(f"{cid}: 이미 등록된 이표번호입니다.")
                        continue
                    if cid in entry_seen:
                        cattle_problems.append(f"{cid}: 표 안에 같은 이표번호가 두 번 있습니다.")
                        continue
                    entry_seen.add(cid)
                    calf_p, comm_f, trans_f = _cell_num(r["송아지 구입금액"]), _cell_num(r["수수료"]), _cell_num(r["운송료"])
                    pen = _cell(r["칸"])
                    cattle_rows.append((
                        cid, clean_excel_text(_cell(r["KPN"])), _cell_date(r["생년월일"]),
                        manual_group_opts[_cell(r["시험군"]) or new_cattle_group],
                        _cell_date(r["입식일"]) or new_admission_date.isoformat(),
                        clean_excel_text(_cell(r["우시장"])) or new_market, _cell(r["동"]),
                        None if pen is None else int(pen),
                        _cell(r["사료구분"]) or "표준", _cell(r["조사료등급"]) or "표준", _cell_date(r["거세일"]),
                        calf_p, comm_f, trans_f, calf_p + comm_f + trans_f,
                        _cell_num(r["보험 가입금액"], None), _cell_num(r["보험료"], None),
                    ))

                if cattle_rows:
                    show_table_total(len(cattle_rows), "구입비용 합계", sum(x[14] for x in cattle_rows))

                if st.button("개체 입식 일괄 등록", type="primary", width="stretch", key="submit_cattle_btn"):
                    if cattle_problems:
                        st.warning("등록하지 않았습니다. 아래 내용을 확인하세요.\n\n" + "\n".join(f"- {p}" for p in dict.fromkeys(cattle_problems)))
                    elif not cattle_rows:
                        st.warning("이표번호를 입력한 줄이 없습니다.")
                    else:
                        # 한 트랜잭션으로 넣어 일부만 들어가는 일이 없게 한다.
                        write_conn = db_connect(DB_FILE)
                        saved = False
                        try:
                            write_conn.executemany(
                                "INSERT INTO cattle (cattle_id, kpn, birth_date, test_group_code, status, admission_date, market_name, building, pen_number, feed_type, roughage_grade, castration_date, calf_price, commission_fee, transport_fee, initial_cost, insurance_value, insurance_premium) VALUES (?,?,?,?,'사육',?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                cattle_rows,
                            )
                            write_conn.commit()
                            saved = True
                        except sqlite3.IntegrityError as e:
                            write_conn.rollback()
                            st.error(f"등록하지 못했습니다 (한 건도 저장되지 않음): {e}")
                        finally:
                            write_conn.close()
                        if saved:
                            st.session_state["_cattle_entry_ver"] = cattle_entry_ver + 1
                            st.session_state.pop("_cattle_entry_base", None)
                            total_cost = sum(x[14] for x in cattle_rows)
                            notify(f"개체 {len(cattle_rows)}마리 입식 등록 완료! (구입비용합계: {total_cost:,.0f}원)", icon="✅")
                            st.rerun()

    
    with sub_tab2:
        col_st, col_dis = st.columns(2)
        
        with col_st:
            st.subheader("🔄 상태 변경 (출하 / 폐사)")
            active_cattle = pd.read_sql("""
                SELECT c.cattle_id, c.test_group_code, t.test_name, c.building, c.pen_number 
                FROM cattle c 
                LEFT JOIN testgroup_master t ON c.test_group_code = t.test_group_code 
                WHERE c.status = '사육' 
                ORDER BY c.cattle_id
            """, conn)
            if active_cattle.empty:
                st.info("현재 사육 중인 개체가 없습니다.")
            else:
                search_st_cid = st.text_input("🔎 대상 개체 이표번호 검색", placeholder="검색할 이표번호 일부 입력", key="search_st_cid")
                if search_st_cid:
                    active_cattle = active_cattle[active_cattle['cattle_id'].astype(str).str.contains(search_st_cid)]
                
                if active_cattle.empty:
                    st.warning("검색된 개체가 없습니다.")
                else:
                    cattle_opts = {}
                    for _, r in active_cattle.iterrows():
                        b = str(r['building']).strip() if pd.notnull(r['building']) and str(r['building']).lower() != 'nan' else ""
                        gname = r['test_name'] if pd.notna(r['test_name']) else r['test_group_code']
                        label = f"{r['cattle_id']} ({gname}) {b}".strip()
                        cattle_opts[label] = r['cattle_id']
                        
                    with st.form("change_status_form", clear_on_submit=True):
                        target_cattle_label = st.selectbox("대상 개체 선택", list(cattle_opts.keys()))
                        new_status = st.selectbox("변경할 상태", ["출하", "폐사"])
                        closure_date = st.date_input("출하/폐사 일자")
                        
                        submitted_status = st.form_submit_button("상태 변경", type="primary", width="stretch")
                            
                        if submitted_status:
                            target_id = cattle_opts[target_cattle_label]
                            wc = db_connect(DB_FILE)
                            wc.execute("UPDATE cattle SET status = ?, closure_date = ? WHERE cattle_id = ?", (new_status, closure_date.isoformat(), target_id))
                            wc.commit(); wc.close()
                            notify(f"개체 '{target_id}' → '{new_status}' 변경 완료", icon="✅")
                            st.rerun()
                    
                    st.markdown("---")
                    st.subheader("📍 개체 위치(동/우방) 및 시험군 이동")
                    st.caption("우방을 변경하면 해당 우방에 있는 개체들과 같은 시험군으로 자동 소속됩니다. 새로 만든 시험군에 기존 개체를 편입할 때도 여기서 이동하세요.")
                    
                    groups_for_move = pd.read_sql("SELECT test_group_code, test_name FROM testgroup_master", conn)
                    move_group_opts = {r['test_name']: r['test_group_code'] for _, r in groups_for_move.iterrows()}
                    
                    with st.form("move_cattle_form", clear_on_submit=True):
                        move_cattle_labels = st.multiselect("이동할 개체 선택 (다중 선택 가능)", list(cattle_opts.keys()), key="move_cattle_sel")
                        
                        m1, m2 = st.columns(2)
                        with m1:
                            move_b = st.selectbox("새로운 동", [f"{i}동" for i in range(1, 7)])
                        with m2:
                            move_p = st.number_input("새로운 우방", min_value=1, max_value=20, value=1)
                            
                        manual_group = st.selectbox("수동 시험군 지정 (자동 할당을 원치 않을 경우)", ["(자동으로 찾기)"] + list(move_group_opts.keys()))
                        
                        submitted_move = st.form_submit_button("위치 및 소속 변경", type="primary", width="stretch")
                            
                        if submitted_move:
                            if not move_cattle_labels:
                                st.warning("이동할 개체를 하나 이상 선택하세요.")
                            else:
                                wc = db_connect(DB_FILE)
                                moving_cids = [cattle_opts[lbl] for lbl in move_cattle_labels]
                                placeholders = ",".join(["?"] * len(moving_cids))
                                
                                target_gcode = None
                                if manual_group != "(자동으로 찾기)":
                                    target_gcode = move_group_opts[manual_group]
                                else:
                                    # 이동할 개체들을 제외한 목적지 우방의 기존 소들 중에서 시험군을 찾음
                                    query = f"SELECT test_group_code FROM cattle WHERE building=? AND pen_number=? AND status='사육' AND test_group_code IS NOT NULL AND cattle_id NOT IN ({placeholders}) LIMIT 1"
                                    dest_row = wc.execute(query, [move_b, move_p] + moving_cids).fetchone()
                                    if dest_row:
                                        target_gcode = dest_row[0]
                                
                                for cid in moving_cids:
                                    indiv_gcode = target_gcode
                                    if indiv_gcode is None:
                                        cur_row = wc.execute("SELECT test_group_code FROM cattle WHERE cattle_id=?", (cid,)).fetchone()
                                        if cur_row:
                                            indiv_gcode = cur_row[0]
                                            
                                    wc.execute("UPDATE cattle SET building = ?, pen_number = ?, test_group_code = ? WHERE cattle_id = ?", (move_b, move_p, indiv_gcode, cid))
                                
                                wc.commit(); wc.close()
                                
                                notify(f"개체 {len(moving_cids)}마리 → {move_b} {move_p}번 우방으로 이동 완료", icon="✅")
                                st.rerun()
        
        with col_dis:
            st.subheader("💉 질병 / 투약 기록")
            all_cattle = pd.read_sql("SELECT cattle_id FROM cattle ORDER BY cattle_id", conn)
            if all_cattle.empty:
                st.info("등록된 개체가 없습니다.")
            else:
                with st.form("add_disease_form", clear_on_submit=True):
                    dis_cattle = st.selectbox("대상 개체", all_cattle['cattle_id'].tolist())
                    dis_date = st.date_input("발병일")
                    dis_symptom = st.text_input("병명/증상", placeholder="예: 호흡기 질환")
                    
                    dc1, dc2 = st.columns(2)
                    with dc1:
                        dis_med1 = st.text_input("약품1", placeholder="예: 항생제")
                        dis_dose1 = st.number_input("수량(ml) 1", min_value=0.0, step=1.0, value=0.0)
                    with dc2:
                        dis_med2 = st.text_input("약품2")
                        dis_dose2 = st.number_input("수량(ml) 2", min_value=0.0, step=1.0, value=0.0)
                    
                    dis_vet = st.text_input("수의사", placeholder="예: 김수의")
                    dis_rx = st.text_input("처방전번호")
                    
                    submitted_dis = st.form_submit_button("질병 기록 등록", type="primary", width="stretch")
                    if submitted_dis:
                        wc = db_connect(DB_FILE)
                        wc.execute(
                            "INSERT INTO disease_record (cattle_id, onset_date, symptom, medicine1, dosage1, medicine2, dosage2, veterinarian, prescription_no) VALUES (?,?,?,?,?,?,?,?,?)",
                            (dis_cattle, dis_date.isoformat(), dis_symptom, dis_med1, dis_dose1, dis_med2, dis_dose2, dis_vet, dis_rx)
                        )
                        wc.commit(); wc.close()
                        notify("질병 기록 등록 완료", icon="✅")
                        st.rerun()
            
            st.markdown("---")
            st.markdown("##### 질병 기록 내역")
            df_disease = pd.read_sql("""
                SELECT d.record_id as ID, d.cattle_id as 이표번호, d.onset_date as 발병일, d.symptom as 병명증상,
                       d.medicine1 as 약품1, d.dosage1 as '수량(ml)', d.medicine2 as 약품2,
                       d.veterinarian as 수의사, d.recovery_date as 완치일, d.prescription_no as 처방전번호
                FROM disease_record d ORDER BY d.onset_date DESC
            """, conn)
            if df_disease.empty:
                st.info("등록된 질병 기록이 없습니다.")
            else:
                for col in ['발병일', '완치일']:
                    df_disease[col] = pd.to_datetime(df_disease[col], errors='coerce')
                farm_dataframe(
                    df_disease, width="stretch", hide_index=True,
                    column_config={
                        "ID": st.column_config.NumberColumn(width="small"),
                        "발병일": st.column_config.DateColumn(format="YYYY-MM-DD"),
                        "완치일": st.column_config.DateColumn(format="YYYY-MM-DD"),
                        "수량(ml)": st.column_config.NumberColumn(format="localized", alignment="right"),
                    },
                )
    
    with sub_tab3:

        st.subheader("📋 전체 개체 현황 (개체관리대장)")
        
        df_all_cattle = pd.read_sql("""
            SELECT c.cattle_id as 이표번호, c.kpn as KPN,
                   c.birth_date as 생년월일, c.admission_date as 입식일,
                   t.test_name as 시험군, c.status as 상태,
                   c.market_name as 우시장, c.building as 동, c.pen_number as 우방,
                   c.feed_type as 사료구분, c.roughage_grade as 조사료등급,
                   c.castration_date as 거세일,
                   c.calf_price as 구입금액, c.commission_fee as 수수료,
                   c.transport_fee as 운송료, c.initial_cost as 구입비용합계,
                   COALESCE((SELECT SUM(l.allocated_variable_cost + l.allocated_fixed_cost)
                               FROM cattle_cost_log l WHERE l.cattle_id = c.cattle_id), 0) as 누적사육비,
                   (SELECT MAX(l.settlement_month) FROM cattle_cost_log l WHERE l.cattle_id = c.cattle_id) as 원가반영월,
                   c.insurance_value as 보험가입금액, c.insurance_premium as 보험료,
                   c.closure_date as 종결일, c.memo as 비고
            FROM cattle c
            LEFT JOIN testgroup_master t ON c.test_group_code = t.test_group_code
            ORDER BY c.status, c.admission_date, c.cattle_id
        """, conn)

        # 월령 자동 계산 (태어나면 1개월 시작)
        df_all_cattle['생년월일_dt'] = pd.to_datetime(df_all_cattle['생년월일'], errors='coerce')
        df_all_cattle['입식일_dt'] = pd.to_datetime(df_all_cattle['입식일'], errors='coerce')
        df_all_cattle['종결일_dt'] = pd.to_datetime(df_all_cattle['종결일'], errors='coerce')
        
        today_dt = pd.to_datetime('today')
        
        import calendar
        def calc_months_days(start, end):
            if pd.isna(start) or pd.isna(end): return ""
            
            months = (end.year - start.year) * 12 + (end.month - start.month)
            if end.day >= start.day:
                days = end.day - start.day
            else:
                months -= 1
                if end.month == 1:
                    prev_year = end.year - 1
                    prev_month = 12
                else:
                    prev_year = end.year
                    prev_month = end.month - 1
                _, days_in_prev_month = calendar.monthrange(prev_year, prev_month)
                days = days_in_prev_month - start.day + end.day
            
            # 태어나면 1개월로 시작하므로 +1
            months += 1
            
            return f"{months}개월 {days}일"
            
        df_all_cattle['입식시월령'] = df_all_cattle.apply(lambda r: calc_months_days(r['생년월일_dt'], r['입식일_dt']), axis=1)
        df_all_cattle['현재월령'] = df_all_cattle.apply(lambda r: calc_months_days(r['생년월일_dt'], r['종결일_dt'] if pd.notna(r['종결일_dt']) else today_dt), axis=1)
        
        # 컬럼 순서 재배치
        cols_order = [
            '이표번호', 'KPN', '생년월일', '입식일', '입식시월령', '현재월령',
            '시험군', '상태', '우시장', '동', '우방', '사료구분', '조사료등급', '거세일',
            '구입금액', '수수료', '운송료', '구입비용합계', '누적사육비', '현재원가', '원가반영월',
            '보험가입금액', '보험료', '종결일', '비고'
        ]
        # 현재원가 = 구입비용합계 + 지금까지 정산(원가 적재)된 모든 달의 변동비·고정비 배분 합계
        df_all_cattle['현재원가'] = (pd.to_numeric(df_all_cattle['구입비용합계'], errors='coerce').fillna(0)
                                  + pd.to_numeric(df_all_cattle['누적사육비'], errors='coerce').fillna(0))
        df_all_cattle = df_all_cattle[cols_order]

        col_f1, col_f2, col_f3, col_f4 = st.container(key="grid2_cattle_filters").columns([1, 1, 1, 1.5])
        with col_f1:
            unique_groups = [g for g in df_all_cattle['시험군'].unique() if pd.notna(g)]
            filter_group = st.selectbox("📌 시험군 필터", ["(전체 보기)", "(미배정)"] + unique_groups)
        with col_f2:
            building_vals = set(str(b).strip() for b in df_all_cattle['동'].unique() if pd.notna(b) and str(b).strip() != '')
            unique_buildings = sorted(building_vals, key=lambda v: (0, int(v.replace("동", ""))) if v.replace("동", "").isdigit() else (1, v))
            filter_building = st.selectbox("🏢 동 필터", ["(전체 보기)", "(미배정)"] + unique_buildings)
        with col_f3:
            pen_vals = set(str(p).strip() for p in df_all_cattle['우방'].unique() if pd.notna(p) and str(p).strip() != '')
            unique_pens = sorted(pen_vals, key=lambda v: (0, int(v)) if v.isdigit() else (1, v))
            filter_pen = st.selectbox("🏠 우방 필터", ["(전체 보기)", "(미배정)"] + unique_pens)
        with col_f4:
            search_cid = st.text_input("🔎 이표번호 검색", placeholder="검색할 이표번호의 일부 또는 전체를 입력하세요...")

        if filter_group != "(전체 보기)":
            if filter_group == "(미배정)":
                df_all_cattle = df_all_cattle[df_all_cattle['시험군'].isna()]
            else:
                df_all_cattle = df_all_cattle[df_all_cattle['시험군'] == filter_group]

        if filter_building != "(전체 보기)":
            if filter_building == "(미배정)":
                df_all_cattle = df_all_cattle[df_all_cattle['동'].isna() | (df_all_cattle['동'] == '')]
            else:
                df_all_cattle = df_all_cattle[df_all_cattle['동'].astype(str).str.strip() == filter_building]

        if filter_pen != "(전체 보기)":
            if filter_pen == "(미배정)":
                df_all_cattle = df_all_cattle[df_all_cattle['우방'].isna() | (df_all_cattle['우방'] == '')]
            else:
                df_all_cattle = df_all_cattle[df_all_cattle['우방'].astype(str).str.strip() == filter_pen]

        if search_cid:
            df_all_cattle = df_all_cattle[df_all_cattle['이표번호'].astype(str).str.contains(search_cid)]
            
        status_badge = {'사육': '🟢 사육', '출하': '🔵 출하', '폐사': '🔴 폐사'}
        cnt_by_status = df_all_cattle['상태'].value_counts()
        st.markdown(
            f"**총 {len(df_all_cattle):,}건** &nbsp;·&nbsp; "
            + " &nbsp; ".join(f"{status_badge[s]} {int(cnt_by_status.get(s, 0)):,}두" for s in ('사육', '출하', '폐사'))
            + f" &nbsp;·&nbsp; 현재원가 합계 **{pd.to_numeric(df_all_cattle['현재원가'], errors='coerce').fillna(0).sum():,.0f}원**"
        )

        # 날짜는 달력 형식, 금액은 숫자형으로 두고 표시 형식만 column_config 로 지정한다.
        for col in ['생년월일', '입식일', '거세일', '종결일']:
            df_all_cattle[col] = pd.to_datetime(df_all_cattle[col], errors='coerce')
        money_cols = ['구입금액', '수수료', '운송료', '구입비용합계', '누적사육비', '현재원가', '보험가입금액', '보험료']
        for col in money_cols:
            df_all_cattle[col] = pd.to_numeric(df_all_cattle[col], errors='coerce')
        df_all_cattle['우방'] = pd.to_numeric(df_all_cattle['우방'], errors='coerce')
        df_all_cattle['상태'] = df_all_cattle['상태'].map(lambda s: status_badge.get(s, s))

        df_all_cattle.insert(0, "선택", False)
        disabled_cols = [c for c in df_all_cattle.columns if c != "선택"]
        date_cfg = {c: st.column_config.DateColumn(c, format="YYYY-MM-DD") for c in ['생년월일', '입식일', '거세일', '종결일']}
        money_cfg = {
            c: st.column_config.NumberColumn(f"{c} (원)", format="localized", alignment="right", step=1)
            for c in money_cols
        }
        edited_all_cattle_df = farm_data_editor(
            df_all_cattle,
            width="stretch",
            hide_index=True,
            disabled=disabled_cols,
            num_rows="fixed",
            key="all_cattle_move_editor",
            column_config={
                "선택": st.column_config.CheckboxColumn("선택", width="small", pinned=True,
                                                       help="체크한 개체를 아래에서 다른 우방으로 이동할 수 있습니다."),
                "이표번호": st.column_config.TextColumn("이표번호", pinned=True),
                "상태": st.column_config.TextColumn("상태", width="small"),
                "우방": st.column_config.NumberColumn("우방", format="%d", alignment="center", width="small"),
                "동": st.column_config.TextColumn("동", alignment="center", width="small"),
                "KPN": st.column_config.TextColumn("KPN", width="small"),
                **date_cfg,
                **money_cfg,
                "누적사육비": st.column_config.NumberColumn(
                    "누적 사육비 (원)", format="localized", alignment="right", step=1,
                    help="지금까지 월말 정산으로 이 개체에 배분된 변동비·고정비 합계"),
                "현재원가": st.column_config.NumberColumn(
                    "현재원가 (원)", format="localized", alignment="right", step=1,
                    help="구입비용합계 + 누적 사육비 (원가 반영월까지의 산입후 원가)"),
                "원가반영월": st.column_config.TextColumn(
                    "원가 반영월", width="small", help="마지막으로 원가가 배분된 정산 연월. 비어 있으면 아직 배분 전"),
            },
        )

        # ----- 보험료 일괄 등록: 이표번호 + 보험료(+ 가입금액) 엑셀로 개체별 보험료를 채운다 -----
        with st.expander("💰 보험료 일괄 등록 (엑셀)"):
            st.caption("엑셀에 **이표번호**와 **보험료** 열만 있으면 됩니다 (열 이름에 '이표'·'보험료'가 들어가면 자동으로 찾습니다). "
                       "'가입금액' 열이 있으면 보험가입금액도 함께 바꿉니다. 반영 전에 바뀌는 내용을 먼저 보여 줍니다.")
            ins_template = pd.DataFrame({"이표번호": ["216585979 (예시)"], "가축보험 가입금액": [770], "가축보험 보험료": [255000]})
            ins_buf = io.BytesIO()
            with pd.ExcelWriter(ins_buf, engine="xlsxwriter") as xw:
                ins_template.to_excel(xw, index=False, sheet_name="보험료")
            st.download_button("보험료 엑셀 양식 내려받기", ins_buf.getvalue(), file_name="보험료_일괄등록_양식.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="ins_template_dl")
            ins_file = st.file_uploader("보험료 엑셀 파일", type=["xlsx", "xls"], key="ins_uploader")
            if ins_file is not None:
                try:
                    ins_df = pd.read_excel(ins_file, dtype=object)
                except Exception as e:
                    st.error(f"엑셀 파일을 읽지 못했습니다: {e}")
                    ins_df = None
                if ins_df is not None:
                    ins_df.columns = [str(col).strip() for col in ins_df.columns]
                    find_col = lambda *keys: next((col for col in ins_df.columns if any(k in col for k in keys)), None)
                    id_col = find_col("이표", "개체번호")
                    prem_col = find_col("보험료")
                    val_col = find_col("가입금액")
                    if id_col is None or prem_col is None:
                        st.error("엑셀에서 '이표번호'와 '보험료' 열을 찾지 못했습니다. (파일 열: %s)" % ", ".join(ins_df.columns))
                    else:
                        def _to_number(v):
                            if v is None or (isinstance(v, float) and pd.isna(v)):
                                return None
                            text = str(v).replace(",", "").replace("원", "").strip()
                            try:
                                return float(text) if text else None
                            except ValueError:
                                return None

                        current = pd.read_sql("SELECT cattle_id, insurance_value, insurance_premium FROM cattle", conn)
                        current["cattle_id"] = current["cattle_id"].astype(str).str.strip()
                        cur_by_id = current.set_index("cattle_id")
                        plan, unknown, bad = [], [], []
                        for _, r in ins_df.iterrows():
                            cid = clean_excel_text(r[id_col])
                            if not cid or "예시" in cid:
                                continue
                            prem = _to_number(r[prem_col])
                            if prem is None:
                                bad.append(cid)
                                continue
                            if cid not in cur_by_id.index:
                                unknown.append(cid)
                                continue
                            new_val = _to_number(r[val_col]) if val_col else None
                            plan.append({
                                "이표번호": cid,
                                "현재 보험료": cur_by_id.at[cid, "insurance_premium"],
                                "새 보험료": prem,
                                "현재 가입금액": cur_by_id.at[cid, "insurance_value"],
                                "새 가입금액": new_val if new_val is not None else cur_by_id.at[cid, "insurance_value"],
                            })
                        plan_df = pd.DataFrame(plan)
                        st.info(f"반영할 개체 **{len(plan_df):,}두** · 보험료 합계 **{plan_df['새 보험료'].sum() if len(plan_df) else 0:,.0f}원**"
                                + (f" · 등록되지 않은 이표번호 {len(unknown)}건" if unknown else "")
                                + (f" · 보험료가 비었거나 숫자가 아닌 줄 {len(bad)}건" if bad else ""))
                        if unknown:
                            st.warning("이 농장에 없는 이표번호 (반영하지 않음): " + ", ".join(unknown[:30]) + (" …" if len(unknown) > 30 else ""))
                        if len(plan_df):
                            ins_money = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right", step=1)
                            farm_dataframe(plan_df, width="stretch", hide_index=True, column_config={
                                "현재 보험료": ins_money("현재 보험료 (원)"), "새 보험료": ins_money("새 보험료 (원)"),
                                "현재 가입금액": ins_money("현재 가입금액"), "새 가입금액": ins_money("새 가입금액"),
                            })
                            if st.button(f"보험료 {len(plan_df):,}두 반영", type="primary", width="stretch", key="ins_apply_btn"):
                                wc = db_connect(DB_FILE)
                                try:
                                    for p_ in plan:
                                        wc.execute("UPDATE cattle SET insurance_premium = ?, insurance_value = ? WHERE cattle_id = ?",
                                                   (p_["새 보험료"], p_["새 가입금액"], p_["이표번호"]))
                                    wc.commit()
                                finally:
                                    wc.close()
                                notify(f"보험료 {len(plan):,}두 반영 완료 (합계 {plan_df['새 보험료'].sum():,.0f}원)", icon="✅")
                                st.rerun()

        selected_move_rows = edited_all_cattle_df[edited_all_cattle_df["선택"]]
        if not selected_move_rows.empty:
            st.markdown("---")
            st.markdown(f"##### 🚚 선택한 {len(selected_move_rows)}마리 우방 이동")
            groups_for_bulk_move = pd.read_sql("SELECT test_group_code, test_name FROM testgroup_master", conn)
            bulk_move_group_opts = {r['test_name']: r['test_group_code'] for _, r in groups_for_bulk_move.iterrows()}

            with st.form("bulk_pen_move_form", clear_on_submit=False):
                bm1, bm2 = st.columns(2)
                with bm1:
                    bulk_move_b = st.selectbox("새로운 동", [f"{i}동" for i in range(1, farm_b_cnt + 1)])
                with bm2:
                    bulk_move_p = st.number_input("새로운 우방", min_value=1, max_value=farm_p_cnt, value=1)
                bulk_manual_group = st.selectbox(
                    "수동 시험군 지정 (선택 안 하면 도착 우방의 기존 시험군을 자동으로 따름)",
                    ["(자동으로 찾기)"] + list(bulk_move_group_opts.keys())
                )
                submitted_bulk_move = st.form_submit_button("선택한 개체 이동", type="primary", width="stretch")

                if submitted_bulk_move:
                    wc = db_connect(DB_FILE)
                    moving_cids = selected_move_rows['이표번호'].astype(str).tolist()
                    placeholders = ",".join(["?"] * len(moving_cids))

                    target_gcode = None
                    if bulk_manual_group != "(자동으로 찾기)":
                        target_gcode = bulk_move_group_opts[bulk_manual_group]
                    else:
                        # 이동할 개체들을 제외한 목적지 우방의 기존 소들 중에서 시험군을 찾음
                        query = f"SELECT test_group_code FROM cattle WHERE building=? AND pen_number=? AND status='사육' AND test_group_code IS NOT NULL AND cattle_id NOT IN ({placeholders}) LIMIT 1"
                        dest_row = wc.execute(query, [bulk_move_b, bulk_move_p] + moving_cids).fetchone()
                        if dest_row:
                            target_gcode = dest_row[0]

                    for cid in moving_cids:
                        indiv_gcode = target_gcode
                        if indiv_gcode is None:
                            cur_row = wc.execute("SELECT test_group_code FROM cattle WHERE cattle_id=?", (cid,)).fetchone()
                            if cur_row:
                                indiv_gcode = cur_row[0]
                        wc.execute("UPDATE cattle SET building = ?, pen_number = ?, test_group_code = ? WHERE cattle_id = ?", (bulk_move_b, bulk_move_p, indiv_gcode, cid))

                    wc.commit(); wc.close()
                    notify(f"개체 {len(moving_cids)}마리 → {bulk_move_b} {bulk_move_p}번 우방으로 이동 완료", icon="✅")
                    st.rerun()

        # ----- 개체 정보 수정: 대장에서 1마리만 체크하면 열린다 (이표번호를 잘못 넣었을 때 등) -----
        if len(selected_move_rows) == 1:
            edit_cid = str(selected_move_rows.iloc[0]["이표번호"]).strip()
            edit_cur = None
            _ec = conn.execute("SELECT * FROM cattle WHERE cattle_id = ?", (edit_cid,))
            _row = _ec.fetchone()
            if _row is not None:
                edit_cur = dict(zip([d[0] for d in _ec.description], _row))
            if edit_cur is not None:
                st.markdown("---")
                st.markdown(f"##### ✏️ 개체 정보 수정 — {edit_cid}")
                st.caption("바꿀 칸만 고치세요. 아래 '바뀌는 내용'을 확인한 뒤 저장합니다. "
                           "이표번호를 바꾸면 질병·투약 기록과 정산 기록의 이표번호도 함께 바뀝니다. "
                           "구입금액·수수료·운송료를 고치면 구입비용합계가 다시 계산되지만, 이미 끝난 월말 정산 금액은 바뀌지 않습니다.")
                ek = f"edit_{edit_cid}_"

                def _opts(options, cur):
                    options = list(options)
                    return options if (cur is None or cur in options) else options + [cur]

                def _to_date(v):
                    if v is None or str(v).strip() in ("", "None", "nan"):
                        return None
                    try:
                        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
                    except ValueError:
                        return None

                def _num(v):
                    return None if v is None or pd.isna(v) else float(v)

                edit_groups = pd.read_sql("SELECT test_group_code, test_name FROM testgroup_master", conn)
                edit_group_opts = {"(미배정)": None, **{r["test_name"]: r["test_group_code"] for _, r in edit_groups.iterrows()}}
                cur_group_label = next((k for k, v in edit_group_opts.items() if v == edit_cur["test_group_code"]), "(미배정)")

                ea, eb, ec_ = st.columns(3)
                with ea:
                    e_id = st.text_input("이표번호 (개체번호)", value=edit_cid, key=ek + "id")
                    e_birth = st.date_input("생년월일", value=_to_date(edit_cur["birth_date"]), key=ek + "birth")
                    e_building = st.selectbox("동", _opts([f"{i}동" for i in range(1, farm_b_cnt + 1)], edit_cur["building"]),
                                              index=_opts([f"{i}동" for i in range(1, farm_b_cnt + 1)], edit_cur["building"]).index(edit_cur["building"])
                                              if edit_cur["building"] else None, key=ek + "bld")
                    e_feed = st.selectbox("사료구분", _opts(["표준", "증량형", "제한형"], edit_cur["feed_type"]),
                                          index=_opts(["표준", "증량형", "제한형"], edit_cur["feed_type"]).index(edit_cur["feed_type"])
                                          if edit_cur["feed_type"] else 0, key=ek + "feed")
                    e_calf = st.number_input("송아지 구입금액 (원)", min_value=0, step=10000, value=int(_num(edit_cur["calf_price"]) or 0), key=ek + "calf")
                    e_ins_v = st.number_input("보험 가입금액 (만원)", min_value=0, step=10, value=int(_num(edit_cur["insurance_value"]) or 0), key=ek + "insv")
                with eb:
                    e_kpn = st.text_input("KPN", value=edit_cur["kpn"] or "", key=ek + "kpn")
                    e_adm = st.date_input("입식일 (구입일)", value=_to_date(edit_cur["admission_date"]), key=ek + "adm")
                    e_pen = st.number_input("우방(칸)", min_value=1, max_value=max(farm_p_cnt, 20), step=1,
                                            value=int(edit_cur["pen_number"]) if edit_cur["pen_number"] is not None and not pd.isna(edit_cur["pen_number"]) else None,
                                            key=ek + "pen")
                    e_rough = st.selectbox("조사료등급", _opts(["표준", "고급", "저급"], edit_cur["roughage_grade"]),
                                           index=_opts(["표준", "고급", "저급"], edit_cur["roughage_grade"]).index(edit_cur["roughage_grade"])
                                           if edit_cur["roughage_grade"] else 0, key=ek + "rough")
                    e_comm = st.number_input("수수료 (원)", min_value=0, step=10000, value=int(_num(edit_cur["commission_fee"]) or 0), key=ek + "comm")
                    e_ins_p = st.number_input("보험료 (원)", min_value=0, step=1000, value=int(_num(edit_cur["insurance_premium"]) or 0), key=ek + "insp")
                with ec_:
                    e_group = st.selectbox("시험군", list(edit_group_opts.keys()), index=list(edit_group_opts.keys()).index(cur_group_label), key=ek + "grp")
                    e_market = st.text_input("우시장", value=edit_cur["market_name"] or "", key=ek + "mkt")
                    e_castr = st.date_input("거세일", value=_to_date(edit_cur["castration_date"]), key=ek + "castr")
                    e_trans = st.number_input("운송료 (원)", min_value=0, step=10000, value=int(_num(edit_cur["transport_fee"]) or 0), key=ek + "trans")
                    e_memo = st.text_input("비고", value=edit_cur["memo"] or "", key=ek + "memo")

                def _iso(d):
                    return d.isoformat() if d else None

                new_vals = {
                    "cattle_id": e_id.strip(), "kpn": e_kpn.strip() or None, "birth_date": _iso(e_birth),
                    "admission_date": _iso(e_adm), "test_group_code": edit_group_opts[e_group],
                    "market_name": e_market.strip() or None, "building": e_building,
                    "pen_number": None if e_pen is None else int(e_pen),
                    "feed_type": e_feed, "roughage_grade": e_rough, "castration_date": _iso(e_castr),
                    "calf_price": e_calf, "commission_fee": e_comm, "transport_fee": e_trans,
                    "insurance_value": e_ins_v, "insurance_premium": e_ins_p, "memo": e_memo.strip() or None,
                }
                new_vals["initial_cost"] = e_calf + e_comm + e_trans
                labels = {
                    "cattle_id": "이표번호", "kpn": "KPN", "birth_date": "생년월일", "admission_date": "입식일",
                    "test_group_code": "시험군", "market_name": "우시장", "building": "동", "pen_number": "우방",
                    "feed_type": "사료구분", "roughage_grade": "조사료등급", "castration_date": "거세일",
                    "calf_price": "송아지 구입금액", "commission_fee": "수수료", "transport_fee": "운송료",
                    "initial_cost": "구입비용합계", "insurance_value": "보험 가입금액", "insurance_premium": "보험료", "memo": "비고",
                }

                def _same(a, b):
                    a = None if a is None or (not isinstance(a, str) and pd.isna(a)) else a
                    b = None if b is None or (not isinstance(b, str) and pd.isna(b)) else b
                    if a in (None, "") and b in (None, ""):
                        return True
                    try:
                        return float(a) == float(b)
                    except (TypeError, ValueError):
                        return str(a) == str(b)

                changes = {k: v for k, v in new_vals.items() if not _same(edit_cur.get(k), v)}
                if changes:
                    group_name_by_code = {v: k for k, v in edit_group_opts.items()}
                    show = lambda k, v: (group_name_by_code.get(v, v) if k == "test_group_code" else v)
                    farm_dataframe(
                        pd.DataFrame([
                            {"항목": labels[k], "현재": str(show(k, edit_cur.get(k)) if show(k, edit_cur.get(k)) is not None else "-"),
                             "변경 후": str(show(k, v) if show(k, v) is not None else "-")}
                            for k, v in changes.items()
                        ]),
                        width="stretch", hide_index=True,
                    )
                else:
                    st.caption("바뀐 내용이 없습니다.")

                if st.button("수정 내용 저장", type="primary", width="stretch", key=ek + "save", disabled=not changes):
                    if not new_vals["cattle_id"]:
                        st.warning("이표번호를 비울 수 없습니다.")
                    else:
                        wc = db_connect(DB_FILE)
                        try:
                            new_id = new_vals["cattle_id"]
                            if new_id != edit_cid:
                                if wc.execute("SELECT 1 FROM cattle WHERE cattle_id = ?", (new_id,)).fetchone():
                                    raise ValueError(f"'{new_id}' 는 이미 등록된 이표번호입니다.")
                                # 이표번호는 다른 표가 참조하므로, 새 번호로 복사 → 참조 이동 → 옛 번호 삭제 순으로 한 트랜잭션에 처리한다.
                                src = wc.execute("SELECT * FROM cattle WHERE cattle_id = ?", (edit_cid,))
                                src_cols = [d[0] for d in src.description]
                                src_row = list(src.fetchone())
                                src_row[src_cols.index("cattle_id")] = new_id
                                wc.execute(
                                    "INSERT INTO cattle (%s) VALUES (%s)" % (", ".join(src_cols), ",".join(["?"] * len(src_cols))),
                                    src_row,
                                )
                                for child in ("disease_record", "cattle_cost_log", "cattle_item_usage_log"):
                                    wc.execute(f"UPDATE {child} SET cattle_id = ? WHERE cattle_id = ?", (new_id, edit_cid))
                                wc.execute("DELETE FROM cattle WHERE cattle_id = ?", (edit_cid,))
                            sets = [k for k in changes if k != "cattle_id"]
                            if sets:
                                wc.execute(
                                    "UPDATE cattle SET %s WHERE cattle_id = ?" % ", ".join(f"{k} = ?" for k in sets),
                                    [new_vals[k] for k in sets] + [new_id],
                                )
                            wc.commit()
                        except Exception as e:
                            wc.rollback()
                            st.error(f"수정하지 못했습니다 (아무것도 바뀌지 않음): {e}")
                        else:
                            notify(f"개체 '{edit_cid}' 정보 수정 완료" + (f" → 이표번호 '{new_vals['cattle_id']}'" if new_vals["cattle_id"] != edit_cid else ""), icon="✅")
                            st.session_state.pop("all_cattle_move_editor", None)
                            st.rerun()
                        finally:
                            try:
                                wc.close()
                            except Exception:
                                pass

with tab1:
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("🐂 사육 개체 요약")
        st.caption("&nbsp;", unsafe_allow_html=True)  # 오른쪽 캡션과 높이를 맞추기 위한 빈 여백
        df_cattle = pd.read_sql("""
            SELECT c.cattle_id as 개체번호, t.test_name as 시험군, c.status as 상태, c.initial_cost as 초기원가
            FROM cattle c
            JOIN testgroup_master t ON c.test_group_code = t.test_group_code
            ORDER BY c.status, c.cattle_id
        """, conn)
        df_cattle['상태'] = df_cattle['상태'].map({'사육': '🟢 사육', '출하': '🔵 출하', '폐사': '🔴 폐사'}).fillna(df_cattle['상태'])
        farm_dataframe(
            df_cattle, width="stretch", hide_index=True,
            column_config={
                "상태": st.column_config.TextColumn(width="small"),
                "초기원가": st.column_config.NumberColumn("초기원가 (원)", format="localized", alignment="right", step=1),
            },
        )
    with col_b:
        st.subheader("📦 품목 및 재고 상태")
        st.caption("매입 시마다 이동평균단가가 자동으로 갱신됩니다.")
        df_item = pd.read_sql("SELECT item_code as 품목코드, item_name as 품목명, category as 분류, unit as 단위, current_stock as 현재재고, moving_avg_price as 이동평균단가 FROM item_master", conn)
        farm_dataframe(
            df_item, width="stretch", hide_index=True,
            column_config={
                "품목코드": st.column_config.TextColumn("코드", width="small"),
                "분류": st.column_config.TextColumn(width="small"),
                "단위": st.column_config.TextColumn(width="small"),
                "현재재고": st.column_config.NumberColumn(format="localized", alignment="right"),
                "이동평균단가": st.column_config.NumberColumn("이동평균단가 (원)", format="localized", alignment="right"),
            },
        )

with tab0:
    # 한 화면에 등록·내역·재고·품목이 몰려 있어 길고 복잡했으므로, 개체 관리 탭처럼 하위 탭으로 나눈다.
    sub_pur_entry, sub_pur_hist, sub_stock, sub_stock_count, sub_items = st.tabs(
        ["🚚 매입 등록", "📋 매입 내역", "📦 재고 현황", "🧾 재고조사표", "🏷️ 품목 관리"])

    with sub_items:
        col_item_form, col_item_list = st.columns([1, 2])
        with col_item_form:
            st.subheader("📋 품목 등록")
            st.caption("사료, 조사료, 약품 등 새 품목을 등록합니다.")
        
            # 품목코드 자동 채번 로직
            df_existing_items = pd.read_sql("SELECT item_code FROM item_master", conn)
            max_item_num = 0
            for code in df_existing_items['item_code']:
                digits = ''.join(filter(str.isdigit, str(code)))
                if digits:
                    max_item_num = max(max_item_num, int(digits))
            next_item_code = f"ITEM{max_item_num + 1}"

            with st.form("add_item_form", clear_on_submit=True):
                st.text_input("품목코드 (자동부여)", value=next_item_code, disabled=True)
                new_item_name = st.text_input("품목명", placeholder="예: TMR사료")
                new_item_category = st.selectbox("분류", ["사료", "조사료", "약품", "기타저장품"])
                new_item_unit = st.selectbox("단위", ["kg", "ml", "개", "병", "통", "포"])
                submitted_item = st.form_submit_button("품목 등록", type="primary", width="stretch")
                if submitted_item:
                    new_item_name = new_item_name.strip()
                    if new_item_name:
                        write_conn = db_connect(DB_FILE)
                        # 같은 품목을 두 번 등록하면 매입·사용량이 둘로 나뉘어 재고와 단가가 어긋난다.
                        existing = pd.read_sql("SELECT item_code AS 품목코드, item_name AS 품목명, category AS 분류, unit AS 단위 FROM item_master", write_conn)
                        dup = existing[existing["품목명"].map(normalize_item_name) == normalize_item_name(new_item_name)]
                        if not dup.empty:
                            write_conn.close()
                            st.session_state["_dup_item_alert"] = (
                                f"'{new_item_name}' 은(는) 이미 등록된 품목입니다. 등록하지 않았습니다.", dup)
                        else:
                            try:
                                # 제출 시점에 한 번 더 최신 코드를 확인하여 동시 접속 시 충돌 방지
                                m_num = 0
                                for c in existing["품목코드"]:
                                    d = ''.join(filter(str.isdigit, str(c)))
                                    if d: m_num = max(m_num, int(d))
                                final_item_code = f"ITEM{m_num + 1}"

                                write_conn.execute(
                                    "INSERT INTO item_master (item_code, item_name, category, unit, current_stock, moving_avg_price) VALUES (?, ?, ?, ?, 0, 0)",
                                    (final_item_code, new_item_name, new_item_category, new_item_unit)
                                )
                                write_conn.commit()
                                write_conn.close()
                                notify(f"품목 '{new_item_name}' ({final_item_code})이 등록되었습니다.", icon="✅")
                                st.rerun()
                            except sqlite3.IntegrityError as e:
                                write_conn.rollback(); write_conn.close()
                                st.error(f"품목 등록 중 오류가 발생했습니다: {e}")
                    else:
                        st.warning("품목명을 입력하세요.")

            # 중복 경고 팝업은 폼 밖에서 띄운다 (폼 안에서는 대화상자를 열 수 없다).
            if "_dup_item_alert" in st.session_state:
                duplicate_item_dialog(*st.session_state.pop("_dup_item_alert"))

        with col_item_list:
            st.subheader("📋 등록된 품목 · 수정 / 삭제")
            st.caption("품목명·분류·단위를 표에서 바로 고치고, 지울 품목은 '삭제'에 체크한 뒤 저장하세요.")
            df_items_all = pd.read_sql("SELECT item_code as 품목코드, item_name as 품목명, category as 분류, unit as 단위, current_stock as 현재재고, moving_avg_price as 이동평균단가 FROM item_master", conn)
            df_items_all.insert(0, "삭제", False)

            edited_item_df = farm_data_editor(
                df_items_all,
                width="stretch",
                hide_index=True,
                disabled=["품목코드", "현재재고", "이동평균단가"],
                column_config={
                    "삭제": st.column_config.CheckboxColumn("삭제", width=50),
                    "품목코드": st.column_config.TextColumn("코드", width="small"),
                    "분류": st.column_config.SelectboxColumn("분류", options=["사료", "조사료", "약품", "기타저장품"], width="small"),
                    "단위": st.column_config.SelectboxColumn("단위", options=["kg", "ml", "개", "병", "통", "포"], width="small"),
                    "현재재고": st.column_config.NumberColumn(format="localized", alignment="right", width="small"),
                    "이동평균단가": st.column_config.NumberColumn("평균단가 (원)", format="localized", alignment="right", width="small"),
                },
                num_rows="fixed",  # 품목 추가는 위 등록 폼으로만, 삭제는 '삭제' 체크박스로
                key="item_master_editor"
            )

            # 표에서 이름을 고쳐 다른 품목과 같아지는 경우도 막는다 (삭제 체크한 줄은 제외).
            kept_items = edited_item_df[~edited_item_df["삭제"].fillna(False).astype(bool) & edited_item_df["품목명"].notna()]
            kept_norm = kept_items["품목명"].map(normalize_item_name)
            dup_items = kept_items[kept_norm.duplicated(keep=False) & kept_norm.ne("")]

            if st.button("품목 수정 사항 저장", type="primary", width="stretch"):
                if not dup_items.empty:
                    duplicate_item_dialog(
                        "같은 이름의 품목이 두 개 이상 있어 저장하지 않았습니다: "
                        + ", ".join(sorted(set(dup_items["품목명"].astype(str)))),
                        dup_items[["품목코드", "품목명", "분류", "단위"]],
                    )
                else:
                    item_changed, item_deleted, _ = editor_changes(
                        df_items_all, edited_item_df, "품목코드", ["품목명", "분류", "단위"])

                    def _save_items():
                        write_conn = db_connect(DB_FILE)
                        for _, row in item_changed.iterrows():
                            write_conn.execute("UPDATE item_master SET item_name=?, category=?, unit=? WHERE item_code=?",
                                               (row['품목명'], row['분류'], row.get('단위'), row['품목코드']))
                        blocked = []
                        for code in item_deleted['품목코드']:
                            try:
                                write_conn.execute("DELETE FROM item_master WHERE item_code=?", (code,))
                            except sqlite3.IntegrityError:
                                blocked.append(code)
                        write_conn.commit()
                        write_conn.close()
                        notify(f"품목 수정 {len(item_changed)}건, 삭제 {len(item_deleted) - len(blocked)}건 저장했습니다.", icon="✅")
                        if blocked:
                            notify(f"매입·사용 내역이 있어 삭제하지 못한 품목: {', '.join(blocked)}", icon="⚠️")

                    save_with_delete_confirm("품목", item_changed, item_deleted, edited_item_df.iloc[0:0],
                                             ["품목코드", "품목명", "분류", "단위", "현재재고"], None, _save_items)

    with sub_pur_entry:
        st.subheader("🚚 매입(입고) 등록")
        st.caption("사료·조사료·약품을 매입하면 재고와 이동평균단가가 자동 갱신됩니다.")
        
        # 품목 목록과 남은 수량(매입 누계 - 사용량 누계). 단위는 품목 등록 시 정한 값을 그대로 사용한다.
        items_df = pd.read_sql("""
            SELECT i.item_code, i.item_name, i.category, i.unit,
                   COALESCE((SELECT SUM(p.quantity) FROM purchase p WHERE p.item_code = i.item_code), 0)
                 - COALESCE((SELECT SUM(u.total_usage) FROM monthly_usage u WHERE u.item_code = i.item_code), 0) AS remaining
            FROM item_master i
        """, conn)
        if items_df.empty:
            st.info("먼저 '🏷️ 품목 관리' 탭에서 품목을 등록해 주세요.")
        else:
            item_units = {r['item_code']: (r['unit'] if pd.notna(r['unit']) else "") for _, r in items_df.iterrows()}

            # 거래명세서 한 장에 여러 품목이 함께 들어오므로, 매입일자는 한 번만 고르고
            # 표에 줄마다 품목을 골라 수량·금액을 넣어 한꺼번에 등록한다 (시험군별 사용량 등록과 같은 방식).
            purchase_date = st.date_input("매입일자", key="purchase_date_input")

            # 선택지 이름에 남은 수량·단위를 붙여, 고른 뒤에도 칸에서 재고를 바로 볼 수 있게 한다.
            # 주의: 줄 추가가 되는 st.data_editor 는 넘기는 표(data)나 선택지(column_config)가 바뀌면
            # 새 표로 보고 입력한 내용을 지운다. 품목 관리 탭에서 품목을 새로 등록하면 선택지가 바뀌므로,
            # 입력 중인 내용을 품목코드 기준으로 따로 저장해 두었다가(draft) 선택지가 바뀔 때 그 내용으로 표를 다시 채운다.
            # 등록 후 표를 비울 때는 키 버전을 올리고 저장해 둔 내용도 지운다.
            purchase_name_to_code = {}
            for _, r in items_df.iterrows():
                unit = item_units[r['item_code']]
                label = f"{r['item_name']} · 남은 {float(r['remaining']):,.1f} {unit}".rstrip()
                if label in purchase_name_to_code:  # 이름·재고가 같은 품목이 겹치면 코드로 구분
                    label = f"{r['item_name']} ({r['item_code']}) · 남은 {float(r['remaining']):,.1f} {unit}".rstrip()
                purchase_name_to_code[label] = r['item_code']
            remaining_by_code = {r['item_code']: float(r['remaining']) for _, r in items_df.iterrows()}
            purchase_ver = st.session_state.setdefault("_purchase_entry_ver", 0)
            purchase_options = list(purchase_name_to_code.keys())
            if (st.session_state.get("_purchase_entry_opts") != purchase_options
                    or "_purchase_entry_base" not in st.session_state):
                code_to_label = {code: label for label, code in purchase_name_to_code.items()}
                draft = st.session_state.get("_purchase_entry_draft", [])
                rows = [
                    {"품목": code_to_label.get(code), "수량": qty, "총매입금액": amt}
                    for code, qty, amt in draft
                    if code is None or code in code_to_label  # 그사이 삭제된 품목의 줄은 버린다
                ]
                rows += [{"품목": None, "수량": None, "총매입금액": None}] * max(8 - len(rows), 0)
                st.session_state["_purchase_entry_base"] = pd.DataFrame({
                    "품목": pd.Series([r["품목"] for r in rows], dtype="object"),
                    "수량": pd.Series([r["수량"] for r in rows], dtype="float"),
                    "총매입금액": pd.Series([r["총매입금액"] for r in rows], dtype="float"),
                })
                st.session_state["_purchase_entry_opts"] = purchase_options
            purchase_edited = farm_data_editor(
                st.session_state["_purchase_entry_base"],
                placeholder="",
                width="stretch",
                hide_index=True,
                num_rows="dynamic",
                key=f"purchase_entry_editor_{purchase_ver}",
                column_config={
                    "품목": st.column_config.SelectboxColumn("품목 (남은 수량)", options=purchase_options, width="medium"),
                    "수량": st.column_config.NumberColumn("매입수량", min_value=0, format="localized", alignment="right"),
                    "총매입금액": st.column_config.NumberColumn("총매입금액 (원)", min_value=0, format="localized", alignment="right"),
                },
            )
            # 입력 중인 내용을 품목코드로 저장해 둔다 (품목 등록 등으로 선택지가 바뀌어도 위에서 되살린다).
            st.session_state["_purchase_entry_draft"] = [
                (purchase_name_to_code.get(r.품목) if pd.notna(r.품목) else None,
                 None if pd.isna(r.수량) else float(r.수량),
                 None if pd.isna(r.총매입금액) else float(r.총매입금액))
                for r in purchase_edited.itertuples(index=False)
                if pd.notna(r.품목) or pd.notna(r.수량) or pd.notna(r.총매입금액)
            ]
            st.caption("품목 칸을 눌러 목록에서 고르고 수량·금액을 입력하세요. 빈 줄은 무시되고, 줄이 모자라면 표 아래 빈 칸을 눌러 추가합니다.")
            purchase_blank_item = purchase_edited[
                purchase_edited["품목"].isna()
                & (purchase_edited["수량"].fillna(0).gt(0) | purchase_edited["총매입금액"].fillna(0).gt(0))
            ]
            purchase_input_df = purchase_edited[purchase_edited["품목"].notna()].copy()
            purchase_input_df["코드"] = purchase_input_df["품목"].map(purchase_name_to_code)
            purchase_input_df["항목"] = purchase_input_df["품목"].str.split(" · ").str[0]

            # 줄마다 단가와 등록 후 남은 수량을 미리 계산한다. 같은 품목이 여러 줄이면 앞 줄 매입분을 더해 이어서 계산.
            problems, rows_to_insert, preview = [], [], []
            if not purchase_blank_item.empty:
                problems.append(f"수량·금액만 있고 품목이 비어 있는 줄이 {len(purchase_blank_item)}개 있습니다. 품목을 고르세요.")
            for row in purchase_input_df.itertuples(index=False):
                qty = float(row.수량) if pd.notna(row.수량) else 0.0
                amt = float(row.총매입금액) if pd.notna(row.총매입금액) else 0.0
                if qty <= 0 or amt <= 0:
                    problems.append(f"{row.항목}: 수량과 금액을 0보다 크게 입력하세요.")
                    continue
                code = row.코드
                rem = remaining_by_code[code]
                remaining_by_code[code] = rem + qty
                rows_to_insert.append((purchase_date.isoformat(), code, qty, item_units.get(code) or "", amt))
                preview.append({
                    "품목": row.항목, "현재 남은 수량": rem, "매입수량": qty, "단가": round(amt / qty, 2),
                    "총매입금액": amt, "등록 후 남은 수량": rem + qty, "단위": item_units.get(code) or "",
                })

            if preview:
                money = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right", step=1)
                farm_dataframe(
                    pd.DataFrame(preview), width="stretch", hide_index=True,
                    column_config={
                        "현재 남은 수량": st.column_config.NumberColumn(format="localized", alignment="right"),
                        "매입수량": st.column_config.NumberColumn(format="localized", alignment="right"),
                        "단가": money("단가 (원)"),
                        "총매입금액": money("총매입금액 (원)"),
                        "등록 후 남은 수량": st.column_config.NumberColumn(format="localized", alignment="right"),
                    },
                )
                show_table_total(len(preview), "총매입금액", sum(r[4] for r in rows_to_insert))

            if st.button("매입 일괄 등록", type="primary", width="stretch", key="submit_purchase_btn"):
                if problems:
                    st.warning("등록하지 않았습니다. 아래 품목을 확인하세요.\n\n" + "\n".join(f"- {p}" for p in problems))
                elif not rows_to_insert:
                    st.warning("들어온 품목을 고르고 수량·금액을 입력하세요.")
                else:
                    # 한 트랜잭션으로 넣어 일부만 들어가는 일이 없게 한다. 재고·이동평균단가는 트리거가 줄마다 갱신.
                    write_conn = db_connect(DB_FILE)
                    try:
                        write_conn.executemany(
                            "INSERT INTO purchase (purchase_date, item_code, quantity, unit, total_amount) VALUES (?, ?, ?, ?, ?)",
                            rows_to_insert,
                        )
                        write_conn.commit()
                    finally:
                        write_conn.close()
                    st.session_state["_purchase_entry_ver"] = purchase_ver + 1
                    for k in ("_purchase_entry_draft", "_purchase_entry_base", "_purchase_entry_opts"):
                        st.session_state.pop(k, None)
                    total_amt = sum(r[4] for r in rows_to_insert)
                    notify(f"매입 {len(rows_to_insert)}건 등록 완료! ({purchase_date.isoformat()}, 합계 {total_amt:,.0f}원)", icon="✅")
                    st.rerun()

    with sub_pur_hist:
        st.subheader("📋 매입 내역")
        st.caption("월·분류·품목명으로 걸러 봅니다. 표에서 바로 고치거나 '삭제'에 체크한 뒤 저장하면 재고와 평균단가가 다시 계산됩니다.")
        df_purchase_all = pd.read_sql("""
            SELECT p.purchase_id as 매입ID, p.purchase_date as 매입일자,
                   p.item_code as 품목코드, i.item_name as 품목명, i.category as 분류,
                   p.quantity as 수량, p.unit as 단위, p.total_amount as 총금액
            FROM purchase p
            JOIN item_master i ON p.item_code = i.item_code
            ORDER BY p.purchase_date DESC, p.purchase_id DESC
        """, conn)
        hist_items = pd.read_sql("SELECT item_code, item_name, unit FROM item_master ORDER BY category, item_name", conn)
        hist_name_to_code = dict(zip(hist_items["item_name"], hist_items["item_code"]))
        hist_unit_by_code = dict(zip(hist_items["item_code"], hist_items["unit"]))

        pur_months = sorted({str(d)[:7] for d in df_purchase_all["매입일자"].dropna()}, reverse=True)
        this_month = datetime.now().strftime("%Y-%m")
        hist_month_opts = ["(전체)"] + pur_months
        default_month = this_month if this_month in pur_months else (pur_months[0] if pur_months else "(전체)")
        hf1, hf2, hf3 = st.columns([1, 1, 2])
        with hf1:
            hist_month = st.selectbox("📅 매입 월", hist_month_opts, index=hist_month_opts.index(default_month), key="pur_hist_month")
        with hf2:
            hist_cat = st.selectbox("🏷️ 분류", ["(전체)", "사료", "조사료", "약품", "기타저장품"], key="pur_hist_cat")
        with hf3:
            hist_search = st.text_input("🔎 품목명 검색", placeholder="품목명의 일부를 입력하세요", key="pur_hist_search")

        df_purchase = df_purchase_all
        if hist_month != "(전체)":
            df_purchase = df_purchase[df_purchase["매입일자"].astype(str).str[:7] == hist_month]
        if hist_cat != "(전체)":
            df_purchase = df_purchase[df_purchase["분류"] == hist_cat]
        if hist_search.strip():
            df_purchase = df_purchase[df_purchase["품목명"].astype(str).str.contains(hist_search.strip(), regex=False)]
        df_purchase = df_purchase.reset_index(drop=True)

        hm1, hm2, hm3 = st.columns(3)
        hm1.metric("매입 건수", f"{len(df_purchase):,}건")
        hm2.metric("품목 수", f"{df_purchase['품목코드'].nunique():,}개")
        hm3.metric("매입 금액", f"{pd.to_numeric(df_purchase['총금액'], errors='coerce').fillna(0).sum():,.0f}원")

        if not df_purchase.empty:
            with st.expander("품목별 합계 보기"):
                by_item = (df_purchase.groupby(["품목명", "분류"], as_index=False)
                           .agg(매입건수=("매입ID", "count"), 매입수량=("수량", "sum"), 단위=("단위", "first"), 매입금액=("총금액", "sum"))
                           .sort_values(["분류", "품목명"]))
                by_item["평균단가"] = (by_item["매입금액"] / by_item["매입수량"].where(by_item["매입수량"] != 0)).round(0)
                farm_dataframe(
                    by_item[["품목명", "분류", "매입건수", "매입수량", "단위", "평균단가", "매입금액"]],
                    width="stretch", hide_index=True,
                    column_config={
                        "매입건수": st.column_config.NumberColumn(width="small"),
                        "매입수량": st.column_config.NumberColumn(format="localized", alignment="right"),
                        "평균단가": st.column_config.NumberColumn("평균단가 (원)", format="localized", alignment="right", step=1),
                        "매입금액": st.column_config.NumberColumn("매입금액 (원)", format="localized", alignment="right", step=1),
                    },
                )

        df_purchase["매입일자"] = pd.to_datetime(df_purchase["매입일자"], errors="coerce")
        df_purchase["단가"] = (pd.to_numeric(df_purchase["총금액"], errors="coerce")
                             / pd.to_numeric(df_purchase["수량"], errors="coerce").where(lambda s: s != 0)).round(0)
        df_purchase = df_purchase[["매입ID", "매입일자", "품목명", "분류", "수량", "단위", "총금액", "단가", "품목코드"]]
        df_purchase.insert(0, "삭제", False)

        # 필터가 바뀌면 줄 구성이 달라지므로 편집표 키도 바꾼다 (저장 안 한 수정이 다른 줄에 붙지 않게).
        hist_sig = hashlib.md5(f"{hist_month}|{hist_cat}|{hist_search.strip()}".encode("utf-8")).hexdigest()[:10]
        purchase_editor_key = f"purchase_editor_{hist_sig}"
        # 수량·금액을 고치면 단가 칸도 바로 다시 계산해 보여 준다.
        if purchase_editor_key in st.session_state:
            for row_idx, changes in st.session_state[purchase_editor_key].get("edited_rows", {}).items():
                row_idx = int(row_idx)
                if row_idx < len(df_purchase):
                    new_amount = changes.get("총금액", df_purchase.at[row_idx, "총금액"])
                    new_qty = changes.get("수량", df_purchase.at[row_idx, "수량"])
                    if pd.notna(new_qty) and pd.notna(new_amount) and float(new_qty) != 0:
                        df_purchase.at[row_idx, "단가"] = round(float(new_amount) / float(new_qty))
            for row in st.session_state[purchase_editor_key].get("added_rows", []):
                amt, qty = row.get("총금액"), row.get("수량")
                if amt is not None and qty:
                    row["단가"] = round(float(amt) / float(qty))

        edited_purchase_df = farm_data_editor(
            df_purchase,
            width="stretch",
            hide_index=True,
            disabled=["매입ID", "분류", "단가", "품목코드"],
            num_rows="add",  # 줄 추가만 허용. 삭제는 '삭제' 체크박스로만 (줄을 빼서 지우는 일이 없게)
            key=purchase_editor_key,
            column_config={
                "삭제": st.column_config.CheckboxColumn("삭제", width=50),
                "매입ID": None,  # 저장에 쓰이지만 화면에서는 숨김
                "품목코드": None,
                "매입일자": st.column_config.DateColumn("매입일자", format="YYYY-MM-DD", width="small"),
                "품목명": st.column_config.SelectboxColumn("품목", options=list(hist_name_to_code.keys()), width="medium"),
                "분류": st.column_config.TextColumn("분류", width="small"),
                "단위": st.column_config.TextColumn("단위", width="small"),
                "수량": st.column_config.NumberColumn(format="localized", alignment="right"),
                "총금액": st.column_config.NumberColumn("총금액 (원)", format="localized", alignment="right", step=1),
                "단가": st.column_config.NumberColumn("단가 (원)", format="localized", alignment="right", step=1, width="small"),
            },
        )
        show_table_total(len(df_purchase), "총금액", float(pd.to_numeric(df_purchase["총금액"], errors="coerce").fillna(0).sum()))
        st.caption("새 매입은 '🚚 매입 등록' 탭에서 넣는 것이 편합니다. 이 표 아래 빈 줄로도 추가할 수 있습니다 (단위를 비우면 품목의 단위를 씁니다).")

        if st.button("매입 수정 사항 저장", type="primary", width="stretch"):
            pur_changed, pur_deleted, pur_added = editor_changes(
                df_purchase, edited_purchase_df, "매입ID", ["매입일자", "품목명", "수량", "단위", "총금액"])
            pur_added = pur_added[pur_added["매입일자"].notna() & pur_added["품목명"].notna()]

            def _iso_day(v):
                return pd.to_datetime(v).strftime("%Y-%m-%d")

            def _purchase_values(row):
                code = hist_name_to_code[row["품목명"]]
                unit = row.get("단위")
                if unit is None or pd.isna(unit) or not str(unit).strip():
                    unit = hist_unit_by_code.get(code) or ""
                return _iso_day(row["매입일자"]), code, row["수량"], unit, row["총금액"]

            def _save_purchases():
                write_conn = db_connect(DB_FILE)
                for _, row in pur_changed.iterrows():
                    write_conn.execute("UPDATE purchase SET purchase_date=?, item_code=?, quantity=?, unit=?, total_amount=? WHERE purchase_id=?",
                                       (*_purchase_values(row), row["매입ID"]))
                for _, row in pur_added.iterrows():
                    write_conn.execute("INSERT INTO purchase (purchase_date, item_code, quantity, unit, total_amount) VALUES (?, ?, ?, ?, ?)",
                                       _purchase_values(row))
                for mid in pur_deleted["매입ID"]:
                    write_conn.execute("DELETE FROM purchase WHERE purchase_id=?", (mid,))

                # 전체 품목 재고 및 단가 재계산
                items = pd.read_sql("SELECT item_code FROM item_master", write_conn)
                for item in items['item_code']:
                    purchases = pd.read_sql("SELECT quantity, total_amount FROM purchase WHERE item_code=? ORDER BY purchase_date ASC", write_conn, params=(item,))
                    stock = float(purchases['quantity'].sum()) if not purchases.empty else 0.0
                    total_val = float(purchases['total_amount'].sum()) if not purchases.empty else 0.0
                    avg_price = round(total_val / stock, 2) if stock > 0 else 0
                    write_conn.execute("UPDATE item_master SET current_stock=?, moving_avg_price=? WHERE item_code=?", (stock, avg_price, item))

                write_conn.commit()
                write_conn.close()
                notify(f"매입 수정 {len(pur_changed)}건, 추가 {len(pur_added)}건, 삭제 {len(pur_deleted)}건 저장 · 재고 재계산 완료", icon="✅")

            if (pur_changed["매입일자"].isna() | pur_changed["품목명"].isna()).any():
                st.warning("매입일자나 품목을 비운 줄이 있어 저장하지 않았습니다. 지우려면 '삭제'에 체크하세요.")
            else:
                save_with_delete_confirm("매입", pur_changed, pur_deleted, pur_added,
                                         ["매입일자", "품목명", "수량", "단위", "총금액"], "총금액", _save_purchases)

    with sub_stock:
        st.subheader("📦 품목별 재고 현황")
        # 남은 수량·금액 = 매입 누계 - 월말 비용 등록(시험군별 사용량) 누계
        stock_df = pd.read_sql("""
            SELECT i.item_name AS 품목명, i.category AS 분류, i.unit AS 단위,
                   COALESCE((SELECT SUM(p.quantity) FROM purchase p WHERE p.item_code = i.item_code), 0) AS 누적매입수량,
                   COALESCE((SELECT SUM(p.total_amount) FROM purchase p WHERE p.item_code = i.item_code), 0) AS 누적매입금액,
                   COALESCE((SELECT SUM(u.total_usage) FROM monthly_usage u WHERE u.item_code = i.item_code), 0) AS 누적사용량,
                   COALESCE((SELECT SUM(u.calculated_amount) FROM monthly_usage u WHERE u.item_code = i.item_code), 0) AS 누적사용금액
            FROM item_master i
            ORDER BY i.category, i.item_name
        """, conn)
        stock_df["남은수량"] = stock_df["누적매입수량"] - stock_df["누적사용량"]
        stock_df["남은금액"] = stock_df["누적매입금액"] - stock_df["누적사용금액"]
        stock_df["평균단가"] = (stock_df["남은금액"] / stock_df["남은수량"].where(stock_df["남은수량"] > 0.005)).round(0)
        last_settled = conn.execute("SELECT MAX(settlement_month) FROM cattle_cost_log").fetchone()[0]
        st.caption("남은 수량 = 매입 누계 − 월말 비용 등록의 시험군별 사용량 누계 · "
                   + (f"최근 정산 {last_settled}" if last_settled else "정산 내역 없음"))

        sf1, sf2, _ = st.columns([1, 1, 2])
        with sf1:
            stock_cat = st.selectbox("🏷️ 분류", ["(전체)", "사료", "조사료", "약품", "기타저장품"], key="stock_cat")
        with sf2:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            stock_remain_only = st.checkbox("남은 품목만 보기", value=True, key="stock_remain_only")
        view_df = stock_df
        if stock_cat != "(전체)":
            view_df = view_df[view_df["분류"] == stock_cat]
        if stock_remain_only:
            view_df = view_df[view_df["남은수량"] > 0.005]

        # 분류별 남은 금액
        cat_sum = stock_df[stock_df["남은수량"] > 0.005].groupby("분류")["남은금액"].sum()
        cat_cols = st.columns(4)
        for col, cat in zip(cat_cols, ["사료", "조사료", "약품", "기타저장품"]):
            col.metric(f"{cat} 남은 금액", f"{float(cat_sum.get(cat, 0)):,.0f}원")

        if view_df.empty:
            st.caption("표시할 품목이 없습니다.")
        else:
            num = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right",
                                                          step=1 if "(원)" in label else None)
            farm_dataframe(
                view_df[["품목명", "분류", "단위", "누적매입수량", "누적매입금액", "누적사용량", "남은수량", "평균단가", "남은금액"]],
                width="stretch", hide_index=True,
                column_config={
                    "누적매입수량": num("누적 매입수량"), "누적매입금액": num("누적 매입금액 (원)"),
                    "누적사용량": num("누적 사용량"), "남은수량": num("남은 수량"),
                    "평균단가": num("평균단가 (원)"), "남은금액": num("남은 금액 (원)"),
                },
            )
            show_table_total(len(view_df), "남은 금액", float(view_df["남은금액"].sum()))

    with sub_stock_count:
        st.subheader("🧾 재고조사표")
        st.caption("회사 재고조사표 양식(A 전월 재고 · B 매입 · C 장부상 재고 · D 실 재고 · E 급여량)으로 인쇄하거나 엑셀로 내려받습니다. "
                   "D(실 재고량)는 '월말 비용 등록'에 넣은 사용량으로 계산한 값이 미리 채워져 있으니, 실제로 센 수량과 다르면 표에서 고치세요.")
        now_m = datetime.now()
        prev_m = f"{now_m.year - 1}-12" if now_m.month == 1 else f"{now_m.year}-{now_m.month - 1:02d}"
        sc_months = pd.read_sql("""
            SELECT DISTINCT substr(purchase_date, 1, 7) AS m FROM purchase
            UNION SELECT DISTINCT settlement_month AS m FROM monthly_usage
        """, conn)["m"].dropna().astype(str).tolist()
        sc_months = sorted(set(sc_months) | {now_m.strftime("%Y-%m"), prev_m}, reverse=True)
        default_farm_label = selected_farm if "시험" in selected_farm else selected_farm.replace("농장", "시험농장")
        sc1, sc2, sc3, sc4 = st.columns(4)
        with sc1:
            sc_month = st.selectbox("📅 조사 월", sc_months, index=sc_months.index(prev_m), key="stock_count_month")
        with sc2:
            # 입력칸 키에 농장 이름을 넣어, 농장을 바꿨을 때 앞 농장에서 쓰던 값(예: 농장 표기)이 따라오지 않게 한다.
            sc_farm_label = st.text_input("농장 표기", value=default_farm_label, key=f"stock_count_farm_label_{selected_farm}")
        with sc3:
            sc_examiner = st.text_input("조사자", value="대리 배성태", key=f"stock_count_examiner_{selected_farm}")
        with sc4:
            sc_witness = st.text_input("입회자", value="팀장 신민석", key=f"stock_count_witness_{selected_farm}")

        sc_rows = stock_count_rows(DB_FILE, sc_month)
        sc_view = sc_rows.drop(columns=["코드"]).copy()
        sc_edited = farm_data_editor(
            sc_view,
            width="stretch",
            hide_index=True,
            num_rows="fixed",
            disabled=["상품명", "규격", "A", "B", "C", "등록사용량"],
            key=f"stock_count_editor_{sc_month}",
            column_config={
                "상품명": st.column_config.TextColumn("상품명", width="medium"),
                "규격": st.column_config.TextColumn("규격", width="small"),
                "A": st.column_config.NumberColumn("전월 재고 (A)", format="localized", alignment="right"),
                "B": st.column_config.NumberColumn("매입량 (B)", format="localized", alignment="right"),
                "C": st.column_config.NumberColumn("장부상 재고 (C=A+B)", format="localized", alignment="right"),
                "등록사용량": st.column_config.NumberColumn("등록된 사용량", format="localized", alignment="right",
                                                       help="월말 비용 등록(시험군별 사용량)에 넣은 이 달 사용량. 한우위탁우는 이 달 출하·폐사 두수"),
                "D": st.column_config.NumberColumn("실 재고량 (D) ✏️", min_value=0, format="localized", alignment="right",
                                                  help="실제로 센 수량을 입력하세요."),
            },
        )
        sc_final = sc_rows.copy()
        sc_final["D"] = pd.to_numeric(sc_edited["D"], errors="coerce").fillna(0).values
        sc_final["E"] = sc_final["C"] - sc_final["D"]
        # 월말 등록 탭의 원가배부 내역이 같은 달의 실 재고량(J)으로 쓴다.
        st.session_state[f"_stock_actual_{sc_month}"] = dict(zip(sc_final["코드"], sc_final["D"]))
        mismatch = sc_final[(sc_final["코드"] != "_cattle") & ((sc_final["E"] - sc_final["등록사용량"]).abs() > 0.005)]
        if not mismatch.empty:
            st.warning("급여량(E = C − D)이 '월말 비용 등록'의 사용량과 다른 품목: "
                       + ", ".join(f"{r.상품명} (급여량 {_stock_qty(r.E)} / 등록 {_stock_qty(r.등록사용량)})" for r in mismatch.itertuples(index=False))
                       + " — 원가 배분에 반영하려면 '💰 월말 등록' 탭의 사용량 등록을 맞춰 주세요.")

        sc_html = generate_stock_count_sheet(sc_final, sc_month, sc_farm_label, sc_examiner, sc_witness)
        dl1, dl2 = st.columns(2)
        dl1.download_button("🖨️ 인쇄용 파일 내려받기 (HTML)", sc_html.encode("utf-8"),
                            file_name=f"재고조사표_{selected_farm}_{sc_month}.html", mime="text/html",
                            width="stretch", key="stock_count_html_dl")
        dl2.download_button("📗 엑셀로 내려받기", stock_count_excel(sc_final, sc_month, sc_farm_label, sc_examiner, sc_witness),
                            file_name=f"재고조사표_{selected_farm}_{sc_month}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width="stretch", key="stock_count_xlsx_dl")
        st.caption("HTML 파일을 열어 '🖨️ 인쇄 / PDF 저장' 버튼을 누르면 A4로 인쇄됩니다.")
        st.markdown("###### 미리보기")
        st.iframe(sc_html, height=700)

with tab2:
    st.subheader("💰 월말 비용 등록 및 조회")
    st.markdown("월말에 재고 조사 후, 시험군별 품목 사용량과 농장 고정비를 등록합니다.")

    # 정산연월은 사용량·고정비 등록이 함께 쓴다. 바꾸면 두 등록 내역의 조회 연월과 정산 대상 연월도 같은 달로 맞춘다.
    month_col, _ = st.columns([1, 3])
    with month_col:
        cost_month = settlement_month_input(
            conn, ("monthly_usage", "monthly_fixedcost"), "정산연월 (사용량·고정비 공통)",
            "cost_month_input", ("usage_view_month", "fc_view_month"),
        )

    col_c, col_d = st.columns(2)
    
    with col_c:
        st.markdown("##### 🌾 시험군별 사용량 등록 (변동비)")
        
        # 시험군 및 품목 목록
        groups_df = pd.read_sql("SELECT test_group_code, test_name FROM testgroup_master", conn)
        # 현재재고(current_stock)는 매입 때만 늘고 사용량 등록으로는 줄지 않으므로,
        # 남은 수량 = 매입 누계 - 전체 기간 사용량 합계, 남은 금액 = 매입금액 합계 - 산출총액 합계 로 계산한다.
        items_df2 = pd.read_sql("""
            SELECT i.item_code, i.item_name, i.category, i.unit, i.moving_avg_price,
                   COALESCE((SELECT SUM(p.quantity) FROM purchase p WHERE p.item_code = i.item_code), 0) AS purchased,
                   COALESCE((SELECT SUM(p.total_amount) FROM purchase p WHERE p.item_code = i.item_code), 0) AS purchased_amt,
                   COALESCE((SELECT SUM(u.total_usage) FROM monthly_usage u WHERE u.item_code = i.item_code), 0) AS used,
                   COALESCE((SELECT SUM(u.calculated_amount) FROM monthly_usage u WHERE u.item_code = i.item_code), 0) AS used_amt
            FROM item_master i
        """, conn)
        items_df2['remaining'] = items_df2['purchased'] - items_df2['used']
        items_df2['remaining_amt'] = items_df2['purchased_amt'] - items_df2['used_amt']
        
        if groups_df.empty or items_df2.empty:
            st.info("시험군 또는 품목이 등록되어 있지 않습니다.")
        else:
            group_options = {f"{r['test_name']} ({r['test_group_code']})": r['test_group_code'] for _, r in groups_df.iterrows()}
            
            usage_month = cost_month  # 탭 맨 위의 공통 정산연월

            if "usage_groups_sel" not in st.session_state:
                st.session_state["usage_groups_sel"] = list(group_options.keys())[:1]
            usage_group_labels = st.multiselect(
                "시험군 (여러 개 선택 가능)", list(group_options.keys()), key="usage_groups_sel",
                help="여러 시험군을 고르면 입력한 사용량을 그 달 시험군별 사육일수(두수×일수) 비율로 나눠 등록합니다.",
            )
            usage_group_codes = [group_options[l] for l in usage_group_labels]
            group_name_of = dict(zip(groups_df['test_group_code'], groups_df['test_name']))

            # 여러 시험군에 나눌 때는 월말 정산(일할계산)과 같은 기준인 '그 달 사육일수 합계'(두수×일수) 비율을 쓴다.
            usage_group_ratios = {}
            if len(usage_group_codes) == 1:
                usage_group_ratios = {usage_group_codes[0]: 1.0}
            elif len(usage_group_codes) > 1 and re.fullmatch(r"\d{4}-\d{2}", usage_month.strip()):
                month_days = settlement_cattle(conn, usage_month.strip()).groupby("test_group_code")["rearing_days"].sum()
                sel_days = {g: float(month_days.get(g, 0)) for g in usage_group_codes}
                total_days = sum(sel_days.values())
                if total_days > 0:
                    usage_group_ratios = {g: d / total_days for g, d in sel_days.items() if d > 0}
                # 현재 사육두수: 지금 상태가 '사육'인 개체 수 (비율 계산에는 쓰지 않고 참고로 보여 준다)
                current_heads = dict(conn.execute(
                    "SELECT test_group_code, COUNT(*) FROM cattle WHERE status = '사육' GROUP BY test_group_code"
                ).fetchall())
                ratio_view = pd.DataFrame([{"시험군": group_name_of.get(g, g), "현재 사육두수": int(current_heads.get(g, 0)),
                                            "사육일수 (두수×일수)": d, "비율": (d / total_days if total_days > 0 else 0.0)}
                                           for g, d in sel_days.items()])
                ratio_view.loc[len(ratio_view)] = {  # 맨 아래 합계 줄 (보여 주기용)
                    "시험군": "합계", "현재 사육두수": int(ratio_view["현재 사육두수"].sum()),
                    "사육일수 (두수×일수)": total_days, "비율": 1.0 if total_days > 0 else 0.0,
                }
                farm_dataframe(
                    ratio_view,
                    width="stretch", hide_index=True,
                    column_config={
                        "현재 사육두수": st.column_config.NumberColumn("현재 사육두수 (두)", format="localized", alignment="right"),
                        "사육일수 (두수×일수)": st.column_config.NumberColumn(format="localized", alignment="right"),
                        "비율": st.column_config.NumberColumn(format="percent", alignment="right"),
                    },
                )
                zero = [group_name_of.get(g, g) for g, d in sel_days.items() if d <= 0]
                if total_days <= 0:
                    st.warning(f"{usage_month.strip()}에 선택한 시험군들의 사육 개체가 없어 나눌 수 없습니다.")
                elif zero:
                    st.caption(f"{usage_month.strip()}에 사육일수가 없는 시험군은 배분에서 빠집니다: {', '.join(zero)}")

            # 여러 품목을 표에 줄줄이 입력해 한 번에 등록한다 (시험군을 여러 개 고르면 위 비율로 나눠 등록).
            # 표에 남은 수량을 함께 보여 줘 재고를 보며 입력할 수 있게 한다.
            stock_state = {
                r['item_code']: {
                    "remaining": float(r['remaining']), "remaining_amt": float(r['remaining_amt']),
                    "fallback": float(r['moving_avg_price']), "unit": r['unit'] if pd.notna(r['unit']) else "",
                }
                for _, r in items_df2.iterrows()
            }
            # 품목은 줄마다 직접 골라 넣는다(목록에서 선택). 선택지 이름에 남은 수량·단위를 붙여,
            # 고른 뒤에도 칸에서 재고를 바로 볼 수 있게 한다.
            # 주의: st.data_editor 는 넘기는 표(data)가 바뀌면 새 표로 보고 입력한 내용을 지운다.
            # 그래서 고른 품목에 맞춰 표 안의 값을 채우지 않고, 표는 항상 같은 빈 표를 넘긴다.
            usage_name_to_code = {}
            for _, r in items_df2.iterrows():
                unit = stock_state[r['item_code']]["unit"]
                label = f"{r['item_name']} · 남은 {float(r['remaining']):,.1f} {unit}".rstrip()
                if label in usage_name_to_code:  # 이름·재고가 같은 품목이 겹치면 코드로 구분
                    label = f"{r['item_name']} ({r['item_code']}) · 남은 {float(r['remaining']):,.1f} {unit}".rstrip()
                usage_name_to_code[label] = r['item_code']
            usage_ver = st.session_state.setdefault("_usage_entry_ver", 0)
            usage_edited = farm_data_editor(
                pd.DataFrame({
                    "품목": pd.Series([None] * 8, dtype="object"),
                    "사용량": pd.Series([None] * 8, dtype="float"),
                }),
                placeholder="",
                width="stretch",
                hide_index=True,
                num_rows="dynamic",
                key=f"usage_entry_editor_{usage_ver}",
                column_config={
                    "품목": st.column_config.SelectboxColumn("품목 (남은 수량)", options=list(usage_name_to_code.keys()), width="medium"),
                    "사용량": st.column_config.NumberColumn("사용량", min_value=0, format="localized", alignment="right"),
                },
            )
            st.caption("품목 칸을 눌러 목록에서 고르고 사용량을 입력하세요. 빈 줄은 무시되고, 줄이 모자라면 표 아래 빈 칸을 눌러 추가합니다.")
            usage_blank_item = usage_edited[usage_edited["품목"].isna() & usage_edited["사용량"].fillna(0).gt(0)]
            usage_input_df = usage_edited[usage_edited["품목"].notna()].copy()
            usage_input_df["코드"] = usage_input_df["품목"].map(usage_name_to_code)
            usage_input_df["항목"] = usage_input_df["품목"].str.split(" · ").str[0]

            # 줄마다 적용단가·산출액을 미리 계산한다. 같은 품목이 여러 줄이면 앞 줄 사용분을 빼고 이어서 계산.
            # 적용단가 = 남은 재고의 평균단가(남은 금액 / 남은 수량). 매입 전체 평균단가를 쓰면
            # 앞서 싸게(또는 비싸게) 쓴 몫이 반영되지 않아, 다 쓰고 나도 산출총액 합계가 매입금액과 어긋난다.
            usage_problems, usage_rows, preview = [], [], []
            if not usage_blank_item.empty:
                usage_problems.append(f"사용량만 있고 품목이 비어 있는 줄이 {len(usage_blank_item)}개 있습니다. 품목을 고르세요.")
            for row in usage_input_df.itertuples(index=False):
                qty = float(row.사용량) if pd.notna(row.사용량) else 0.0
                if qty <= 0:
                    usage_problems.append(f"{row.항목}: 사용량을 0보다 크게 입력하세요.")
                    continue
                code = row.코드
                st_ = stock_state[code]
                rem, rem_amt = st_["remaining"], st_["remaining_amt"]
                unit_price = round(rem_amt / rem, 2) if rem > 0 else st_["fallback"]
                if rem > 0 and qty >= rem - 1e-9:
                    # 남은 재고를 모두 쓰는 경우: 남은 금액을 그대로 배정해 반올림 오차 없이 매입금액과 맞춘다.
                    amount = round(rem_amt + (qty - rem) * unit_price, 2)
                    price = round(amount / qty, 2)
                else:
                    price = float(unit_price)
                    amount = round(qty * price, 2)
                st_["remaining"], st_["remaining_amt"] = rem - qty, rem_amt - amount
                usage_rows.append((code, qty, price, amount))
                preview.append({
                    "품목": row.항목, "현재 남은 수량": rem, "사용량": qty, "적용단가": price, "산출액": amount,
                    "등록 후 남은 수량": st_["remaining"], "단위": st_["unit"],
                })

            if preview:
                money = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right", step=1)
                farm_dataframe(
                    pd.DataFrame(preview), width="stretch", hide_index=True,
                    column_config={
                        "현재 남은 수량": st.column_config.NumberColumn(format="localized", alignment="right"),
                        "사용량": st.column_config.NumberColumn(format="localized", alignment="right"),
                        "적용단가": money("적용단가 (원)"),
                        "산출액": money("산출액 (원)"),
                        "등록 후 남은 수량": st.column_config.NumberColumn(format="localized", alignment="right"),
                    },
                )
                show_table_total(len(preview), "산출액", sum(r[3] for r in usage_rows))
                short = [p_ for p_ in preview if p_["등록 후 남은 수량"] < -1e-9]
                if short:
                    st.warning("남은 수량보다 많이 쓰는 품목이 있습니다: " +
                               ", ".join(f"{p_['품목']} ({p_['등록 후 남은 수량']:,.1f} {p_['단위']})" for p_ in short) +
                               " — 매입 등록이 빠지지 않았는지 확인하세요.")

            # 품목별 사용량·산출액을 시험군 비율로 나눈다. 반올림 오차는 마지막 시험군에 몰아 합계가 정확히 맞게 한다.
            def _split(total, ratio_items):
                parts, acc = [], 0.0
                for i, (g, r) in enumerate(ratio_items):
                    v = round(total - acc, 2) if i == len(ratio_items) - 1 else round(total * r, 2)
                    acc += v
                    parts.append((g, v))
                return parts

            ratio_items = list(usage_group_ratios.items())
            usage_split_rows = []  # (시험군코드, 품목코드, 사용량, 적용단가, 산출액)
            if ratio_items:
                for code, qty, price, amount in usage_rows:
                    for (g, q_g), (_, a_g) in zip(_split(qty, ratio_items), _split(amount, ratio_items)):
                        usage_split_rows.append((g, code, q_g, round(a_g / q_g, 2) if q_g else price, a_g))
            if len(ratio_items) > 1 and usage_split_rows:
                item_name_of = dict(zip(items_df2['item_code'], items_df2['item_name']))
                st.markdown("###### 시험군별 배분 결과")
                farm_dataframe(
                    pd.DataFrame([{"시험군": group_name_of.get(g, g), "품목": item_name_of.get(c, c),
                                   "배분 사용량": q, "적용단가": p, "산출액": a} for g, c, q, p, a in usage_split_rows]),
                    width="stretch", hide_index=True,
                    column_config={
                        "배분 사용량": st.column_config.NumberColumn(format="localized", alignment="right"),
                        "적용단가": st.column_config.NumberColumn("적용단가 (원)", format="localized", alignment="right"),
                        "산출액": st.column_config.NumberColumn("산출액 (원)", format="localized", alignment="right"),
                    },
                )

            if st.button("사용량 일괄 등록", type="primary", width="stretch", key="submit_usage_btn"):
                if not re.fullmatch(r"\d{4}-\d{2}", usage_month.strip()):
                    st.warning("정산연월(YYYY-MM, 예: 2026-04)을 올바르게 입력하세요.")
                elif usage_problems:
                    st.warning("등록하지 않았습니다. 아래 품목을 확인하세요.\n\n" + "\n".join(f"- {p_}" for p_ in usage_problems))
                elif not usage_rows:
                    st.warning("사용한 품목 옆에 사용량을 입력하세요.")
                elif not usage_group_codes:
                    st.warning("시험군을 하나 이상 고르세요.")
                elif not usage_split_rows:
                    st.warning("선택한 시험군들에 이 달 사육 개체가 없어 사용량을 나눌 수 없습니다. 시험군이나 정산연월을 확인하세요.")
                else:
                    usage_month = usage_month.strip()
                    write_conn = db_connect(DB_FILE)
                    try:
                        # 한 트랜잭션으로 넣어 일부만 들어가는 일이 없게 한다.
                        write_conn.executemany(
                            "INSERT INTO monthly_usage (settlement_month, test_group_code, item_code, total_usage, applied_price, calculated_amount) VALUES (?, ?, ?, ?, ?, ?)",
                            [(usage_month, g, code, qty, price, amount) for g, code, qty, price, amount in usage_split_rows],
                        )
                        write_conn.commit()
                    finally:
                        write_conn.close()
                    st.session_state["usage_view_month"] = usage_month
                    st.session_state["_usage_entry_ver"] = usage_ver + 1
                    group_text = ", ".join(group_name_of.get(g, g) for g, _ in ratio_items)
                    notify(f"사용량 {len(usage_rows)}품목 → {len(usage_split_rows)}건 등록 완료! [{usage_month}] {group_text} | "
                           f"산출액 합계 {sum(r[4] for r in usage_split_rows):,.0f}원", icon="✅")
                    st.rerun()
        
    st.markdown("---")
    st.markdown("##### 등록된 사용 내역")
    usage_view_month = month_view_select(conn, "monthly_usage", "usage_view_month",
                                         st.session_state.get("cost_month_input") or datetime.now().strftime('%Y-%m'))
    df_usage = pd.read_sql("""
        SELECT u.usage_id as ID, u.settlement_month as 정산연월,
               t.test_name as 시험군, i.item_name as 품목명,
               u.total_usage as 사용량, u.applied_price as 적용단가,
               u.calculated_amount as 산출총액
        FROM monthly_usage u
        JOIN testgroup_master t ON u.test_group_code = t.test_group_code
        JOIN item_master i ON u.item_code = i.item_code
        WHERE u.settlement_month = ?
        ORDER BY t.test_name
    """, conn, params=(usage_view_month,))
    df_usage.insert(0, "삭제", False)

    # 편집표 키에 연월을 넣어, 달을 바꾸면 이전 달에서 하던 편집이 새 달의 같은 줄에 붙지 않게 한다.
    usage_editor_key = f"usage_editor_{usage_view_month}"
    if usage_editor_key in st.session_state:
        edits = st.session_state[usage_editor_key].get("edited_rows", {})
        for row_idx, changes in edits.items():
            row_idx = int(row_idx)
            if row_idx < len(df_usage):
                new_qty = changes.get("사용량", df_usage.at[row_idx, "사용량"])
                new_price = changes.get("적용단가", df_usage.at[row_idx, "적용단가"])
                if pd.notna(new_qty) and pd.notna(new_price):
                    df_usage.at[row_idx, "산출총액"] = round(float(new_qty) * float(new_price))
            
        added = st.session_state[usage_editor_key].get("added_rows", [])
        for row in added:
            qty = row.get("사용량", 0)
            price = row.get("적용단가", 0)
            row["산출총액"] = round(float(qty) * float(price))

    edited_usage_df = farm_data_editor(
        df_usage,
        width="stretch",
        hide_index=True,
        disabled=["ID", "시험군", "품목명", "산출총액"],
        num_rows="fixed",  # 행 추가는 위 등록 폼으로만, 삭제는 '삭제' 체크박스로
        key=usage_editor_key,
        column_config={
            "삭제": st.column_config.CheckboxColumn("삭제", width=50),
            "ID": None,
            "정산연월": None,  # 위 '조회 연월'과 같은 값만 반복되므로 숨김
            "사용량": st.column_config.NumberColumn(format="localized", alignment="right", width="small"),
            "적용단가": st.column_config.NumberColumn("적용단가 (원)", format="localized", alignment="right"),
            "산출총액": st.column_config.NumberColumn("산출총액 (원)", format="localized", alignment="right", step=1),
        },
    )
    # 삭제 체크한 행은 빼고, 표에서 고친 값(새로 추가한 행 포함)을 그대로 반영해 합산
    usage_kept = edited_usage_df[~edited_usage_df["삭제"].fillna(False).astype(bool)]
    show_table_total(
        len(usage_kept), "산출총액",
        pd.to_numeric(usage_kept["산출총액"], errors="coerce").fillna(0).sum(),
    )

    if st.button("사용 내역 수정 사항 저장", type="primary", width="stretch"):
        use_changed, use_deleted, _ = editor_changes(df_usage, edited_usage_df, "ID", ["정산연월", "사용량", "적용단가"])

        def _save_usage():
            write_conn = db_connect(DB_FILE)
            for _, row in use_changed.iterrows():
                qty = float(row['사용량'])
                price = float(row['적용단가'])
                write_conn.execute(
                    "UPDATE monthly_usage SET settlement_month=?, total_usage=?, applied_price=?, calculated_amount=? WHERE usage_id=?",
                    (row['정산연월'], qty, price, round(qty * price, 2), row['ID'])
                )
            for uid in use_deleted['ID']:
                write_conn.execute("DELETE FROM monthly_usage WHERE usage_id=?", (uid,))
            write_conn.commit()
            write_conn.close()
            notify(f"사용 내역 수정 {len(use_changed)}건, 삭제 {len(use_deleted)}건 저장했습니다.", icon="✅")

        save_with_delete_confirm("사용 내역", use_changed, use_deleted, edited_usage_df.iloc[0:0],
                                 ["ID", "정산연월", "시험군", "품목명", "사용량", "산출총액"], "산출총액", _save_usage)
    
    with col_d:
        st.markdown("##### ⚡ 농장 고정비 등록")

        fc_month = cost_month  # 탭 맨 위의 공통 정산연월

        # 여러 지출 항목을 한 번에 등록한다. 금액 칸은 천단위 콤마로 표시된다.
        fc_items = ["인건비", "전기세", "시험사양수고비", "CCTV사용료", "우수등급장려금", "가축보험료", "기타"]
        fc_input_df = item_entry_table(
            "fc_entry", fc_items, fc_items,
            {"금액": st.column_config.NumberColumn("총 청구금액 (원)", min_value=0, format="localized", alignment="right")},
            name_label="지출 항목",
        )
        if not fc_input_df.empty:
            show_table_total(len(fc_input_df), "청구금액",
                             pd.to_numeric(fc_input_df["금액"], errors="coerce").fillna(0).sum())

        if st.button("고정비 일괄 등록", type="primary", width="stretch", key="submit_fc_btn"):
            fc_rows = [(r.항목, float(r.금액)) for r in fc_input_df.itertuples(index=False)]
            if not re.fullmatch(r"\d{4}-\d{2}", fc_month.strip()):
                st.warning("정산연월(YYYY-MM, 예: 2026-04)을 올바르게 입력하세요.")
            elif not fc_rows:
                st.warning("청구된 지출 항목 옆에 금액을 입력하세요.")
            else:
                fc_month = fc_month.strip()
                write_conn = db_connect(DB_FILE)
                try:
                    write_conn.executemany(
                        "INSERT INTO monthly_fixedcost (settlement_month, expense_item, total_billed_amount) VALUES (?, ?, ?)",
                        [(fc_month, item, amt) for item, amt in fc_rows],
                    )
                    write_conn.commit()
                finally:
                    write_conn.close()
                st.session_state["fc_view_month"] = fc_month
                reset_item_entry("fc_entry")
                notify(f"고정비 {len(fc_rows)}건 등록 완료! [{fc_month}] 합계 {sum(a for _, a in fc_rows):,.0f}원", icon="✅")
                st.rerun()
        
    st.markdown("---")
    st.markdown("##### 등록된 고정비 내역 (체크박스로 삭제 가능)")
    fc_view_month = month_view_select(conn, "monthly_fixedcost", "fc_view_month",
                                      st.session_state["cost_month_input"])
    df_fixed = pd.read_sql("""
        SELECT fixed_cost_id as ID, settlement_month as 정산연월,
               expense_item as 지출항목, total_billed_amount as 총청구금액
        FROM monthly_fixedcost
        WHERE settlement_month = ?
        ORDER BY fixed_cost_id
    """, conn, params=(fc_view_month,))
    df_fixed.insert(0, "삭제", False)
        
    edited_fc_df = farm_data_editor(
        df_fixed,
        width="stretch",
        hide_index=True,
        disabled=["ID"],
        num_rows="fixed",  # 행 추가는 위 등록 폼으로만, 삭제는 '삭제' 체크박스로
        key=f"fixedcost_editor_{fc_view_month}",
        column_config={
            "삭제": st.column_config.CheckboxColumn("삭제", width=50),
            "ID": None,
            "정산연월": None,  # 위 '조회 연월'과 같은 값만 반복되므로 숨김
            "총청구금액": st.column_config.NumberColumn("총청구금액 (원)", format="localized", alignment="right", step=1),
        },
    )
    fc_kept = edited_fc_df[~edited_fc_df["삭제"].fillna(False).astype(bool)]
    show_table_total(
        len(fc_kept), "총청구금액",
        pd.to_numeric(fc_kept["총청구금액"], errors="coerce").fillna(0).sum(),
    )

    if st.button("고정비 수정 사항 저장", type="primary", width="stretch"):
        fc_changed, fc_deleted, _ = editor_changes(df_fixed, edited_fc_df, "ID", ["정산연월", "지출항목", "총청구금액"])
        fc_changed = fc_changed[fc_changed['정산연월'].notna() & (fc_changed['정산연월'].astype(str).str.strip() != "")]

        def _save_fixedcost():
            write_conn = db_connect(DB_FILE)
            for _, row in fc_changed.iterrows():
                write_conn.execute(
                    "UPDATE monthly_fixedcost SET settlement_month=?, expense_item=?, total_billed_amount=? WHERE fixed_cost_id=?",
                    (row['정산연월'], row['지출항목'], row['총청구금액'], int(row['ID']))
                )
            for mid in fc_deleted['ID']:
                write_conn.execute("DELETE FROM monthly_fixedcost WHERE fixed_cost_id=?", (int(mid),))
            write_conn.commit()
            write_conn.close()
            notify(f"고정비 수정 {len(fc_changed)}건, 삭제 {len(fc_deleted)}건 저장했습니다.", icon="✅")

        save_with_delete_confirm("고정비", fc_changed, fc_deleted, edited_fc_df.iloc[0:0],
                                 ["ID", "정산연월", "지출항목", "총청구금액"], "총청구금액", _save_fixedcost)


    st.markdown("---")
    st.markdown("##### 🖨️ 원가배부 내역 출력")
    st.caption("위 정산연월의 시험군별 사용량과 재고(전월 재고 · 매입 · 장부상 재고 · 실 재고)를 회사 양식으로 인쇄하거나 엑셀로 내려받습니다. "
               "실 재고량(J)은 '📦 품목·매입 관리 → 🧾 재고조사표'에서 같은 달로 입력한 값을 쓰고, 없으면 장부상 재고 − 사용량으로 채웁니다.")
    ca_groups, ca_rows = cost_allocation_rows(DB_FILE, cost_month)
    if ca_rows.empty:
        st.info(f"{cost_month}에 사용량·매입·재고가 있는 품목이 없습니다.")
    else:
        _fm = re.fullmatch(r"(구미)(.+)농장", selected_farm)
        default_ca_label = f"{_fm.group(1)}({_fm.group(2)})시험농장" if _fm else selected_farm.replace("농장", "시험농장")
        # 입력칸 키에 농장 이름을 넣어, 농장을 바꿨을 때 앞 농장의 제목(농장 표기)이 따라오지 않게 한다.
        ca1, ca2, ca3, ca4 = st.columns(4)
        with ca1:
            ca_farm_label = st.text_input("농장 표기", value=default_ca_label, key=f"cost_alloc_farm_label_{selected_farm}")
        # 재고조사표의 조사자·입회자처럼 작성자도 한 사람씩 칸을 둔다. 빈 칸은 인쇄물에서 빠진다.
        ca_writer_names = []
        for col, default_name, n in zip((ca2, ca3, ca4), ("신민석", "배성태", "이상욱"), (1, 2, 3)):
            with col:
                ca_writer_names.append(st.text_input(f"작성자 {n}", value=default_name,
                                                     key=f"cost_alloc_writer{n}_{selected_farm}"))
        ca_writers = ", ".join(w.strip().replace(",", " ") for w in ca_writer_names if w.strip())
        ca_diff_note = st.text_input("차액분 메모", placeholder="차액이 있으면 사유를 적으세요", key=f"cost_alloc_diff_note_{selected_farm}")
        ca_actual = st.session_state.get(f"_stock_actual_{cost_month}")
        ca_heads, _, _ = _cost_alloc_headers(ca_groups, cost_month)
        ca_vals = _cost_alloc_values(ca_rows, len(ca_groups), ca_actual)
        ca_view = pd.DataFrame(ca_vals, columns=[h.replace("<br>", " ") for h in ca_heads])
        ca_view.insert(0, "규격", ca_rows["규격"].values)
        ca_view.insert(0, "구분", ca_rows["품목"].values)
        farm_dataframe(ca_view, width="stretch", hide_index=True,
                       column_config={c: st.column_config.NumberColumn(c, format="localized", alignment="right")
                                      for c in ca_view.columns[2:]})
        ca_diff = [(r["품목"], v[-1]) for r, v in zip(ca_rows.to_dict("records"), ca_vals) if abs(v[-1]) > 0.005]
        if ca_diff:
            st.warning("차액분이 있는 품목: " + ", ".join(f"{n} ({_stock_qty(d)})" for n, d in ca_diff)
                       + " — 실 재고로 계산한 원가배부량과 등록한 사용량이 다릅니다.")
        ca_html = generate_cost_allocation_sheet(ca_groups, ca_rows, cost_month, ca_farm_label, ca_actual, ca_diff_note, ca_writers)
        cd1, cd2 = st.columns(2)
        cd1.download_button("🖨️ 인쇄용 파일 내려받기 (HTML)", ca_html.encode("utf-8"),
                            file_name=f"원가배부내역_{selected_farm}_{cost_month}.html", mime="text/html",
                            width="stretch", key="cost_alloc_html_dl")
        cd2.download_button("📗 엑셀로 내려받기",
                            cost_allocation_excel(ca_groups, ca_rows, cost_month, ca_farm_label, ca_actual, ca_diff_note, ca_writers),
                            file_name=f"원가배부내역_{selected_farm}_{cost_month}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width="stretch", key="cost_alloc_xlsx_dl")
        with st.expander("인쇄 미리보기"):
            st.iframe(ca_html, height=650)

# 월말 정산 실행은 비용 등록(💰 월말 등록)과 따로 탭으로 뺀다.
with tab_settle:
    st.subheader("🚀 월말 정산(일할계산) 실행")
    st.markdown("아래 버튼을 누르면 '💰 월말 등록' 탭에서 등록한 변동비·고정비를 분석하여, 이번 달 사육 이력이 있는 각 개체에 **실제 사육일수에 비례해(일할계산)** 변동비와 고정비를 배분합니다.")
    
    # 정산 대상 연월은 비용(사용량·고정비)이 등록됐거나 이미 정산된 연월 중에서 고른다.
    # 직접 입력하게 두면 비용이 없는 달(예: 이번 달)이 그대로 정산되는 실수가 생긴다.
    # '💰 월말 등록' 탭의 공통 정산연월을 바꾸면 여기 값도 따라 바뀐다(settlement_month_input). 처음에는 그 정산연월로 시작.
    calc_months = sorted({r[0] for r in conn.execute("""
        SELECT settlement_month FROM monthly_usage
        UNION SELECT settlement_month FROM monthly_fixedcost
        UNION SELECT settlement_month FROM cattle_cost_log
    """).fetchall() if r[0] and re.fullmatch(r"\d{4}-\d{2}", str(r[0]))}, reverse=True)
    if not calc_months:
        # 아래 탭들이 계속 그려져야 하므로 st.stop() 대신 이번 달 하나만 선택지로 둔다.
        st.info("비용(사용량·고정비)이 등록된 연월이 없습니다. '💰 월말 등록' 탭에서 비용을 먼저 등록하세요.")
        calc_months = [datetime.now().strftime('%Y-%m')]
    settled_months_set = {r[0] for r in conn.execute("SELECT DISTINCT settlement_month FROM cattle_cost_log").fetchall()}
    if "calc_target_month" not in st.session_state:
        start_month = st.session_state.get("cost_month_input")
        st.session_state["calc_target_month"] = start_month if start_month in calc_months else calc_months[0]
    elif st.session_state["calc_target_month"] not in calc_months:
        # 월말 등록 탭에서 비용이 없는 달로 바꾼 경우: 선택지에 없으니 가장 최근 등록 연월로 둔다.
        st.session_state["calc_target_month"] = calc_months[0]
    target_month = st.columns([1, 3])[0].selectbox(
        "정산 대상 연월", calc_months, key="calc_target_month",
        format_func=lambda m: f"{m}  (정산 완료)" if m in settled_months_set else m)
    
    # 정산 전 요약 미리보기 — 실제 정산과 같은 기준(settlement_cattle)으로 두수와 고정비 몫을 계산한다.
    preview_usage = pd.read_sql("""
        SELECT test_group_code, SUM(calculated_amount) as 변동비
        FROM monthly_usage WHERE settlement_month = ?
        GROUP BY test_group_code
    """, conn, params=(target_month,))
    fixed_total = conn.execute(
        "SELECT COALESCE(SUM(total_billed_amount), 0) FROM monthly_fixedcost WHERE settlement_month = ?",
        (target_month,),
    ).fetchone()[0]

    if not preview_usage.empty or fixed_total:
        preview_cattle = settlement_cattle(conn, target_month)
        total_days = preview_cattle['rearing_days'].sum()
        by_group = preview_cattle.groupby('test_group_code').agg(
            두수=('cattle_id', 'count'), 사육일수=('rearing_days', 'sum'))
        summary = (
            by_group.join(preview_usage.set_index('test_group_code'), how='outer')
            .fillna(0).reset_index()
            .merge(pd.read_sql("SELECT test_group_code, test_name as 시험군 FROM testgroup_master", conn),
                   on='test_group_code', how='left')
        )
        summary['시험군'] = summary['시험군'].fillna(summary['test_group_code'])
        # 고정비는 농장 전체 사육일수 비례로 나뉘므로, 시험군 몫도 같은 비율로 미리 계산
        summary['고정비'] = (fixed_total * summary['사육일수'] / total_days) if total_days else 0.0
        summary = summary[['시험군', '두수', '변동비', '고정비']]
        if len(summary) > 1:  # 시험군이 하나뿐이면 전체 행이 같은 숫자의 반복이라 생략
            summary.loc[len(summary)] = ['전체', summary['두수'].sum(), summary['변동비'].sum(), float(fixed_total)]
        summary['두수'] = summary['두수'].astype(int)
        summary['총합계'] = summary['변동비'] + summary['고정비']
        summary['두당평균'] = (summary['총합계'] / summary['두수'].where(summary['두수'] > 0)).fillna(0)

        st.markdown(f"**[{target_month}] 정산 대상 비용 요약** (입식 당일 제외, 실제 배분 기준 두수)")
        money = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right", step=1)
        farm_dataframe(
            summary.style.apply(
                lambda r: ['font-weight: 800; background-color: #EAF2ED' if r['시험군'] == '전체' else ''] * len(r),
                axis=1,
            ).format({"두수": "{:,} 두", "변동비": "{:,.0f}", "고정비": "{:,.0f}", "총합계": "{:,.0f}", "두당평균": "{:,.0f}"}),
            width="stretch", hide_index=True,
            column_config={
                "두수": st.column_config.TextColumn("정산 두수", alignment="right"),
                "변동비": money("변동비 합계 (원)"),
                "고정비": money("고정비 합계 (원)"),
                "총합계": money("총합계 (원)"),
                "두당평균": money("두당 평균 (원)"),
            },
        )
        if summary['두수'].iloc[-1] == 0:
            st.warning("이 달에 사육일수가 있는 개체가 없어 정산할 수 없습니다.")
    else:
        st.info(f"[{target_month}] 에 등록된 사용량·고정비가 없습니다. 정산 전에 '💰 월말 등록' 탭에서 비용을 먼저 등록하세요.")
    
    st.caption("※ 입식 당일은 절식하므로 배분에서 제외하고, 입식 다음 날부터 사육일수로 계산합니다.")
    settled_count = conn.execute(
        "SELECT COUNT(*) FROM cattle_cost_log WHERE settlement_month = ?", (target_month,)
    ).fetchone()[0]

    if settled_count == 0:
        if st.button("🚀 정산 실행(일할계산) 및 누적원가 반영", type="primary", width="stretch"):
            success, msg = distribute_monthly_costs(DB_FILE, target_month)
            if success:
                st.success(msg)
                st.balloons()
            else:
                st.error(msg)
    else:
        st.warning(
            f"[{target_month}] 은 이미 {settled_count}마리로 정산되어 있습니다. 비용이나 개체 정보를 고쳤다면 "
            "기존 배분 내역을 지우고 다시 정산하세요. (실행 직전 자동 백업)"
        )
        confirm_resettle = st.checkbox(f"[{target_month}] 기존 배분 내역 {settled_count}건을 지우고 다시 정산합니다",
                                       key=f"resettle_confirm_{target_month}")
        if st.button("🔄 기존 정산 지우고 다시 정산", type="secondary", width="stretch",
                     disabled=not confirm_resettle, key=f"danger_resettle_{target_month}"):
            backup_db(DB_FILE, f"before-resettle-{target_month}")
            success, msg = distribute_monthly_costs(DB_FILE, target_month, replace=True)
            if success:
                notify(f"재정산 완료 — {msg}", icon="✅")
                st.rerun()
            else:
                st.error(msg)

        # 잘못 정산한 달(예: 비용이 없는 달)은 배분 내역을 통째로 지운다. 등록된 비용(사용량·고정비)은 그대로 둔다.
        with st.expander(f"🗑️ [{target_month}] 정산 취소 (배분 내역 삭제)"):
            st.caption("이 달에 개체별로 배분된 원가 기록만 지웁니다. 등록된 사용량·고정비는 지우지 않으며, 실행 직전 자동 백업됩니다.")
            confirm_cancel = st.checkbox(f"[{target_month}] 정산 {settled_count}마리 배분 내역을 삭제합니다", key=f"cancel_settle_confirm_{target_month}")
            if st.button("정산 취소", disabled=not confirm_cancel, width="stretch", key=f"danger_cancel_settle_{target_month}"):
                backup_db(DB_FILE, f"before-cancel-settle-{target_month}")
                write_conn = db_connect(DB_FILE)
                write_conn.execute("DELETE FROM cattle_item_usage_log WHERE settlement_month = ?", (target_month,))
                write_conn.execute("DELETE FROM cattle_cost_log WHERE settlement_month = ?", (target_month,))
                write_conn.commit()
                write_conn.close()
                notify(f"[{target_month}] 정산을 취소했습니다. (배분 내역 {settled_count}건 삭제)", icon="🗑️")
                st.rerun()

    st.markdown("---")
    st.subheader(f"📒 [{target_month}] 개체별 원가 적재 결과")
    try:
        # 누적 원가 = 구입비용합계(initial_cost) + 정산 대상 연월까지 적재된 변동비·고정비 합계
        df_log = pd.read_sql("""
            SELECT l.cattle_id, l.settlement_month,
                   l.allocated_variable_cost AS 변동비_할당,
                   l.allocated_fixed_cost AS 고정비_할당,
                   (l.allocated_variable_cost + l.allocated_fixed_cost) AS 당월_추가원가,
                   COALESCE(c.initial_cost, 0) AS 구입원가,
                   (SELECT SUM(h.allocated_variable_cost + h.allocated_fixed_cost)
                      FROM cattle_cost_log h
                     WHERE h.cattle_id = l.cattle_id AND h.settlement_month <= l.settlement_month) AS 누적_사육비
            FROM cattle_cost_log l
            LEFT JOIN cattle c ON c.cattle_id = l.cattle_id
            WHERE l.settlement_month = ?
            ORDER BY l.cattle_id
        """, conn, params=(target_month,))
        if df_log.empty:
            st.info("해당 연월에 아직 정산된 내역이 없습니다.")
        else:
            df_log['누적_원가'] = df_log['구입원가'] + df_log['누적_사육비']
            money = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right", step=1)
            farm_dataframe(
                df_log, width="stretch", hide_index=True,
                column_config={
                    "cattle_id": st.column_config.TextColumn("이표번호", pinned=True),
                    "settlement_month": None,  # 제목의 연월과 같은 값만 반복되므로 숨김
                    "변동비_할당": money("변동비 할당 (원)"),
                    "고정비_할당": money("고정비 할당 (원)"),
                    "당월_추가원가": money("당월 추가원가 (원)"),
                    "구입원가": money("구입원가 (원)"),
                    "누적_사육비": money("누적 사육비 (원)"),
                    "누적_원가": money("누적 원가 (원)"),
                },
            )
            st.caption("누적 사육비 = 정산 대상 연월까지 적재된 변동비·고정비 합계 · 누적 원가 = 구입원가(구입비용합계) + 누적 사육비")
            show_table_total(len(df_log), "누적 원가", df_log['누적_원가'].sum())

            # 두당 평균 원가 (이번 정산 대상 개체 기준)
            st.markdown(f"##### 📊 [{target_month}] 두당 평균 원가 ({len(df_log):,}두)")
            avg_cols = st.container(key="kpi_avg").columns(4)
            avg_cols[0].metric("평균 당월 추가원가", f"{df_log['당월_추가원가'].mean():,.0f}원")
            avg_cols[1].metric("평균 구입원가", f"{df_log['구입원가'].mean():,.0f}원")
            avg_cols[2].metric("평균 누적 사육비", f"{df_log['누적_사육비'].mean():,.0f}원")
            avg_cols[3].metric("평균 누적 원가", f"{df_log['누적_원가'].mean():,.0f}원")

            # 시험군별 두당 평균 원가
            group_of = pd.read_sql(
                "SELECT c.cattle_id, COALESCE(t.test_name, '(시험군 없음)') AS 시험군 FROM cattle c "
                "LEFT JOIN testgroup_master t ON t.test_group_code = c.test_group_code", conn)
            by_group = (
                df_log.merge(group_of, on="cattle_id", how="left")
                .fillna({"시험군": "(시험군 없음)"})
                .groupby("시험군")
                .agg(두수=("cattle_id", "count"), 평균_당월_추가원가=("당월_추가원가", "mean"),
                     평균_구입원가=("구입원가", "mean"), 평균_누적_사육비=("누적_사육비", "mean"),
                     평균_누적_원가=("누적_원가", "mean"))
                .reset_index()
            )
            if len(by_group) > 1:
                farm_dataframe(
                    by_group, width="stretch", hide_index=True,
                    column_config={
                        "두수": st.column_config.NumberColumn("두수", format="localized", alignment="right"),
                        "평균_당월_추가원가": money("평균 당월 추가원가 (원)"),
                        "평균_구입원가": money("평균 구입원가 (원)"),
                        "평균_누적_사육비": money("평균 누적 사육비 (원)"),
                        "평균_누적_원가": money("평균 누적 원가 (원)"),
                    },
                )

            # 엑셀 저장: 개체별 원가 / 두당 평균 / 시험군별 평균(시험군이 둘 이상일 때)
            excel_cattle = (
                df_log.merge(group_of, on="cattle_id", how="left")
                .fillna({"시험군": "(시험군 없음)"})
                [["cattle_id", "시험군", "settlement_month", "변동비_할당", "고정비_할당", "당월_추가원가",
                  "구입원가", "누적_사육비", "누적_원가"]]
                .rename(columns={
                    "cattle_id": "이표번호", "settlement_month": "정산연월", "변동비_할당": "변동비 할당 (원)",
                    "고정비_할당": "고정비 할당 (원)", "당월_추가원가": "당월 추가원가 (원)", "구입원가": "구입원가 (원)",
                    "누적_사육비": "누적 사육비 (원)", "누적_원가": "누적 원가 (원)",
                })
            )
            excel_avg = pd.DataFrame([{
                "정산연월": target_month, "두수": len(df_log),
                "평균 당월 추가원가 (원)": df_log['당월_추가원가'].mean(), "평균 구입원가 (원)": df_log['구입원가'].mean(),
                "평균 누적 사육비 (원)": df_log['누적_사육비'].mean(), "평균 누적 원가 (원)": df_log['누적_원가'].mean(),
            }])
            excel_group = by_group.rename(columns={
                "평균_당월_추가원가": "평균 당월 추가원가 (원)", "평균_구입원가": "평균 구입원가 (원)",
                "평균_누적_사육비": "평균 누적 사육비 (원)", "평균_누적_원가": "평균 누적 원가 (원)",
            })
            xlsx_buf = io.BytesIO()
            with pd.ExcelWriter(xlsx_buf, engine="xlsxwriter") as xw:
                sheets = [("개체별 원가", excel_cattle), ("두당 평균", excel_avg)]
                if len(by_group) > 1:
                    sheets.append(("시험군별 평균", excel_group))
                won = xw.book.add_format({"num_format": "#,##0"})
                for sheet_name, sheet_df in sheets:
                    sheet_df.to_excel(xw, sheet_name=sheet_name, index=False)
                    ws = xw.sheets[sheet_name]
                    for i, col in enumerate(sheet_df.columns):
                        width = max(12, min(40, int(max(len(str(col)) * 1.8, sheet_df[col].astype(str).str.len().max() * 1.2)) + 2))
                        ws.set_column(i, i, width, won if pd.api.types.is_numeric_dtype(sheet_df[col]) else None)
                    ws.freeze_panes(1, 0)
            st.download_button(
                "📥 엑셀로 저장", xlsx_buf.getvalue(),
                file_name=f"{selected_farm}_개체별원가_{target_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch", key="cost_log_xlsx",
            )
    except:
        st.info("아직 정산된 내역이 없습니다.")

with tab_report:
    st.subheader("🧾 결산 리포트 생성")
    st.markdown("정산이 완료된 연월을 골라 인쇄·저장 가능한 결산서(HTML)를 만듭니다. 브라우저에서 열어 **PDF 인쇄 및 저장** 버튼으로 PDF로도 저장할 수 있습니다.")

    settled_months = list_settled_months(DB_FILE)
    if not settled_months:
        st.info("아직 정산된 연월이 없습니다. 먼저 '🚀 월말 정산' 탭에서 정산을 실행하세요.")
    else:
        report_col, _ = st.columns([1, 3])
        report_month = report_col.selectbox("리포트 연월", settled_months, key="report_month_select")
        if report_col.button("📄 결산 리포트 생성", type="primary", key="gen_report_btn", width="stretch"):
            ok, result = generate_settlement_report(DB_FILE, selected_farm, report_month)
            if ok:
                st.session_state["report_html"] = result
                st.session_state["report_month_generated"] = report_month
            else:
                st.session_state.pop("report_html", None)
                st.error(result)

        report_html = st.session_state.get("report_html")
        if report_html and st.session_state.get("report_month_generated") == report_month:
            st.success(f"'{report_month}' 결산 리포트가 생성되었습니다.")
            st.download_button(
                "⬇️ HTML 파일로 내려받기",
                report_html.encode("utf-8"),
                file_name=f"결산리포트_{selected_farm}_{report_month}.html",
                mime="text/html",
                width="stretch",
            )
            st.markdown("###### 미리보기")
            st.iframe(report_html, height=900)

    st.markdown("---")
    st.subheader("🖨️ 월말 팔공 입력 자료")
    st.markdown("회사 회계 프로그램에 옮겨 적기 위한 인쇄용 자료입니다. 매입 내역 · 품목 수불부(기초·입고·출고·기말) · 고정비 · "
                "시험군별 원가 배분 · 입식/출하/폐사 · 월말 사육 장부가를 A4 한 묶음으로 만들고, 줄마다 **입력확인(□)** 칸이 있습니다.")
    acct_months = list_accounting_months(DB_FILE)
    if not acct_months:
        st.info("아직 매입·비용·입식 자료가 없습니다.")
    else:
        acct_col, _ = st.columns([1, 3])
        acct_month = acct_col.selectbox("대상 월", acct_months, key="acct_month_select")
        if acct_col.button("🖨️ 팔공 입력 자료 만들기", type="primary", key="gen_acct_btn", width="stretch"):
            ok, result = generate_accounting_sheet(DB_FILE, selected_farm, acct_month)
            if ok:
                st.session_state["acct_html"] = result
                st.session_state["acct_month_generated"] = acct_month
            else:
                st.session_state.pop("acct_html", None)
                st.error(result)

        acct_html = st.session_state.get("acct_html")
        if acct_html and st.session_state.get("acct_month_generated") == acct_month:
            st.success(f"'{acct_month}' 팔공 입력 자료를 만들었습니다. 내려받은 파일을 열어 **🖨️ 인쇄 / PDF 저장** 버튼을 누르세요.")
            st.download_button(
                "⬇️ 인쇄용 파일 내려받기 (HTML)",
                acct_html.encode("utf-8"),
                file_name=f"팔공입력자료_{selected_farm}_{acct_month}.html",
                mime="text/html",
                width="stretch",
                key="acct_download",
            )
            st.markdown("###### 미리보기")
            st.iframe(acct_html, height=900)

with tab_slaughter:
    st.markdown(
        '<div class="farm-coming"><h3>🥩 도축 성적 관리 <span class="farm-badge">준비중</span></h3>'
        '<p>출하(도축)된 개체의 도축 성적을 기록하고 시험군별로 비교하는 메뉴입니다. 아래 항목이 들어갈 예정입니다.</p>'
        '<div class="farm-coming-grid">'
        '<div class="farm-coming-card"><b>등급</b><span>육질·육량 등급 (1++, 1+, 1, 2, 3 / A·B·C)</span></div>'
        '<div class="farm-coming-card"><b>도체중</b><span>kg, 시험군별 평균 비교</span></div>'
        '<div class="farm-coming-card"><b>등심단면적</b><span>㎠, 개체·시험군별 분포</span></div>'
        '<div class="farm-coming-card"><b>근내지방도</b><span>No.1~9, 등급 판정 근거</span></div>'
        '</div></div>',
        unsafe_allow_html=True,
    )

conn.close()
