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

st.set_page_config(page_title="대구축협 시험농장 관리 시스템", layout="wide", page_icon="🐮")

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
    padding-top: 3.6rem !important;
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
    flex-wrap: wrap;
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
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)


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
    return st.selectbox("조회 연월", months, key=key)


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
    edited = st.data_editor(
        df, width="stretch", hide_index=True, num_rows="fixed",
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
    st.dataframe(dup_df, width="stretch", hide_index=True)
    st.caption("품목명은 띄어쓰기·대소문자를 무시하고 비교합니다. 다른 품목이면 이름을 구분되게 바꿔 주세요.")
    if st.button("확인", type="primary", width="stretch"):
        st.rerun()


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
    db_adapter.ensure_schema(db_file, PG_DDL, extra_sql=[
        "ALTER TABLE testgroup_master ADD COLUMN IF NOT EXISTS location_mapping TEXT",
        "ALTER TABLE purchase ADD COLUMN IF NOT EXISTS unit TEXT",
        "ALTER TABLE item_master ADD COLUMN IF NOT EXISTS unit TEXT",
    ])

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
    <div style="display:flex; align-items:center; justify-content:center; gap:10px; padding: 10px 0 20px 0;">
        <img src="{MEDAL_ICON_DATA_URI}" alt="심볼" style="height: 40px;">
        <span style="font-size: 2.5rem;">🐮</span>
        <h2 style="margin:0; font-size:1.2rem; text-align:left;">대구축협 시험농장<br>관리 시스템</h2>
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


with st.sidebar.expander("🐄 등록된 개체 전체 삭제"):
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
        <div style="display:flex; align-items:center; gap:14px; margin-bottom:0.4rem;">
            <img src="{MEDAL_ICON_DATA_URI}" alt="심볼" style="height: 44px;">
            <div>
                <h1 style="margin:0; padding:0; font-size:1.9rem;">대구축협 시험농장 관리 시스템</h1>
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

        m1, m2, m3, m4, m5 = st.columns(5)
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
        money_col = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right")

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
        st.dataframe(
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
            st.dataframe(
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
            st.dataframe(
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
    <div style="display:flex; align-items:center; justify-content:space-between; gap:16px; flex-wrap:wrap; margin-bottom:0.4rem;">
        <div style="display:flex; align-items:center; gap:14px;">
            <img src="{MEDAL_ICON_DATA_URI}" alt="심볼" style="height: 44px;">
            <div>
                <h1 style="margin:0; padding:0; font-size:1.9rem;">대구축협 시험농장 관리 시스템</h1>
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

kpi_cols = st.columns(6)
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

tab_cattle, tab1, tab0, tab2, tab_report, tab_slaughter = st.tabs(["🐄 개체 관리", "📊 사육 및 재고 현황", "📦 품목·매입 관리", "💰 월말 정산 및 청구 내역", "🧾 결산 리포트", "🥩 도축 성적"])

# ===== 개체 관리 탭 =====
with tab_cattle:
    sub_tab1, sub_tab2, sub_tab3 = st.tabs(["🐮 입식 등록", "📋 상태 변경 / 질병 기록", "📊 전체 현황"])
    
    with sub_tab1:
        col_reg1, col_reg2 = st.columns(2)
        
        with col_reg1:
            st.subheader("시험군 등록")
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

                # 버튼 넓이를 절반으로 줄이기 위해 컬럼 사용
                btn_g1, btn_g2 = st.columns(2)
                with btn_g1:
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
            st.caption("기존 개체를 이 시험군으로 배정하려면 '상태 변경 / 질병 기록' 탭의 '개체 위치 및 시험군 이동'을 이용하세요.")

            st.markdown("---")
            st.markdown("##### 등록된 시험군")
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
                
            if '자동할당조건' in df_groups.columns:
                df_groups['자동할당조건'] = df_groups['자동할당조건'].apply(format_loc)
            
            # 시험군코드는 UI 화면 테이블에서 숨김 처리
            st.dataframe(df_groups[['시험명칭', '시작일', '종료일', '자동할당조건']], width="stretch", hide_index=True)
            
            if not df_groups.empty:
                st.markdown("##### 📝 시험군 수정 / 삭제")
                
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
                st.caption("기존 개체를 다른 시험군으로 옮기려면 '상태 변경 / 질병 기록' 탭의 '개체 위치 및 시험군 이동'을 이용하세요.")

        with col_reg2:
            st.subheader("개체 입식 등록")
            st.caption("개체를 한 마리씩 등록하거나, 엑셀 파일을 통해 일괄 등록할 수 있습니다.")
            
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

                            # 버튼 넓이를 절반으로 줄이기 위해 컬럼 사용
                            btn_b1, btn_b2 = st.columns(2)
                            with btn_b1:
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
                with st.form("add_cattle_form", clear_on_submit=True):
                    c1, c2 = st.columns(2)
                    with c1:
                        new_cattle_id = st.text_input("이표번호 (개체번호)", placeholder="예: 216585979")
                    with c2:
                        new_kpn = st.text_input("KPN", placeholder="예: 1654")
                    
                    c3, c4 = st.columns(2)
                    with c3:
                        new_birth_date = st.date_input("생년월일")
                    with c4:
                        new_admission_date = st.date_input("입식일 (구입일)")
                    
                    c5, c6 = st.columns(2)
                    with c5:
                        new_cattle_group = st.selectbox("소속 시험군", list(cattle_group_opts.keys()))
                    with c6:
                        new_market = st.text_input("우시장", placeholder="예: 순정축협(정읍)")
                    
                    c7, c8, c9 = st.columns(3)
                    with c7:
                        new_building = st.selectbox("동", [f"{i}동" for i in range(1, 7)])
                    with c8:
                        new_pen = st.number_input("우방 (칸번호)", min_value=1, max_value=20, value=1)
                    with c9:
                        new_feed_type = st.selectbox("사료구분", ["표준", "증량형", "제한형"])
                    
                    c10, c11 = st.columns(2)
                    with c10:
                        new_roughage = st.selectbox("조사료등급", ["표준", "고급", "저급"])
                    with c11:
                        new_castration = st.date_input("거세일")
                    
                    st.markdown("**입식 비용 내역**")
                    cc1, cc2, cc3 = st.columns(3)
                    with cc1:
                        new_calf_price = st.number_input("송아지 구입금액 (원)", min_value=0, step=100000, value=5000000)
                    with cc2:
                        new_commission = st.number_input("수수료 (원)", min_value=0, step=10000, value=30000)
                    with cc3:
                        new_transport = st.number_input("운송료 (원)", min_value=0, step=10000, value=0)
                    
                    st.markdown("**가축보험 정보**")
                    ci1, ci2 = st.columns(2)
                    with ci1:
                        new_ins_value = st.number_input("가입금액 (만원)", min_value=0, step=10, value=770)
                    with ci2:
                        new_ins_premium = st.number_input("보험료 (원)", min_value=0, step=1000, value=0)
                    
                    # 버튼 넓이를 절반으로 줄이기 위해 컬럼 사용
                    btn_c1, btn_c2 = st.columns(2)
                    with btn_c1:
                        submitted_cattle = st.form_submit_button("개체 입식 등록", type="primary", width="stretch")
                        
                    if submitted_cattle:
                        if new_cattle_id:
                            total_init_cost = new_calf_price + new_commission + new_transport
                            sel_group_code = cattle_group_opts[new_cattle_group]
                            try:
                                wc = db_connect(DB_FILE)
                                wc.execute(
                                    "INSERT INTO cattle (cattle_id, kpn, birth_date, test_group_code, status, admission_date, market_name, building, pen_number, feed_type, roughage_grade, castration_date, calf_price, commission_fee, transport_fee, initial_cost, insurance_value, insurance_premium) VALUES (?,?,?,?,'사육',?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                    (new_cattle_id, new_kpn, new_birth_date.isoformat(), sel_group_code, new_admission_date.isoformat(), new_market, new_building, new_pen, new_feed_type, new_roughage, new_castration.isoformat(), new_calf_price, new_commission, new_transport, total_init_cost, new_ins_value, new_ins_premium)
                                )
                                wc.commit(); wc.close()
                                notify(f"개체 '{new_cattle_id}' 입식 등록 완료! (구입비용합계: {total_init_cost:,}원)", icon="✅")
                                st.rerun()
                            except sqlite3.IntegrityError:
                                wc.rollback(); wc.close()
                                st.error("이미 등록된 이표번호입니다.")
                        else:
                            st.warning("이표번호를 입력하세요.")

    
    with sub_tab2:
        col_st, col_dis = st.columns(2)
        
        with col_st:
            st.subheader("상태 변경 (출하 / 폐사)")
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
                        
                        btn_s1, btn_s2 = st.columns(2)
                        with btn_s1:
                            submitted_status = st.form_submit_button("상태 변경", type="primary", width="stretch")
                            
                        if submitted_status:
                            target_id = cattle_opts[target_cattle_label]
                            wc = db_connect(DB_FILE)
                            wc.execute("UPDATE cattle SET status = ?, closure_date = ? WHERE cattle_id = ?", (new_status, closure_date.isoformat(), target_id))
                            wc.commit(); wc.close()
                            notify(f"개체 '{target_id}' → '{new_status}' 변경 완료", icon="✅")
                            st.rerun()
                    
                    st.markdown("---")
                    st.subheader("개체 위치(동/우방) 및 시험군 이동")
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
                        
                        btn_m1, btn_m2 = st.columns(2)
                        with btn_m1:
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
                st.dataframe(
                    df_disease, width="stretch", hide_index=True,
                    column_config={
                        "ID": st.column_config.NumberColumn(width="small"),
                        "발병일": st.column_config.DateColumn(format="YYYY-MM-DD"),
                        "완치일": st.column_config.DateColumn(format="YYYY-MM-DD"),
                        "수량(ml)": st.column_config.NumberColumn(format="localized", alignment="right"),
                    },
                )
    
    with sub_tab3:

        st.subheader("전체 개체 현황 (개체관리대장)")
        
        df_all_cattle = pd.read_sql("""
            SELECT c.cattle_id as 이표번호, c.kpn as KPN,
                   c.birth_date as 생년월일, c.admission_date as 입식일,
                   t.test_name as 시험군, c.status as 상태,
                   c.market_name as 우시장, c.building as 동, c.pen_number as 우방,
                   c.feed_type as 사료구분, c.roughage_grade as 조사료등급,
                   c.castration_date as 거세일,
                   c.calf_price as 구입금액, c.commission_fee as 수수료,
                   c.transport_fee as 운송료, c.initial_cost as 구입비용합계,
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
            '구입금액', '수수료', '운송료', '구입비용합계', '보험가입금액', '보험료', '종결일', '비고'
        ]
        df_all_cattle = df_all_cattle[cols_order]

        col_f1, col_f2, col_f3, col_f4 = st.columns([1, 1, 1, 1.5])
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
        )

        # 날짜는 달력 형식, 금액은 숫자형으로 두고 표시 형식만 column_config 로 지정한다.
        for col in ['생년월일', '입식일', '거세일', '종결일']:
            df_all_cattle[col] = pd.to_datetime(df_all_cattle[col], errors='coerce')
        money_cols = ['구입금액', '수수료', '운송료', '구입비용합계', '보험가입금액', '보험료']
        for col in money_cols:
            df_all_cattle[col] = pd.to_numeric(df_all_cattle[col], errors='coerce')
        df_all_cattle['우방'] = pd.to_numeric(df_all_cattle['우방'], errors='coerce')
        df_all_cattle['상태'] = df_all_cattle['상태'].map(lambda s: status_badge.get(s, s))

        df_all_cattle.insert(0, "선택", False)
        disabled_cols = [c for c in df_all_cattle.columns if c != "선택"]
        date_cfg = {c: st.column_config.DateColumn(c, format="YYYY-MM-DD") for c in ['생년월일', '입식일', '거세일', '종결일']}
        money_cfg = {
            c: st.column_config.NumberColumn(f"{c} (원)", format="localized", alignment="right")
            for c in money_cols
        }
        edited_all_cattle_df = st.data_editor(
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
                "우방": st.column_config.NumberColumn("우방", format="%d", alignment="center"),
                "동": st.column_config.TextColumn("동", alignment="center"),
                **date_cfg,
                **money_cfg,
            },
        )

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

with tab1:
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("사육 개체 요약")
        st.caption("&nbsp;", unsafe_allow_html=True)  # 오른쪽 캡션과 높이를 맞추기 위한 빈 여백
        df_cattle = pd.read_sql("""
            SELECT c.cattle_id as 개체번호, t.test_name as 시험군, c.status as 상태, c.initial_cost as 초기원가
            FROM cattle c
            JOIN testgroup_master t ON c.test_group_code = t.test_group_code
            ORDER BY c.status, c.cattle_id
        """, conn)
        df_cattle['상태'] = df_cattle['상태'].map({'사육': '🟢 사육', '출하': '🔵 출하', '폐사': '🔴 폐사'}).fillna(df_cattle['상태'])
        st.dataframe(
            df_cattle, width="stretch", hide_index=True,
            column_config={
                "상태": st.column_config.TextColumn(width="small"),
                "초기원가": st.column_config.NumberColumn("초기원가 (원)", format="localized", alignment="right"),
            },
        )
    with col_b:
        st.subheader("품목 및 재고 상태")
        st.caption("매입 시마다 이동평균단가가 자동으로 갱신됩니다.")
        df_item = pd.read_sql("SELECT item_code as 품목코드, item_name as 품목명, category as 분류, unit as 단위, current_stock as 현재재고, moving_avg_price as 이동평균단가 FROM item_master", conn)
        st.dataframe(
            df_item, width="stretch", hide_index=True,
            column_config={
                "현재재고": st.column_config.NumberColumn(format="localized", alignment="right"),
                "이동평균단가": st.column_config.NumberColumn("이동평균단가 (원)", format="localized", alignment="right"),
            },
        )

with tab0:
    col_left, col_right = st.columns(2)
    
    with col_left:
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
            new_item_unit = st.selectbox("단위", ["kg", "ml", "개"])
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

        st.markdown("---")
        st.markdown("##### 등록된 품목 목록 (체크박스로 삭제 가능)")
        df_items_all = pd.read_sql("SELECT item_code as 품목코드, item_name as 품목명, category as 분류, unit as 단위, current_stock as 현재재고, moving_avg_price as 이동평균단가 FROM item_master", conn)
        df_items_all.insert(0, "삭제", False)

        edited_item_df = st.data_editor(
            df_items_all,
            width="stretch",
            hide_index=True,
            disabled=["품목코드", "현재재고", "이동평균단가"],
            column_config={
                "삭제": st.column_config.CheckboxColumn("삭제", width="small"),
                "단위": st.column_config.SelectboxColumn("단위", options=["kg", "ml", "개"]),
                "현재재고": st.column_config.NumberColumn(format="localized", alignment="right"),
                "이동평균단가": st.column_config.NumberColumn("이동평균단가 (원)", format="localized", alignment="right"),
            },
            num_rows="dynamic",
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
                write_conn = db_connect(DB_FILE)
                current_codes = edited_item_df[~edited_item_df["삭제"]]['품목코드'].dropna().tolist()

                for _, row in edited_item_df.iterrows():
                    if not row.get("삭제", False) and pd.notna(row['품목코드']):
                        write_conn.execute("UPDATE item_master SET item_name=?, category=?, unit=? WHERE item_code=?", (row['품목명'], row['분류'], row.get('단위'), row['품목코드']))
            
                original_codes = df_items_all['품목코드'].dropna().tolist()
                missing_codes = set(original_codes) - set(current_codes)
                for code in missing_codes:
                    try:
                        write_conn.execute("DELETE FROM item_master WHERE item_code=?", (code,))
                    except sqlite3.IntegrityError:
                        st.error(f"'{code}' 품목은 매입 등 사용 내역이 있어 삭제할 수 없습니다.")
            
                write_conn.commit()
                write_conn.close()
                notify("품목 내역이 업데이트 되었습니다.", icon="✅")
                st.rerun()
    
    with col_right:
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
            st.info("먼저 좌측에서 품목을 등록해 주세요.")
        else:
            item_units = {r['item_code']: (r['unit'] if pd.notna(r['unit']) else "") for _, r in items_df.iterrows()}

            # 거래명세서 한 장에 여러 품목이 함께 들어오므로, 매입일자는 한 번만 고르고
            # 표에 줄마다 품목을 골라 수량·금액을 넣어 한꺼번에 등록한다 (시험군별 사용량 등록과 같은 방식).
            purchase_date = st.date_input("매입일자", key="purchase_date_input")

            # 선택지 이름에 남은 수량·단위를 붙여, 고른 뒤에도 칸에서 재고를 바로 볼 수 있게 한다.
            # 주의: 줄 추가가 되는 st.data_editor 는 넘기는 표(data)나 선택지(column_config)가 바뀌면
            # 새 표로 보고 입력한 내용을 지운다. 왼쪽에서 품목을 새로 등록하면 선택지가 바뀌므로,
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
            purchase_edited = st.data_editor(
                st.session_state["_purchase_entry_base"],
                width="stretch",
                hide_index=True,
                num_rows="dynamic",
                key=f"purchase_entry_editor_{purchase_ver}",
                column_config={
                    "품목": st.column_config.SelectboxColumn("품목 (남은 수량)", options=purchase_options, width="large"),
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
                money = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right")
                st.dataframe(
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

            # ----- 품목 현황 -----
            # 고른 매입일자가 속한 달의 매입 품목은 바로 보이게 두고,
            # 월말정산 후 남은 품목과 전체 매입 품목은 접어 두었다가 펼쳐서 본다.
            num = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right")
            purchase_month = purchase_date.strftime("%Y-%m")
            month_df = pd.read_sql("""
                SELECT i.item_name AS 품목명, i.category AS 분류, COUNT(*) AS 매입건수,
                       SUM(p.quantity) AS 매입수량, MAX(i.unit) AS 단위, SUM(p.total_amount) AS 매입금액
                FROM purchase p
                JOIN item_master i ON p.item_code = i.item_code
                WHERE substr(p.purchase_date, 1, 7) = ?
                GROUP BY p.item_code, i.item_name, i.category
                ORDER BY i.category, i.item_name
            """, conn, params=(purchase_month,))
            st.markdown(f"##### 📅 {purchase_month} 매입 품목")
            if month_df.empty:
                st.caption(f"{purchase_month}에 매입한 품목이 없습니다. (위 매입일자를 바꾸면 그 달의 매입 품목을 보여 줍니다)")
            else:
                month_df["평균단가"] = (month_df["매입금액"] / month_df["매입수량"].where(month_df["매입수량"] != 0)).round(0)
                st.dataframe(
                    month_df[["품목명", "분류", "매입건수", "매입수량", "단위", "평균단가", "매입금액"]],
                    width="stretch", hide_index=True,
                    column_config={
                        "매입건수": st.column_config.NumberColumn(width="small"),
                        "매입수량": num("매입수량"), "평균단가": num("평균단가 (원)"), "매입금액": num("매입금액 (원)"),
                    },
                )
                show_table_total(len(month_df), "매입금액", float(month_df["매입금액"].sum()))

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
            last_settled = conn.execute("SELECT MAX(settlement_month) FROM cattle_cost_log").fetchone()[0]

            remain_df = stock_df[stock_df["남은수량"] > 0.005].copy()
            settled_note = f"최근 정산 {last_settled}" if last_settled else "정산 내역 없음"
            with st.expander(f"📦 월말정산 후 남은 품목 ({len(remain_df)}개 · {settled_note})"):
                st.caption("남은 수량 = 매입 누계 − 월말 비용 등록의 시험군별 사용량 누계")
                if remain_df.empty:
                    st.caption("남은 품목이 없습니다.")
                else:
                    remain_df["평균단가"] = (remain_df["남은금액"] / remain_df["남은수량"]).round(0)
                    st.dataframe(
                        remain_df[["품목명", "분류", "남은수량", "단위", "평균단가", "남은금액"]],
                        width="stretch", hide_index=True,
                        column_config={"남은수량": num("남은 수량"), "평균단가": num("평균단가 (원)"), "남은금액": num("남은 금액 (원)")},
                    )
                    show_table_total(len(remain_df), "남은 금액", float(remain_df["남은금액"].sum()))

            all_df = stock_df[stock_df["누적매입수량"] > 0]
            with st.expander(f"🗂️ 전체 매입 품목 ({len(all_df)}개)"):
                if all_df.empty:
                    st.caption("매입한 품목이 없습니다.")
                else:
                    st.dataframe(
                        all_df[["품목명", "분류", "단위", "누적매입수량", "누적매입금액", "누적사용량", "남은수량"]],
                        width="stretch", hide_index=True,
                        column_config={
                            "누적매입수량": num("누적 매입수량"), "누적매입금액": num("누적 매입금액 (원)"),
                            "누적사용량": num("누적 사용량"), "남은수량": num("남은 수량"),
                        },
                    )
                    show_table_total(len(all_df), "누적 매입금액", float(all_df["누적매입금액"].sum()))

        st.markdown("---")
        st.markdown("##### 매입 내역 (체크박스로 삭제 가능)")
        df_purchase = pd.read_sql("""
            SELECT p.purchase_id as 매입ID, p.purchase_date as 매입일자, 
                   p.item_code as 품목코드, i.item_name as 품목명,
                   p.quantity as 수량, p.unit as 단위, p.total_amount as 총금액,
                   ROUND(p.total_amount / NULLIF(p.quantity, 0), 0) as 단가
            FROM purchase p
            JOIN item_master i ON p.item_code = i.item_code
            ORDER BY p.purchase_date DESC
        """, conn)
        df_purchase.insert(0, "삭제", False)
        
        if "purchase_editor" in st.session_state:
            edits = st.session_state["purchase_editor"].get("edited_rows", {})
            for row_idx, changes in edits.items():
                row_idx = int(row_idx)
                if row_idx < len(df_purchase):
                    new_amount = changes.get("총금액", df_purchase.at[row_idx, "총금액"])
                    new_qty = changes.get("수량", df_purchase.at[row_idx, "수량"])
                    if pd.notna(new_qty) and float(new_qty) != 0:
                        df_purchase.at[row_idx, "단가"] = round(float(new_amount) / float(new_qty))
            
            added = st.session_state["purchase_editor"].get("added_rows", [])
            for row in added:
                amt = row.get("총금액", 0)
                qty = row.get("수량", 0)
                if qty and float(qty) != 0:
                    row["단가"] = round(float(amt) / float(qty))
        
        edited_purchase_df = st.data_editor(
            df_purchase,
            width="stretch",
            hide_index=True,
            disabled=["매입ID", "품목명", "단가"],
            num_rows="dynamic",
            key="purchase_editor",
            column_config={
                "삭제": st.column_config.CheckboxColumn("삭제", width="small"),
                "매입ID": st.column_config.NumberColumn(width="small"),
                "수량": st.column_config.NumberColumn(format="localized", alignment="right"),
                "총금액": st.column_config.NumberColumn("총금액 (원)", format="localized", alignment="right"),
                "단가": st.column_config.NumberColumn("단가 (원)", format="localized", alignment="right"),
            },
        )
        
        if st.button("매입 수정 사항 저장", type="primary", width="stretch"):
            write_conn = db_connect(DB_FILE)
            current_ids = edited_purchase_df[~edited_purchase_df["삭제"]]['매입ID'].dropna().tolist()
            
            for _, row in edited_purchase_df.iterrows():
                if not row.get("삭제", False):
                    pid = row['매입ID']
                    if pd.notna(pid):
                        write_conn.execute("UPDATE purchase SET purchase_date=?, item_code=?, quantity=?, unit=?, total_amount=? WHERE purchase_id=?", 
                                         (row['매입일자'], row['품목코드'], row['수량'], row.get('단위', ''), row['총금액'], pid))
                    else:
                        if pd.notna(row['매입일자']) and pd.notna(row['품목코드']):
                            write_conn.execute("INSERT INTO purchase (purchase_date, item_code, quantity, unit, total_amount) VALUES (?, ?, ?, ?, ?)",
                                             (row['매입일자'], row['품목코드'], row['수량'], row.get('단위', ''), row['총금액']))
            
            original_ids = df_purchase['매입ID'].dropna().tolist()
            missing_ids = set(original_ids) - set(current_ids)
            for mid in missing_ids:
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
            notify("매입 내역이 업데이트 및 재고가 재계산 되었습니다.", icon="✅")
            st.rerun()

with tab2:
    st.subheader("월말 비용 등록 및 조회")
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

            usage_group_label = st.selectbox("시험군", list(group_options.keys()), key="usage_group_sel")

            # 한 시험군의 여러 품목을 표에 줄줄이 입력해 한 번에 등록한다.
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
            usage_edited = st.data_editor(
                pd.DataFrame({
                    "품목": pd.Series([None] * 8, dtype="object"),
                    "사용량": pd.Series([None] * 8, dtype="float"),
                }),
                width="stretch",
                hide_index=True,
                num_rows="dynamic",
                key=f"usage_entry_editor_{usage_ver}",
                column_config={
                    "품목": st.column_config.SelectboxColumn("품목 (남은 수량)", options=list(usage_name_to_code.keys()), width="large"),
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
                money = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right")
                st.dataframe(
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

            if st.button("사용량 일괄 등록", type="primary", width="stretch", key="submit_usage_btn"):
                if not re.fullmatch(r"\d{4}-\d{2}", usage_month.strip()):
                    st.warning("정산연월(YYYY-MM, 예: 2026-04)을 올바르게 입력하세요.")
                elif usage_problems:
                    st.warning("등록하지 않았습니다. 아래 품목을 확인하세요.\n\n" + "\n".join(f"- {p_}" for p_ in usage_problems))
                elif not usage_rows:
                    st.warning("사용한 품목 옆에 사용량을 입력하세요.")
                else:
                    usage_month = usage_month.strip()
                    sel_group = group_options[usage_group_label]
                    write_conn = db_connect(DB_FILE)
                    try:
                        # 한 트랜잭션으로 넣어 일부만 들어가는 일이 없게 한다.
                        write_conn.executemany(
                            "INSERT INTO monthly_usage (settlement_month, test_group_code, item_code, total_usage, applied_price, calculated_amount) VALUES (?, ?, ?, ?, ?, ?)",
                            [(usage_month, sel_group, code, qty, price, amount) for code, qty, price, amount in usage_rows],
                        )
                        write_conn.commit()
                    finally:
                        write_conn.close()
                    st.session_state["usage_view_month"] = usage_month
                    st.session_state["_usage_entry_ver"] = usage_ver + 1
                    notify(f"사용량 {len(usage_rows)}건 등록 완료! [{usage_month}] {usage_group_label} | "
                           f"산출액 합계 {sum(r[3] for r in usage_rows):,.0f}원", icon="✅")
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

        edited_usage_df = st.data_editor(
            df_usage,
            width="stretch",
            hide_index=True,
            disabled=["ID", "시험군", "품목명", "산출총액"],
            num_rows="fixed",  # 행 추가는 위 등록 폼으로만, 삭제는 '삭제' 체크박스로
            key=usage_editor_key,
            column_config={
                "삭제": st.column_config.CheckboxColumn("삭제", width="small"),
                "ID": st.column_config.NumberColumn(width="small"),
                "사용량": st.column_config.NumberColumn(format="localized", alignment="right"),
                "적용단가": st.column_config.NumberColumn("적용단가 (원)", format="localized", alignment="right"),
                "산출총액": st.column_config.NumberColumn("산출총액 (원)", format="localized", alignment="right"),
            },
        )
        # 삭제 체크한 행은 빼고, 표에서 고친 값(새로 추가한 행 포함)을 그대로 반영해 합산
        usage_kept = edited_usage_df[~edited_usage_df["삭제"].fillna(False).astype(bool)]
        show_table_total(
            len(usage_kept), "산출총액",
            pd.to_numeric(usage_kept["산출총액"], errors="coerce").fillna(0).sum(),
        )

        if st.button("사용 내역 수정 사항 저장", type="primary", width="stretch"):
            write_conn = db_connect(DB_FILE)
            current_usage_ids = edited_usage_df[~edited_usage_df["삭제"]]['ID'].dropna().tolist()

            for _, row in edited_usage_df.iterrows():
                uid = row['ID']
                if not row.get("삭제", False) and pd.notna(uid):
                    qty = float(row['사용량'])
                    price = float(row['적용단가'])
                    calc_amount = round(qty * price, 2)
                    write_conn.execute(
                        "UPDATE monthly_usage SET settlement_month=?, total_usage=?, applied_price=?, calculated_amount=? WHERE usage_id=?",
                        (row['정산연월'], qty, price, calc_amount, uid)
                    )

            original_usage_ids = df_usage['ID'].dropna().tolist()
            missing_usage_ids = set(original_usage_ids) - set(current_usage_ids)
            for uid in missing_usage_ids:
                write_conn.execute("DELETE FROM monthly_usage WHERE usage_id=?", (uid,))

            write_conn.commit()
            write_conn.close()
            notify("사용 내역이 업데이트 되었습니다.", icon="✅")
            st.rerun()
    
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
        
        edited_fc_df = st.data_editor(
            df_fixed,
            width="stretch",
            hide_index=True,
            disabled=["ID"],
            num_rows="fixed",  # 행 추가는 위 등록 폼으로만, 삭제는 '삭제' 체크박스로
            key=f"fixedcost_editor_{fc_view_month}",
            column_config={
                "삭제": st.column_config.CheckboxColumn("삭제", width="small"),
                "ID": st.column_config.NumberColumn(width="small"),
                "총청구금액": st.column_config.NumberColumn("총청구금액 (원)", format="localized", alignment="right"),
            },
        )
        fc_kept = edited_fc_df[~edited_fc_df["삭제"].fillna(False).astype(bool)]
        show_table_total(
            len(fc_kept), "총청구금액",
            pd.to_numeric(fc_kept["총청구금액"], errors="coerce").fillna(0).sum(),
        )

        if st.button("고정비 수정 사항 저장", type="primary", width="stretch"):
            write_conn = db_connect(DB_FILE)
            current_ids = edited_fc_df[~edited_fc_df["삭제"]]['ID'].dropna().tolist()
            
            original_ids = df_fixed['ID'].dropna().tolist()
            missing_ids = set(original_ids) - set(current_ids)
            for mid in missing_ids:
                write_conn.execute("DELETE FROM monthly_fixedcost WHERE fixed_cost_id=?", (int(mid),))
                
            for _, row in edited_fc_df.iterrows():
                if not row.get("삭제", False) and pd.notna(row['정산연월']) and str(row['정산연월']).strip() != "":
                    fid = row['ID']
                    if pd.isna(fid):
                        write_conn.execute(
                            "INSERT INTO monthly_fixedcost (settlement_month, expense_item, total_billed_amount) VALUES (?, ?, ?)",
                            (row['정산연월'], row['지출항목'], row['총청구금액'])
                        )
                    else:
                        write_conn.execute(
                            "UPDATE monthly_fixedcost SET settlement_month=?, expense_item=?, total_billed_amount=? WHERE fixed_cost_id=?",
                            (row['정산연월'], row['지출항목'], row['총청구금액'], int(fid))
                        )
            write_conn.commit()
            write_conn.close()
            notify("고정비 내역이 업데이트 되었습니다.", icon="✅")
            st.rerun()

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("### 🚀 월말 정산(일할계산) 실행")
    st.markdown("아래 버튼을 누르면 위에서 등록한 변동비·고정비를 분석하여, 이번 달 사육 이력이 있는 각 개체에 **실제 사육일수에 비례해(일할계산)** 변동비와 고정비를 배분합니다.")
    
    # 정산 대상 연월은 비용(사용량·고정비)이 등록됐거나 이미 정산된 연월 중에서 고른다.
    # 직접 입력하게 두면 비용이 없는 달(예: 이번 달)이 그대로 정산되는 실수가 생긴다.
    # 위 공통 정산연월을 바꾸면 여기 값도 따라 바뀐다(settlement_month_input). 처음에는 그 정산연월로 시작.
    calc_months = sorted({r[0] for r in conn.execute("""
        SELECT settlement_month FROM monthly_usage
        UNION SELECT settlement_month FROM monthly_fixedcost
        UNION SELECT settlement_month FROM cattle_cost_log
    """).fetchall() if r[0] and re.fullmatch(r"\d{4}-\d{2}", str(r[0]))}, reverse=True)
    if not calc_months:
        # 아래 탭들이 계속 그려져야 하므로 st.stop() 대신 이번 달 하나만 선택지로 둔다.
        st.info("비용(사용량·고정비)이 등록된 연월이 없습니다. 위에서 비용을 먼저 등록하세요.")
        calc_months = [datetime.now().strftime('%Y-%m')]
    settled_months_set = {r[0] for r in conn.execute("SELECT DISTINCT settlement_month FROM cattle_cost_log").fetchall()}
    if "calc_target_month" not in st.session_state:
        start_month = st.session_state.get("cost_month_input")
        st.session_state["calc_target_month"] = start_month if start_month in calc_months else calc_months[0]
    elif st.session_state["calc_target_month"] not in calc_months:
        # 위에서 비용이 없는 달로 바꾼 경우: 선택지에 없으니 가장 최근 등록 연월로 둔다.
        st.session_state["calc_target_month"] = calc_months[0]
    target_month = st.selectbox("정산 대상 연월", calc_months, key="calc_target_month",
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
        money = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right")
        st.dataframe(
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
        st.info(f"[{target_month}] 에 등록된 사용량·고정비가 없습니다. 정산 전에 위에서 비용을 먼저 등록하세요.")
    
    st.caption("※ 입식 당일은 절식하므로 배분에서 제외하고, 입식 다음 날부터 사육일수로 계산합니다.")
    settled_count = conn.execute(
        "SELECT COUNT(*) FROM cattle_cost_log WHERE settlement_month = ?", (target_month,)
    ).fetchone()[0]

    if settled_count == 0:
        if st.button("🚀 정산 실행(일할계산) 및 누적원가 반영", type="primary"):
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
        if st.button("🔄 기존 정산 지우고 다시 정산", type="primary"):
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
            if st.button("정산 취소", disabled=not confirm_cancel, key=f"cancel_settle_btn_{target_month}"):
                backup_db(DB_FILE, f"before-cancel-settle-{target_month}")
                write_conn = db_connect(DB_FILE)
                write_conn.execute("DELETE FROM cattle_item_usage_log WHERE settlement_month = ?", (target_month,))
                write_conn.execute("DELETE FROM cattle_cost_log WHERE settlement_month = ?", (target_month,))
                write_conn.commit()
                write_conn.close()
                notify(f"[{target_month}] 정산을 취소했습니다. (배분 내역 {settled_count}건 삭제)", icon="🗑️")
                st.rerun()

    st.markdown("---")
    st.subheader(f"[{target_month}] 개체별 원가 적재 결과 (Cattle_Cost_Log)")
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
            money = lambda label: st.column_config.NumberColumn(label, format="localized", alignment="right")
            st.dataframe(
                df_log, width="stretch", hide_index=True,
                column_config={
                    "cattle_id": st.column_config.TextColumn("이표번호"),
                    "settlement_month": st.column_config.TextColumn("정산연월"),
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
    except:
        st.info("아직 정산된 내역이 없습니다.")

with tab_report:
    st.subheader("🧾 결산 리포트 생성")
    st.markdown("정산이 완료된 연월을 골라 인쇄·저장 가능한 결산서(HTML)를 만듭니다. 브라우저에서 열어 **PDF 인쇄 및 저장** 버튼으로 PDF로도 저장할 수 있습니다.")

    settled_months = list_settled_months(DB_FILE)
    if not settled_months:
        st.info("아직 정산된 연월이 없습니다. 먼저 '🚀 월말 정산(일할계산) 실행' 탭에서 정산을 실행하세요.")
    else:
        report_month = st.selectbox("리포트 연월", settled_months, key="report_month_select")
        if st.button("📄 결산 리포트 생성", type="primary", key="gen_report_btn"):
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

with tab_slaughter:
    st.subheader("🥩 도축 성적 관리")
    st.info("이 탭은 추후 출하(도축)된 개체들의 도축 성적(등급, 도체중, 등심단면적, 근내지방도 등)을 기록하고 확인하기 위한 메뉴입니다.\n\n현재 준비중입니다.")

conn.close()
