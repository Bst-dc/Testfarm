import streamlit as st
import sqlite3
import pandas as pd
import os
import io
import glob
import shutil
import tempfile
import html
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
DB_DIR = os.environ.get("ERP_DB_DIR") or os.path.join(os.path.expanduser("~"), "시험농장DB")
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

st.markdown(
    """
    <style>
    /* 웹 폰트 적용 (Pretendard) */
    @import url("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/static/pretendard.css");
    html, body, [class*="st-"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, 'Helvetica Neue', 'Segoe UI', 'Apple SD Gothic Neo', 'Noto Sans KR', 'Malgun Gothic', sans-serif !important;
    }
    /* 위 폰트 강제 적용이 Streamlit 내장 아이콘(사이드바 접기 화살표, expander 화살표 등)의
       전용 아이콘 폰트까지 덮어써서 "keyboard_double_arrow_left" 같은 글자가 그대로 보이는
       문제를 막기 위해 아이콘 요소는 원래 아이콘 폰트로 되돌린다. */
    [data-testid="stIconMaterial"] {
        font-family: "Material Symbols Rounded" !important;
    }

    /* 전체 배경 - 세이지그린 → 더스티블루 대각선 그라데이션 (한우 스마트 컨설팅 스타일) */
    .stApp {
        background:
            radial-gradient(1200px 700px at 0% 0%, rgba(176, 205, 183, 0.9) 0%, rgba(176, 205, 183, 0) 60%),
            radial-gradient(1000px 800px at 100% 100%, rgba(171, 196, 216, 0.9) 0%, rgba(171, 196, 216, 0) 60%),
            linear-gradient(135deg, #C6D8C3 0%, #C0D2DF 100%);
        background-attachment: fixed;
        color: #2B2B28;
    }

    /* 사이드바 배경 - 같은 계열의 반투명 그라데이션 */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, rgba(166, 198, 174, 0.85) 0%, rgba(160, 186, 210, 0.85) 100%) !important;
        border-right: 1px solid rgba(255, 255, 255, 0.4);
    }
    section[data-testid="stSidebar"] * {
        color: #2B2B28;
    }

    /* 메트릭 카드 (핵심 지표) 스타일링 - 민트 카드 + 웜톤 그림자 */
    div[data-testid="metric-container"], div[data-testid="stMetric"] {
        background-color: #E6EFE3;
        border: 1px solid #CFDCCB;
        padding: 20px 24px;
        border-radius: 14px;
        box-shadow: 0 4px 6px -1px rgba(60, 55, 40, 0.05), 0 2px 4px -1px rgba(60, 55, 40, 0.03);
        border-left: 6px solid #5E7F66; /* 세이지그린 포인트 */
        transition: transform 0.2s ease;
    }
    div[data-testid="metric-container"]:hover, div[data-testid="stMetric"]:hover {
        transform: translateY(-2px);
    }
    div[data-testid="metric-container"] > div, div[data-testid="stMetric"] > div {
        color: #2B2B28; /* 제목 색상 */
    }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"],
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        font-weight: 800;
        font-size: 2.2rem;
        color: #1F2A22;
    }

    /* 탭 메뉴(헤더)를 최신 웹앱 버튼형(Pill) 스타일로 변경
       (예전 BaseWeb 마크업 기준 [data-baseweb="tab"] 셀렉터는 현재 Streamlit 버전에서
        div[data-testid="stTab"] 로 바뀌어 더 이상 매치되지 않았음 — 실제 DOM 기준으로 수정) */
    div[data-testid="stTabs"] div[role="tablist"] {
        gap: 12px;
        border-bottom: none;
        padding-bottom: 10px;
    }
    div[data-testid="stTab"] {
        background-color: rgba(255, 255, 255, 0.6) !important;
        border: 1px solid rgba(255, 255, 255, 0.8) !important;
        border-radius: 30px !important;
        padding: 12px 24px !important;
        box-shadow: 0 2px 4px rgba(60, 55, 40, 0.04);
        transition: all 0.2s ease;
    }
    div[data-testid="stTab"]:hover {
        background-color: rgba(255, 255, 255, 0.85) !important;
        border-color: rgba(255, 255, 255, 1) !important;
    }
    div[data-testid="stTab"][aria-selected="true"] {
        background: linear-gradient(135deg, #5E7F66 0%, #4A6F8A 100%) !important;
        border-color: transparent !important;
        box-shadow: 0 4px 10px rgba(74, 111, 138, 0.3);
    }
    div[data-testid="stTab"][aria-selected="true"] * {
        color: #FFFFFF !important;
    }
    div[data-testid="stTab"] [data-testid="stMarkdownContainer"] p {
        font-size: 20px !important;
        font-weight: 700 !important;
        margin: 0;
    }
    /* Streamlit 기본 탭 활성화 시 나타나는 하단 붉은 선(Indicator) 제거
       (실제 DOM에서는 stTab 안의 .react-aria-SelectionIndicator 요소였음 —
        stTabIndicator/tab-highlight 셀렉터는 존재하지 않는 이름이라 안 지워지고 있었음) */
    div[data-testid="stTab"] .react-aria-SelectionIndicator {
        display: none !important;
    }

    /* 사이드바 라디오 버튼(농장 선택)을 크고 예쁜 글래스 카드로 변경
       (Streamlit이 라디오를 React Aria 기반으로 바꾸면서 label이 radiogroup의 직계 자식이
        아니게 되었고, 선택 상태도 data-checked/aria-checked가 아니라 label 자체의
        data-selected="true" 로 표시됨 — 실제 DOM 기준으로 셀렉터를 다시 작성) */
    section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"] {
        padding: 16px 20px !important;
        background-color: rgba(255, 255, 255, 0.55) !important;
        border: 2px solid rgba(255, 255, 255, 0.7) !important;
        border-radius: 12px !important;
        margin-bottom: 12px !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 1px 3px rgba(60, 55, 40, 0.05) !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"]:hover {
        border-color: rgba(255, 255, 255, 1) !important;
        background-color: rgba(255, 255, 255, 0.75) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 4px 6px rgba(60, 55, 40, 0.1) !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"][data-selected="true"] {
        background: linear-gradient(135deg, #5E7F66 0%, #4A6F8A 100%) !important;
        border-color: transparent !important;
        box-shadow: 0 4px 10px rgba(74, 111, 138, 0.3) !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"][data-selected="true"] p {
        color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"] p {
        font-size: 22px !important;
        font-weight: 800 !important;
        color: #2B2B28 !important;
        margin: 0 !important;
    }

    /* 라디오 버튼의 동그라미 숨기기 (텍스트만 돋보이게) */
    section[data-testid="stSidebar"] div[role="radiogroup"] label[data-testid="stRadioOption"] > div > div:first-child {
        display: none !important;
    }

    /* 사이드바 라디오 버튼 제목(위젯 라벨) 크기 키우기 */
    section[data-testid="stSidebar"] div[data-testid="stWidgetLabel"] p {
        font-size: 26px !important;
        font-weight: 900 !important;
        color: #1F2A22 !important;
        padding-bottom: 10px !important;
    }

    /* 표(Dataframe) 디자인 깔끔하게 */
    div[data-testid="stDataFrame"] {
        border: 1px solid #CFDCCB;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 2px 4px rgba(60, 55, 40, 0.03);
    }

    /* 버튼 - 세이지그린 → 스틸블루 그라데이션
       (앱의 등록/수정/삭제 버튼은 대부분 st.form_submit_button 이라 stFormSubmitButton
        래퍼를 쓰는데, 기존 셀렉터에 빠져 있어서 대부분의 버튼이 기본 빨간색으로 남아 있었음) */
    .stButton > button, .stDownloadButton > button,
    div[data-testid="stFormSubmitButton"] > button {
        border-radius: 10px !important;
        border: none !important;
        background: linear-gradient(135deg, #5E7F66 0%, #4A6F8A 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        box-shadow: 0 2px 6px rgba(74, 111, 138, 0.25);
        transition: transform 0.15s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover,
    div[data-testid="stFormSubmitButton"] > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 10px rgba(74, 111, 138, 0.3);
    }

    /* 입력창/선택창 - 반투명 글래스 톤
       (셀렉트박스/멀티셀렉트도 BaseWeb에서 React Aria로 바뀌어 data-baseweb="select"가
        더 이상 없음 — 실제 래퍼는 stSelectbox/stMultiSelect 안의 role="group" 요소) */
    div[data-testid="stSelectbox"] div[role="group"],
    div[data-testid="stMultiSelect"] div[role="group"],
    .stTextInput input, .stNumberInput input, .stDateInput input {
        background-color: rgba(255, 255, 255, 0.75) !important;
        border-color: #CFDCCB !important;
        border-radius: 10px !important;
    }

    /* expander도 카드 톤으로 통일 */
    div[data-testid="stExpander"] {
        background-color: rgba(255, 255, 255, 0.55);
        border: 1px solid rgba(255, 255, 255, 0.7);
        border-radius: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True
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

SQLITE_DDL = """
CREATE TABLE IF NOT EXISTS testgroup_master (
    test_group_code TEXT PRIMARY KEY,
    test_name TEXT NOT NULL,
    start_date DATE NOT NULL,
    end_date DATE,
    location_mapping TEXT
);

CREATE TABLE IF NOT EXISTS cattle (
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
);

CREATE TABLE IF NOT EXISTS disease_record (
    record_id INTEGER PRIMARY KEY AUTOINCREMENT,
    cattle_id TEXT REFERENCES cattle(cattle_id),
    onset_date DATE NOT NULL,
    symptom TEXT,
    medicine1 TEXT,
    dosage1 NUMERIC(10, 2),
    medicine2 TEXT,
    dosage2 NUMERIC(10, 2),
    medicine3 TEXT,
    dosage3 NUMERIC(10, 2),
    veterinarian TEXT,
    recovery_date DATE,
    prescription_no TEXT,
    memo TEXT
);

CREATE TABLE IF NOT EXISTS item_master (
    item_code TEXT PRIMARY KEY,
    item_name TEXT NOT NULL,
    category TEXT CHECK (category IN ('사료', '조사료', '약품', '기타저장품')) NOT NULL,
    unit TEXT,
    current_stock NUMERIC(10, 2) NOT NULL DEFAULT 0,
    moving_avg_price NUMERIC(12, 2) NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS purchase (
    purchase_id INTEGER PRIMARY KEY AUTOINCREMENT,
    purchase_date DATE NOT NULL,
    item_code TEXT REFERENCES item_master(item_code),
    quantity NUMERIC(10, 2) NOT NULL,
    unit TEXT,
    total_amount NUMERIC(12, 2) NOT NULL
);

CREATE TABLE IF NOT EXISTS monthly_usage (
    usage_id INTEGER PRIMARY KEY AUTOINCREMENT,
    settlement_month TEXT NOT NULL,
    test_group_code TEXT REFERENCES testgroup_master(test_group_code),
    item_code TEXT REFERENCES item_master(item_code),
    total_usage NUMERIC(10, 2) NOT NULL,
    applied_price NUMERIC(12, 2) NOT NULL,
    calculated_amount NUMERIC(12, 2) NOT NULL
);

CREATE TABLE IF NOT EXISTS monthly_fixedcost (
    fixed_cost_id INTEGER PRIMARY KEY AUTOINCREMENT,
    settlement_month TEXT NOT NULL,
    expense_item TEXT NOT NULL,
    total_billed_amount NUMERIC(12, 2) NOT NULL
);

CREATE TABLE IF NOT EXISTS cattle_cost_log (
    log_id INTEGER PRIMARY KEY AUTOINCREMENT,
    cattle_id TEXT REFERENCES cattle(cattle_id),
    settlement_month TEXT NOT NULL,
    allocated_variable_cost NUMERIC(12, 2) NOT NULL DEFAULT 0,
    allocated_fixed_cost NUMERIC(12, 2) NOT NULL DEFAULT 0,
    UNIQUE (cattle_id, settlement_month)
);

CREATE TABLE IF NOT EXISTS cattle_item_usage_log (
    log_id INTEGER PRIMARY KEY AUTOINCREMENT,
    cattle_id TEXT REFERENCES cattle(cattle_id),
    settlement_month TEXT NOT NULL,
    item_code TEXT REFERENCES item_master(item_code),
    allocated_usage NUMERIC(10, 2) NOT NULL,
    allocated_amount NUMERIC(12, 2) NOT NULL,
    UNIQUE (cattle_id, settlement_month, item_code)
);

CREATE TRIGGER IF NOT EXISTS trg_after_insert_purchase
AFTER INSERT ON purchase
FOR EACH ROW
BEGIN
    UPDATE item_master
    SET current_stock = current_stock + NEW.quantity,
        moving_avg_price = CASE 
            WHEN current_stock + NEW.quantity > 0 
            THEN ROUND(((current_stock * moving_avg_price) + NEW.total_amount) / (current_stock + NEW.quantity), 2)
            ELSE 0 
        END
    WHERE item_code = NEW.item_code;
END;
"""

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
    conn.execute("PRAGMA busy_timeout = 30000")
    conn.execute("PRAGMA journal_mode = WAL")  # 읽는 중에도 쓰기가 막히지 않는다
    conn.execute("PRAGMA foreign_keys = ON")   # 스키마에 선언된 참조 무결성을 실제로 적용
    _conn_registry().append(conn)
    return conn


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
        _repair_dangling_cattle_refs(conn)
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


def _table_ddl(table):
    """SQLITE_DDL 에서 해당 표의 CREATE TABLE 문만 잘라 낸다."""
    marker = f"CREATE TABLE IF NOT EXISTS {table} ("
    start = SQLITE_DDL.find(marker)
    if start < 0:
        return None
    end = SQLITE_DDL.find(");", start)
    return SQLITE_DDL[start:end + 1]


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
    if not os.path.exists(db_file):
        return None
    name = os.path.splitext(os.path.basename(db_file))[0]
    out_dir = dest_dir or BACKUP_DIR
    os.makedirs(out_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  # 생성 날짜·시간을 파일명에 넣는다
    out = os.path.join(out_dir, "%s_%s_%s.db" % (name, stamp, reason))
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
        probe = sqlite3.connect(tmp_path)
        if probe.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            probe.close()
            return False, "파일이 손상되었습니다. 다른 백업 파일을 사용하세요."
        tables = {r[0] for r in probe.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        probe.close()
        missing = {"cattle", "testgroup_master", "item_master"} - tables
        if missing:
            return False, "이 시스템의 DB 파일이 아닙니다. (없는 표: %s)" % ", ".join(sorted(missing))

        if os.path.exists(db_file):
            backup_db(db_file, "before-restore")
        close_stale_connections()
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

def distribute_monthly_costs(db_file, settlement_month):
    import calendar
    from datetime import datetime
    conn = db_connect(db_file)
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
        conn.close()
        return False, "사육 중이거나 정산월에 포함되는 개체가 없습니다."

    year, month = map(int, settlement_month.split('-'))
    last_day = calendar.monthrange(year, month)[1]
    month_start = datetime(year, month, 1)
    month_end = datetime(year, month, last_day)

    def calc_days(row):
        adm = row['admission_date']
        adm_d = datetime.strptime(str(adm)[:10], '%Y-%m-%d') if pd.notna(adm) and str(adm).strip() else month_start
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
    cattle_df = cattle_df[cattle_df['rearing_days'] > 0]
    
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
    if not os.path.exists(db_file):
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


# ========== UI 메인 ==========

# 이전 실행(rerun)에서 닫히지 않은 연결부터 정리한다. -> "database is locked" 방지
close_stale_connections()

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

with st.sidebar.expander("➕ 새 농장 추가"):
    with st.form("add_farm_form", clear_on_submit=True):
        new_farm_name = st.text_input("새 농장 이름")
        new_farm_color = st.color_picker("테마 색상", "#3B82F6")
        new_farm_b = st.number_input("동 개수 (예: 6)", min_value=1, max_value=20, value=6)
        new_farm_p = st.number_input("동별 우방 개수 (예: 20)", min_value=1, max_value=100, value=20)
        if st.form_submit_button("추가"):
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
                    st.success(f"'{new_farm_name}' 농장이 추가되었습니다!")
                    st.rerun()


if selected_farm == "시험농장 전체 현황":
    st.markdown(
        f"""
        <div style="display:flex; align-items:center; gap:14px; margin-bottom:0.5rem;">
            <img src="{MEDAL_ICON_DATA_URI}" alt="심볼" style="height: 44px;">
            <h1 style="margin:0;">대구축협 시험농장 관리 시스템</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("데이터베이스 트리거에 의한 **단가 자동 갱신** 및 Pandas를 이용한 **월말 1/n 비용 분배**를 시각적으로 확인하는 대시보드입니다.")
    st.subheader("🌐 농장 통합 대시보드")
    st.caption("등록된 모든 관리 농장(선산, 고아 등)의 개체 현황을 통합하여 보여줍니다.")
    
    all_cattle_dfs = []
    
    for farm_nm, farm_cfg_info in FARM_CONFIG.items():
        f_db = farm_cfg_info["db_file"]
        if os.path.exists(f_db):
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
        
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("전체 누적 입식", f"{total_admission}두")
        m2.metric("현재 사육중", f"{current_breeding}두")
        m3.metric("누적 폐사", f"{dead_cattle}두")
        m4.metric("누적 출하", f"{shipped_cattle}두")
        m5.metric("총 구입비용", f"{total_initial_cost // 10000:,}만원")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.markdown("##### 🏢 농장별 요약 현황")
        farm_summary = df_all.groupby('농장명').agg(
            전체입식=('개체번호', 'count'),
            사육중=('상태', lambda x: (x == '사육').sum()),
            출하=('상태', lambda x: (x == '출하').sum()),
            폐사=('상태', lambda x: (x == '폐사').sum()),
            총구입비용_만원=('초기원가', lambda x: int(x.sum(skipna=True)) // 10000 if '초기원가' in df_all.columns else 0),
            평균구입금액_만원=('초기원가', lambda x: int(x.mean(skipna=True)) // 10000 if '초기원가' in df_all.columns and not x.isna().all() else 0)
        ).reset_index()
        
        farm_summary.rename(columns={
            '전체입식': '전체 입식 (두)', 
            '사육중': '현재 사육중 (두)', 
            '출하': '누적 출하 (두)', 
            '폐사': '누적 폐사 (두)', 
            '총구입비용_만원': '총 구입비용 (만원)',
            '평균구입금액_만원': '두당 평균구입금액 (만원)'
        }, inplace=True)
        
        for col in ['전체 입식 (두)', '현재 사육중 (두)', '누적 출하 (두)', '누적 폐사 (두)', '총 구입비용 (만원)', '두당 평균구입금액 (만원)']:
            farm_summary[col] = farm_summary[col].apply(lambda x: f"{int(x):,}")
            
        st.dataframe(
            farm_summary.style.set_properties(**{'text-align': 'center'})
            .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}]),
            width="stretch", hide_index=True
        )
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.markdown("##### 🏪 농장별 우시장 구입 현황")
        df_market = df_all[df_all['우시장'].notna() & (df_all['우시장'].astype(str).str.strip() != '')]
        if not df_market.empty:
            market_summary = df_market.groupby(['농장명', '우시장']).agg(
                구입마릿수=('개체번호', 'count'),
                총구입비용_만원=('초기원가', lambda x: int(x.sum(skipna=True)) // 10000 if '초기원가' in df_market.columns else 0),
                평균구입비용_만원=('초기원가', lambda x: int(x.mean(skipna=True)) // 10000 if '초기원가' in df_market.columns and not x.isna().all() else 0)
            ).reset_index()
            
            market_summary.rename(columns={
                '구입마릿수': '구입 마릿수 (두)',
                '총구입비용_만원': '총 구입비용 (만원)',
                '평균구입비용_만원': '두당 평균 (만원)'
            }, inplace=True)
            
            for col in ['구입 마릿수 (두)', '총 구입비용 (만원)', '두당 평균 (만원)']:
                market_summary[col] = market_summary[col].apply(lambda x: f"{int(x):,}")
                
            st.dataframe(
                market_summary.style.set_properties(**{'text-align': 'center'})
                .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}]),
                width="stretch", hide_index=True
            )
        else:
            st.info("등록된 우시장 구입 이력이 없습니다.")
            
        st.markdown("<br>", unsafe_allow_html=True)
                
        with st.expander("통합 데이터 상세 표"):
            df_all.insert(0, '순번', range(1, len(df_all) + 1))
            st.dataframe(
                df_all, 
                width="stretch", 
                hide_index=True,
                column_config={"순번": st.column_config.NumberColumn(width=60)}
            )
            
        st.markdown("---")
        st.subheader("📑 통합 보고서 생성")
        st.markdown("현재 전체 현황 대시보드의 요약 수치 및 농장별 데이터 표를 기반으로 인쇄 가능한 HTML 보고서를 생성합니다.")
        if st.button("📄 보고서 생성", type="primary"):
            if not df_market.empty:
                market_html_parts = []
                farms_in_market = [f for f in FARM_CONFIG.keys() if f in market_summary['농장명'].values]
                for i, farm_nm in enumerate(farms_in_market, start=1):
                    farm_market = market_summary[market_summary['농장명'] == farm_nm].drop(columns=['농장명'])
                    market_html_parts.append(f"<h3>{i}) {html.escape(farm_nm)}</h3>")
                    market_html_parts.append(farm_market.to_html(index=False, classes='table', justify='center'))
                market_html = "".join(market_html_parts)
            else:
                market_html = '<p>우시장 구입 이력이 없습니다.</p>'
            html_content = f"""
            <html>
            <head>
                <meta charset="utf-8">
                <title>대구축협 시험농장 현황 보고</title>
                <style>
                    body {{ font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; line-height: 1.6; padding: 20px; background: #f8fafc; }}
                    h1, h3 {{ color: #333; }}
                    h2 {{ color: #1e293b; border-left: 6px solid #4F46E5; padding-left: 12px; margin-top: 36px; }}
                    .summary-box {{ display: grid; grid-template-columns: repeat(5, 1fr); gap: 14px; margin-bottom: 24px; }}
                    .metric {{
                        background: #ffffff; border-radius: 12px; padding: 20px 10px;
                        text-align: center; box-shadow: 0 2px 8px rgba(30, 41, 59, 0.08);
                        border-top: 5px solid var(--accent, #4F46E5);
                    }}
                    .metric.total {{ --accent: #4F46E5; }}
                    .metric.breeding {{ --accent: #059669; }}
                    .metric.dead {{ --accent: #DC2626; }}
                    .metric.shipped {{ --accent: #2563EB; }}
                    .metric.cost {{ --accent: #D97706; }}
                    .metric .title {{ font-size: 13px; color: #64748b; font-weight: 600; letter-spacing: 0.03em; }}
                    .metric .value {{ font-size: 28px; font-weight: 800; color: var(--accent, #2c3e50); margin-top: 6px; }}
                    .table {{ width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 14px; text-align: center; background: #fff; }}
                    .table th, .table td {{ border: 1px solid #ddd; padding: 8px; }}
                    .table th {{ background-color: #2c3e50; color: white; text-align: center !important; }}
                    @media print {{
                        body {{ background: #fff; }}
                        .metric {{ border: 1px solid #ccc; box-shadow: none; break-inside: avoid; }}
                        .table th {{ color: black; }}
                    }}
                    @media (max-width: 700px) {{
                        .summary-box {{ grid-template-columns: repeat(2, 1fr); }}
                    }}
                </style>
            </head>
            <body>
                <div style="display: flex; align-items: center; justify-content: center; margin-bottom: 20px;">
                    <img src="{MEDAL_ICON_DATA_URI}" alt="심볼" style="height: 40px; margin-right: 15px;">
                    <h1 style="margin: 0;">대구축협 시험농장 현황 보고</h1>
                </div>
                <p><strong>기준일:</strong> {datetime.now().strftime('%Y년 %m월 %d일')}</p>
                
                <h2>1. 전체 요약 현황</h2>
                <div class="summary-box">
                    <div class="metric total"><div class="title">전체 누적 입식</div><div class="value">{total_admission:,}두</div></div>
                    <div class="metric breeding"><div class="title">현재 사육중</div><div class="value">{current_breeding:,}두</div></div>
                    <div class="metric dead"><div class="title">누적 폐사</div><div class="value">{dead_cattle:,}두</div></div>
                    <div class="metric shipped"><div class="title">누적 출하</div><div class="value">{shipped_cattle:,}두</div></div>
                    <div class="metric cost"><div class="title">총 구입비용</div><div class="value">{total_initial_cost // 10000:,}만원</div></div>
                </div>
                
                <h2>2. 농장별 요약 현황</h2>
                {farm_summary.to_html(index=False, classes='table', justify='center')}
                
                <h2>3. 농장별 우시장 구입 현황</h2>
                {market_html}
            </body>
            </html>
            """
            st.session_state["overall_report_html"] = html_content
            
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
            st.components.v1.html(overall_html, height=800, scrolling=True)

    else:
        st.info("데이터가 있는 농장이 없습니다.")
    st.stop()

farm_cfg = FARM_CONFIG[selected_farm]
DB_FILE = farm_cfg["db_file"]
farm_color = farm_cfg["color"]
farm_b_cnt = farm_cfg.get("buildings_count", 6)
farm_p_cnt = farm_cfg.get("pens_count", 20)


# DB 자동 생성
if not os.path.exists(DB_FILE):
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
            st.success(msg)
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
    if not os.path.exists(db_file):
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


def cattle_reset_preview(db_file):
    """개체 전체 삭제로 함께 지워질 데이터 건수 (cattle 및 cattle을 참조하는 표)."""
    tables = [
        ("개체", "cattle"),
        ("질병·처방 기록", "disease_record"),
        ("개체별 원가 내역", "cattle_cost_log"),
        ("개체별 품목 사용 내역", "cattle_item_usage_log"),
    ]
    rows = []
    if not os.path.exists(db_file):
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
def cattle_reset_dialog(farm_name, db_file):
    st.error(f"**{farm_name}**에 등록된 모든 개체(입식 내역)가 삭제됩니다. 이 작업은 되돌릴 수 없습니다.")
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
    if col_run.button("네, 삭제합니다", type="primary", width="stretch", key="cattle_reset_run"):
        backup_db(db_file, "before-cattle-reset")
        wc = db_connect(db_file)
        wc.execute("DELETE FROM cattle_item_usage_log")
        wc.execute("DELETE FROM cattle_cost_log")
        wc.execute("DELETE FROM disease_record")
        wc.execute("DELETE FROM cattle")
        wc.commit(); wc.close()
        st.session_state["cattle_reset_done"] = True
        st.rerun()


st.sidebar.markdown("---")
with st.sidebar.expander("🐄 등록된 개체 전체 삭제"):
    st.caption(f"**{selected_farm}**에 등록된 개체(입식 내역)만 삭제됩니다. 시험군·품목 등은 유지됩니다. 실행 직전 자동으로 백업본을 만듭니다.")
    if st.button("개체 전체 삭제 실행", width="stretch"):
        st.session_state["show_cattle_reset_dialog"] = True

if st.session_state.pop("show_cattle_reset_dialog", False):
    cattle_reset_dialog(selected_farm, DB_FILE)
if st.session_state.pop("cattle_reset_done", False):
    st.sidebar.success("등록된 개체가 모두 삭제되었습니다. (직전 상태는 백업에 보관)")

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
st.sidebar.caption("저장 위치: %s" % DB_DIR)

# 타이틀 (선택된 농장 표시)
st.markdown(
    f"""
    <div style="display:flex; align-items:center; gap:16px; margin-bottom:4px;">
        <div style="background:{farm_color}; color:white; padding:6px 18px; border-radius:8px; font-weight:800; font-size:0.95rem; letter-spacing:1px;">
            {selected_farm}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown(
    f"""
    <div style="display:flex; align-items:center; gap:14px; margin-bottom:0.5rem;">
        <img src="{MEDAL_ICON_DATA_URI}" alt="심볼" style="height: 44px;">
        <h1 style="margin:0;">대구축협 시험농장 관리 시스템</h1>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown("데이터베이스 트리거에 의한 **단가 자동 갱신** 및 Pandas를 이용한 **월말 1/n 비용 분배**를 시각적으로 확인하는 대시보드입니다.")

conn = db_connect(DB_FILE)

# KPI Metrics 표시
col_m1, col_m2, col_m3, col_m4 = st.columns(4)

total_admitted = pd.read_sql("SELECT COUNT(*) as cnt FROM cattle", conn).iloc[0]['cnt']
current_raising = pd.read_sql("SELECT COUNT(*) as cnt FROM cattle WHERE status='사육'", conn).iloc[0]['cnt']
dead_count = pd.read_sql("SELECT COUNT(*) as cnt FROM cattle WHERE status='폐사'", conn).iloc[0]['cnt']

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

col_m1.metric("입식 두수", f"{total_admitted} 마리")
col_m2.metric("사육 두수", f"{current_raising} 마리")
col_m3.metric("폐사 두수", f"{dead_count} 마리")
col_m4.metric("평균 개월령", avg_months_str)
st.markdown("---")

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
                            st.success(f"시험군 '{new_group_name}' 등록 완료")
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
                            st.success("시험군 수정 완료")
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
                            st.success(f"시험군 '{target_code}' 삭제 완료")
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
                                do_bulk_upload = st.button("개체 일괄등록", type="primary", use_container_width=True)

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
                                st.success(f"개체 '{new_cattle_id}' 입식 등록 완료! (구입비용합계: {total_init_cost:,}원)")
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
                            st.success(f"개체 '{target_id}' → '{new_status}' 변경 완료")
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
                                
                                st.success(f"개체 {len(moving_cids)}마리 → {move_b} {move_p}번 우방으로 이동 완료")
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
                        st.success("질병 기록 등록 완료")
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
                st.dataframe(df_disease, width="stretch", hide_index=True)
    
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
            
        st.markdown(f"**총 {len(df_all_cattle)}건**")
            
        # 금액 관련 컬럼 천단위 콤마 표시
        money_cols = ['구입금액', '수수료', '운송료', '구입비용합계', '보험가입금액', '보험료']
        for col in money_cols:
            if col in df_all_cattle.columns:
                df_all_cattle[col] = df_all_cattle[col].apply(lambda x: f"{int(x):,}" if pd.notnull(x) and str(x).strip() != '' else "")
            
        df_all_cattle.insert(0, "선택", False)
        disabled_cols = [c for c in df_all_cattle.columns if c != "선택"]
        edited_all_cattle_df = st.data_editor(
            df_all_cattle,
            width="stretch",
            hide_index=True,
            disabled=disabled_cols,
            num_rows="fixed",
            key="all_cattle_move_editor",
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
                    st.success(f"개체 {len(moving_cids)}마리 → {bulk_move_b} {bulk_move_p}번 우방으로 이동 완료")
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
        st.dataframe(df_cattle, use_container_width=True, hide_index=True)
    with col_b:
        st.subheader("품목 및 재고 상태")
        st.caption("매입 시마다 이동평균단가가 자동으로 갱신됩니다.")
        df_item = pd.read_sql("SELECT item_code as 품목코드, item_name as 품목명, category as 분류, unit as 단위, current_stock as 현재재고, moving_avg_price as 이동평균단가 FROM item_master", conn)
        st.dataframe(df_item, use_container_width=True, hide_index=True)

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
            submitted_item = st.form_submit_button("품목 등록", type="primary", use_container_width=True)
            if submitted_item:
                if new_item_name:
                    try:
                        write_conn = db_connect(DB_FILE)
                        # 제출 시점에 한 번 더 최신 코드를 확인하여 동시 접속 시 충돌 방지
                        cur_w = write_conn.cursor()
                        cur_w.execute("SELECT item_code FROM item_master")
                        curr_codes = [r[0] for r in cur_w.fetchall()]
                        m_num = 0
                        for c in curr_codes:
                            d = ''.join(filter(str.isdigit, str(c)))
                            if d: m_num = max(m_num, int(d))
                        final_item_code = f"ITEM{m_num + 1}"

                        write_conn.execute(
                            "INSERT INTO item_master (item_code, item_name, category, unit, current_stock, moving_avg_price) VALUES (?, ?, ?, ?, 0, 0)",
                            (final_item_code, new_item_name, new_item_category, new_item_unit)
                        )
                        write_conn.commit()
                        write_conn.close()
                        st.success(f"품목 '{new_item_name}' ({final_item_code})이 등록되었습니다.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        write_conn.rollback(); write_conn.close()
                        st.error("품목 등록 중 오류가 발생했습니다.")
                else:
                    st.warning("품목명을 입력하세요.")
        
        st.markdown("---")
        st.markdown("##### 등록된 품목 목록 (체크박스로 삭제 가능)")
        df_items_all = pd.read_sql("SELECT item_code as 품목코드, item_name as 품목명, category as 분류, unit as 단위, current_stock as 현재재고, moving_avg_price as 이동평균단가 FROM item_master", conn)
        df_items_all.insert(0, "삭제", False)

        edited_item_df = st.data_editor(
            df_items_all,
            use_container_width=True,
            hide_index=True,
            disabled=["품목코드", "현재재고", "이동평균단가"],
            column_config={
                "단위": st.column_config.SelectboxColumn("단위", options=["kg", "ml", "개"]),
            },
            num_rows="dynamic",
            key="item_master_editor"
        )

        if st.button("품목 수정 사항 저장", type="secondary", use_container_width=True):
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
            st.success("품목 내역이 업데이트 되었습니다.")
            st.rerun()
    
    with col_right:
        st.subheader("🚚 매입(입고) 등록")
        st.caption("사료·조사료·약품을 매입하면 재고와 이동평균단가가 자동 갱신됩니다.")
        
        # 품목 목록 가져오기 (단위는 품목 등록 시 정한 값을 그대로 사용한다)
        items_df = pd.read_sql("SELECT item_code, item_name, category, unit FROM item_master", conn)
        if items_df.empty:
            st.info("먼저 좌측에서 품목을 등록해 주세요.")
        else:
            item_options = {f"{r['item_name']} ({r['item_code']}) [{r['category']} · {r['unit'] or '단위 미지정'}]": r['item_code'] for _, r in items_df.iterrows()}
            item_units = {r['item_code']: r['unit'] for _, r in items_df.iterrows()}

            # 매입수량/총매입금액은 천단위 콤마를 보여줘야 해서 st.number_input이 아닌
            # text_input + on_change 콜백으로 처리한다. 콤마 입력창은 st.form 안에서는
            # 제출 전까지 rerun이 안 돼 즉시 반영되지 않으므로, 이 폼은 st.form을 쓰지 않는다.
            if st.session_state.pop("_reset_purchase_fields", False):
                st.session_state["purchase_qty_text"] = ""
                st.session_state["purchase_amount_text"] = ""

            purchase_item_label = st.selectbox("매입 품목", list(item_options.keys()), key="purchase_item_sel")
            purchase_date = st.date_input("매입일자", key="purchase_date_input")
            col_q, col_a2 = st.columns(2)
            with col_q:
                st.text_input(
                    "매입수량", key="purchase_qty_text", placeholder="예: 1,000",
                    on_change=format_thousands_input, args=("purchase_qty_text",),
                )
            with col_a2:
                st.text_input(
                    "총매입금액 (원)", key="purchase_amount_text", placeholder="예: 3,000,000",
                    on_change=format_thousands_input, args=("purchase_amount_text",),
                )

            if st.button("매입 등록", type="primary", use_container_width=True, key="submit_purchase_btn"):
                purchase_qty = parse_thousands_input("purchase_qty_text")
                purchase_amount = parse_thousands_input("purchase_amount_text")
                if purchase_qty > 0 and purchase_amount > 0:
                    selected_item_code = item_options[purchase_item_label]
                    purchase_unit = item_units.get(selected_item_code) or ""
                    write_conn = db_connect(DB_FILE)
                    write_conn.execute(
                        "INSERT INTO purchase (purchase_date, item_code, quantity, unit, total_amount) VALUES (?, ?, ?, ?, ?)",
                        (purchase_date.isoformat(), selected_item_code, purchase_qty, purchase_unit, purchase_amount)
                    )
                    write_conn.commit()
                    write_conn.close()
                    unit_price = purchase_amount / purchase_qty
                    st.session_state["_reset_purchase_fields"] = True
                    st.success(f"매입 완료! {purchase_item_label} | {purchase_qty:,.1f} {purchase_unit} | {purchase_amount:,}원 (단가 {unit_price:,.0f}원)")
                    st.rerun()
                else:
                    st.warning("수량과 금액을 올바르게 입력하세요.")
        
        st.markdown("---")
        st.markdown("##### 매입 내역 (체크박스로 삭제 가능)")
        df_purchase = pd.read_sql("""
            SELECT p.purchase_id as 매입ID, p.purchase_date as 매입일자, 
                   p.item_code as 품목코드, i.item_name as 품목명,
                   p.quantity as 수량, p.unit as 단위, p.total_amount as 총금액,
                   ROUND(p.total_amount / p.quantity, 0) as 단가
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
            use_container_width=True,
            hide_index=True,
            disabled=["매입ID", "품목명", "단가"],
            num_rows="dynamic",
            key="purchase_editor"
        )
        
        if st.button("매입 수정 사항 저장", type="secondary", use_container_width=True):
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
            st.success("매입 내역이 업데이트 및 재고가 재계산 되었습니다.")
            st.rerun()

with tab2:
    st.subheader("월말 비용 등록 및 조회")
    st.markdown("월말에 재고 조사 후, 시험군별 품목 사용량과 농장 고정비를 등록합니다.")
    
    col_c, col_d = st.columns(2)
    
    with col_c:
        st.markdown("##### 🌾 시험군별 사용량 등록 (변동비)")
        
        # 시험군 및 품목 목록
        groups_df = pd.read_sql("SELECT test_group_code, test_name FROM testgroup_master", conn)
        items_df2 = pd.read_sql("SELECT item_code, item_name, category, moving_avg_price FROM item_master", conn)
        
        if groups_df.empty or items_df2.empty:
            st.info("시험군 또는 품목이 등록되어 있지 않습니다.")
        else:
            group_options = {f"{r['test_name']} ({r['test_group_code']})": r['test_group_code'] for _, r in groups_df.iterrows()}
            item_options2 = {f"{r['item_name']} ({r['item_code']}) [{r['category']}]": r['item_code'] for _, r in items_df2.iterrows()}
            
            # 정산연월은 같은 달에 품목을 여러 번 등록하는 경우가 많아서, 폼이
            # clear_on_submit 으로 초기화되어도 마지막에 등록한 연월이 유지되도록 한다.
            if "last_usage_month" not in st.session_state:
                st.session_state["last_usage_month"] = "2023-10"

            with st.form("add_usage_form", clear_on_submit=True):
                usage_month = st.text_input("정산연월", value=st.session_state["last_usage_month"], help="형식: YYYY-MM")
                usage_group_label = st.selectbox("시험군", list(group_options.keys()))
                usage_item_label = st.selectbox("사용 품목", list(item_options2.keys()))
                usage_qty = st.number_input("총 사용량 (kg/개)", min_value=0.01, step=1.0, format="%.2f")

                submitted_usage = st.form_submit_button("사용량 등록", type="primary", width="stretch")
                if submitted_usage:
                    if usage_qty > 0 and usage_month:
                        sel_group = group_options[usage_group_label]
                        sel_item = item_options2[usage_item_label]
                        # 현재 이동평균단가 조회
                        # (moving_avg_price가 정수값이면 pandas가 numpy.int64로 읽어오는데,
                        #  sqlite3가 이를 숫자로 인식하지 못하고 그대로 바이너리로 저장해버려
                        #  float()로 순수 파이썬 숫자로 바꿔서 넘긴다)
                        avg_price = float(items_df2[items_df2['item_code'] == sel_item]['moving_avg_price'].iloc[0])
                        calc_amount = round(usage_qty * avg_price, 2)

                        write_conn = db_connect(DB_FILE)
                        write_conn.execute(
                            "INSERT INTO monthly_usage (settlement_month, test_group_code, item_code, total_usage, applied_price, calculated_amount) VALUES (?, ?, ?, ?, ?, ?)",
                            (usage_month, sel_group, sel_item, usage_qty, avg_price, calc_amount)
                        )
                        write_conn.commit()
                        write_conn.close()
                        st.session_state["last_usage_month"] = usage_month
                        st.success(f"등록 완료! {usage_group_label} | {usage_item_label} | {usage_qty:,.1f} 사용 | 적용단가 {avg_price:,.0f}원 | 산출액 {calc_amount:,.0f}원")
                        st.rerun()
                    else:
                        st.warning("정산연월과 사용량을 올바르게 입력하세요.")
        
        st.markdown("---")
        st.markdown("##### 등록된 사용 내역")
        df_usage = pd.read_sql("""
            SELECT u.usage_id as ID, u.settlement_month as 정산연월,
                   t.test_name as 시험군, i.item_name as 품목명,
                   u.total_usage as 사용량, u.applied_price as 적용단가,
                   u.calculated_amount as 산출총액
            FROM monthly_usage u
            JOIN testgroup_master t ON u.test_group_code = t.test_group_code
            JOIN item_master i ON u.item_code = i.item_code
            ORDER BY u.settlement_month DESC, t.test_name
        """, conn)
        df_usage.insert(0, "삭제", False)

        if "usage_editor" in st.session_state:
            edits = st.session_state["usage_editor"].get("edited_rows", {})
            for row_idx, changes in edits.items():
                row_idx = int(row_idx)
                if row_idx < len(df_usage):
                    new_qty = changes.get("사용량", df_usage.at[row_idx, "사용량"])
                    new_price = changes.get("적용단가", df_usage.at[row_idx, "적용단가"])
                    if pd.notna(new_qty) and pd.notna(new_price):
                        df_usage.at[row_idx, "산출총액"] = round(float(new_qty) * float(new_price))
            
            added = st.session_state["usage_editor"].get("added_rows", [])
            for row in added:
                qty = row.get("사용량", 0)
                price = row.get("적용단가", 0)
                row["산출총액"] = round(float(qty) * float(price))

        edited_usage_df = st.data_editor(
            df_usage,
            width="stretch",
            hide_index=True,
            disabled=["ID", "시험군", "품목명", "산출총액"],
            num_rows="dynamic",
            key="usage_editor",
        )

        if st.button("사용 내역 수정 사항 저장", type="secondary", width="stretch"):
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
            st.success("사용 내역이 업데이트 되었습니다.")
            st.rerun()
    
    with col_d:
        st.markdown("##### ⚡ 농장 고정비 등록")

        if "last_fc_month" not in st.session_state:
            st.session_state["last_fc_month"] = "2023-10"

        with st.form("add_fixedcost_form", clear_on_submit=True):
            fc_month = st.text_input("정산연월 ", value=st.session_state["last_fc_month"], help="형식: YYYY-MM")
            fc_item = st.selectbox("지출 항목", ["인건비", "전기세", "시험사양수고비", "CCTV사용료", "우수등급장려금", "기타"])
            fc_amount = st.number_input("총 청구금액 (원)", min_value=0, step=10000)

            submitted_fc = st.form_submit_button("고정비 등록", type="primary", width="stretch")
            if submitted_fc:
                if fc_amount > 0 and fc_month:
                    write_conn = db_connect(DB_FILE)
                    write_conn.execute(
                        "INSERT INTO monthly_fixedcost (settlement_month, expense_item, total_billed_amount) VALUES (?, ?, ?)",
                        (fc_month, fc_item, fc_amount)
                    )
                    write_conn.commit()
                    write_conn.close()
                    st.session_state["last_fc_month"] = fc_month
                    st.success(f"등록 완료! [{fc_month}] {fc_item} | {fc_amount:,}원")
                    st.rerun()
                else:
                    st.warning("정산연월과 금액을 올바르게 입력하세요.")
        
        st.markdown("---")
        st.markdown("##### 등록된 고정비 내역 (체크박스로 삭제 가능)")
        df_fixed = pd.read_sql("""
            SELECT fixed_cost_id as ID, settlement_month as 정산연월,
                   expense_item as 지출항목, total_billed_amount as 총청구금액
            FROM monthly_fixedcost
            ORDER BY settlement_month DESC
        """, conn)
        df_fixed.insert(0, "삭제", False)
        
        edited_fc_df = st.data_editor(
            df_fixed,
            width="stretch",
            hide_index=True,
            disabled=["ID"],
            num_rows="dynamic",
            key="fixedcost_editor"
        )
        
        if st.button("고정비 수정 사항 저장", type="secondary", width="stretch"):
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
            st.success("고정비 내역이 업데이트 되었습니다.")
            st.rerun()

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("### 🚀 월말 정산(일할계산) 실행")
    st.markdown("아래 버튼을 누르면 위에서 등록한 변동비·고정비를 분석하여, 이번 달 사육 이력이 있는 각 개체에 **실제 사육일수에 비례해(일할계산)** 변동비와 고정비를 배분합니다.")
    
    target_month = st.text_input("정산 대상 연월 (예: 2023-10)", value="2023-10", key="calc_target_month")
    
    # 정산 전 요약 미리보기
    preview_usage = pd.read_sql("""
        SELECT t.test_name as 시험군, SUM(u.calculated_amount) as 변동비_합계
        FROM monthly_usage u
        JOIN testgroup_master t ON u.test_group_code = t.test_group_code
        WHERE u.settlement_month = ?
        GROUP BY t.test_name
    """, conn, params=(target_month,))
    preview_fixed = pd.read_sql("SELECT SUM(total_billed_amount) as 고정비_합계 FROM monthly_fixedcost WHERE settlement_month = ?", conn, params=(target_month,))
    
    if not preview_usage.empty or (not preview_fixed.empty and pd.notna(preview_fixed.iloc[0]['고정비_합계'])):
        st.markdown(f"**[{target_month}] 정산 대상 비용 요약:**")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.dataframe(preview_usage, width="stretch", hide_index=True)
        with col_p2:
            fixed_val = preview_fixed.iloc[0]['고정비_합계'] if pd.notna(preview_fixed.iloc[0]['고정비_합계']) else 0
            st.metric("고정비 합계", f"{fixed_val:,.0f} 원")
    
    if st.button("🚀 정산 실행(일할계산) 및 누적원가 반영", type="primary"):
        success, msg = distribute_monthly_costs(DB_FILE, target_month)
        if success:
            st.success(msg)
            st.balloons()
        else:
            st.error(msg)
            
    st.markdown("---")
    st.subheader(f"[{target_month}] 개체별 원가 적재 결과 (Cattle_Cost_Log)")
    try:
        df_log = pd.read_sql("SELECT cattle_id, settlement_month, allocated_variable_cost as 변동비_할당, allocated_fixed_cost as 고정비_할당, (allocated_variable_cost + allocated_fixed_cost) as 당월_추가원가 FROM cattle_cost_log WHERE settlement_month=?", conn, params=(target_month,))
        if df_log.empty:
            st.info("해당 연월에 아직 정산된 내역이 없습니다.")
        else:
            st.dataframe(df_log, width="stretch", hide_index=True)
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
