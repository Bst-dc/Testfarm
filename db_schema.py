"""농장 DB 스키마 정의. erp_ui.py 와 migrate_to_supabase.py 가 함께 쓴다."""
import re

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
    total_amount NUMERIC(12, 2) NOT NULL,
    feed_zone TEXT
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

-- 사료 재고 예측: 사료빈 구역(동 묶음), 구역별 사료 변경일별 두당 일급여량, 사료빈 실제 재고 확인값
CREATE TABLE IF NOT EXISTS feed_zone (
    zone_id INTEGER PRIMARY KEY AUTOINCREMENT,
    zone_name TEXT NOT NULL UNIQUE,
    buildings TEXT,
    test_groups TEXT
);

CREATE TABLE IF NOT EXISTS feed_schedule (
    zone_name TEXT NOT NULL,
    change_date DATE NOT NULL,
    kg_per_head NUMERIC(10, 2) NOT NULL,
    PRIMARY KEY (zone_name, change_date)
);

CREATE TABLE IF NOT EXISTS feed_stock_check (
    zone_name TEXT NOT NULL,
    check_date DATE NOT NULL,
    quantity NUMERIC(12, 2) NOT NULL,
    PRIMARY KEY (zone_name, check_date)
);

-- 위탁농장 월말 점검표: 달마다 점검횟수·위탁사육자·항목별 평점과 평가이유(items, JSON)
CREATE TABLE IF NOT EXISTS inspection_sheet (
    sheet_month TEXT PRIMARY KEY,
    inspect_count INTEGER,
    farmer TEXT,
    items TEXT
);

CREATE TRIGGER IF NOT EXISTS trg_after_insert_purchase
AFTER INSERT ON purchase
FOR EACH ROW
BEGIN
    UPDATE item_master
    SET current_stock = current_stock + NEW.quantity,
        moving_avg_price = CASE 
            WHEN current_stock + NEW.quantity > 0 
            THEN ROUND(((current_stock * moving_avg_price) + NEW.total_amount) * 1.0 / (current_stock + NEW.quantity), 2)
            ELSE 0 
        END
    WHERE item_code = NEW.item_code;
END;
"""

# ========== Supabase(PostgreSQL) 스키마 ==========
# SQLITE_DDL 을 기준으로 만든다. SQLite 와 똑같이 동작하도록
#   - 날짜는 'YYYY-MM-DD' 문자열(TEXT)로 둔다 (앱이 substr/문자열 비교로 연월을 다룬다).
#   - NUMERIC(p, s) 의 자릿수 제한을 없앤다 (SQLite 는 소수점 자리를 자르지 않는다).
#   - AUTOINCREMENT → SERIAL, 트리거는 PL/pgSQL 함수로 바꾼다.
PG_PURCHASE_TRIGGER = """
CREATE OR REPLACE FUNCTION trg_after_insert_purchase_fn() RETURNS trigger
SET search_path FROM CURRENT
AS $$
BEGIN
    UPDATE item_master
    SET current_stock = current_stock + NEW.quantity,
        moving_avg_price = CASE
            WHEN current_stock + NEW.quantity > 0
            THEN ROUND(((current_stock * moving_avg_price) + NEW.total_amount) * 1.0 / (current_stock + NEW.quantity), 2)
            ELSE 0
        END
    WHERE item_code = NEW.item_code;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE OR REPLACE TRIGGER trg_after_insert_purchase
AFTER INSERT ON purchase
FOR EACH ROW EXECUTE FUNCTION trg_after_insert_purchase_fn();
"""
PG_DDL = re.sub(
    r"NUMERIC\(\d+,\s*\d+\)", "NUMERIC",
    re.sub(r"\bDATE\b", "TEXT", SQLITE_DDL[:SQLITE_DDL.index("CREATE TRIGGER")]),
).replace("INTEGER PRIMARY KEY AUTOINCREMENT", "SERIAL PRIMARY KEY") + PG_PURCHASE_TRIGGER

# 백업/복원 때 옮기는 순서 (참조하는 쪽이 뒤)
DB_TABLES = [
    "testgroup_master", "cattle", "item_master", "disease_record", "purchase",
    "monthly_usage", "monthly_fixedcost", "cattle_cost_log", "cattle_item_usage_log",
    "feed_zone", "feed_schedule", "feed_stock_check", "inspection_sheet",
]
