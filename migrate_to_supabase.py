"""로컬 SQLite 농장 DB를 Supabase(PostgreSQL)로 옮긴다.

사용법 (프로젝트 폴더에서):
    python migrate_to_supabase.py                 # 확인만 (옮길 파일과 행 수 표시)
    python migrate_to_supabase.py --run           # 실제로 옮기기
    python migrate_to_supabase.py --run --drop-extra   # + 농장이 아닌 스키마(테스트용 등) 삭제

- 원본은 실제 DB 폴더(ERP_DB_DIR, 기본: 사용자폴더\\시험농장DB)의 farms.json 에 등록된 농장 DB다.
  프로젝트 폴더(구글 드라이브)에 남아 있는 옛 사본은 쓰지 않는다.
- 농장마다 스키마를 통째로 지우고 앱과 같은 스키마(기본키·제약·트리거 포함)로 다시 만든 뒤 데이터를 넣는다.
  한 농장 단위로 한 트랜잭션이라, 도중에 실패하면 그 농장은 이전 상태 그대로 남는다.
- 원본 SQLite 파일은 읽기만 한다.
"""
import json
import os
import sqlite3
import sys

import toml

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.chdir(HERE)

if not os.environ.get("DATABASE_URL"):
    os.environ["DATABASE_URL"] = toml.load(os.path.join(HERE, ".streamlit", "secrets.toml"))["DATABASE_URL"]

import db_adapter  # noqa: E402
from db_schema import DB_TABLES, PG_DDL  # noqa: E402

DB_DIR = os.environ.get("ERP_DB_DIR") or os.path.join(os.path.expanduser("~"), "시험농장DB")
SYSTEM_SCHEMAS = {
    "public", "auth", "storage", "realtime", "extensions", "graphql", "graphql_public", "vault",
    "pgsodium", "pgsodium_masks", "supabase_functions", "supabase_migrations", "net", "cron",
    "_realtime", "pgbouncer", "information_schema",
}


def farm_db_files():
    with open(os.path.join(DB_DIR, "farms.json"), encoding="utf-8") as f:
        farms = json.load(f)
    return {name: os.path.join(DB_DIR, cfg.get("db_filename", "erp_%s.db" % name)) for name, cfg in farms.items()}


def local_counts(path):
    conn = sqlite3.connect(path)
    try:
        tables = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        return {t: conn.execute("SELECT COUNT(*) FROM %s" % t).fetchone()[0] for t in DB_TABLES if t in tables}
    finally:
        conn.close()


def remote_counts(db_file):
    conn = db_adapter.connect(db_file)
    try:
        return {t: conn.execute("SELECT COUNT(*) FROM %s" % t).fetchone()[0] for t in DB_TABLES}
    finally:
        conn.close()


def extra_schemas(keep):
    conn = db_adapter.connect("public")
    try:
        rows = conn.execute("SELECT schema_name FROM information_schema.schemata").fetchall()
    finally:
        conn.close()
    return sorted(s for (s,) in rows if s not in keep and s not in SYSTEM_SCHEMAS and not s.startswith("pg_"))


def main():
    run = "--run" in sys.argv
    drop_extra = "--drop-extra" in sys.argv
    files = farm_db_files()
    print("원본 폴더:", DB_DIR)

    for farm, path in files.items():
        schema = db_adapter.schema_for(path)
        if not os.path.exists(path):
            print("\n[%s] 파일 없음: %s → 건너뜀" % (farm, path))
            continue
        counts = local_counts(path)
        print("\n[%s] %s → 스키마 '%s'" % (farm, os.path.basename(path), schema))
        print("  원본 행 수:", counts)
        if not run:
            continue
        db_adapter.import_from_sqlite(path, path, PG_DDL, DB_TABLES, trigger_tables=["purchase"])
        after = remote_counts(path)
        print("  Supabase 행 수:", after)
        diff = {t: (counts.get(t, 0), after.get(t, 0)) for t in DB_TABLES if counts.get(t, 0) != after.get(t, 0)}
        print("  검증:", "일치" if not diff else "불일치 %s" % diff)

    keep = {db_adapter.schema_for(p) for p in files.values()}
    extras = extra_schemas(keep)
    if extras:
        print("\n농장이 아닌 스키마:", extras)
        if run and drop_extra:
            conn = db_adapter.connect("public")
            try:
                for s in extras:
                    conn.execute('DROP SCHEMA "%s" CASCADE' % s.replace('"', '""'))
                conn.commit()
            finally:
                conn.close()
            print("  → 삭제함")
        elif not drop_extra:
            print("  (지우려면 --drop-extra 를 함께 지정)")

    if not run:
        print("\n확인만 했습니다. 실제로 옮기려면 --run 을 붙여 실행하세요.")


if __name__ == "__main__":
    main()
