"""sqlite3 모듈처럼 쓸 수 있는 Supabase(PostgreSQL) 어댑터.

erp_ui.py 는 원래 SQLite 전용으로 작성되어 있어서, 여기서 SQLite 와 다르게 동작하는 부분을
흡수해 앱 코드를 거의 그대로 쓸 수 있게 한다.

- 농장별 DB 파일(erp_sunsan.db) → 같은 이름의 Postgres 스키마(sunsan) 로 대응시킨다.
- SQL 문법 차이(? 자리표시자, strftime, 작은따옴표 별칭, sqlite_master, PRAGMA)를 변환한다.
- SQLite 처럼 느슨한 타입: 숫자 파라미터를 타입 미정 리터럴로 보내 TEXT 컬럼과 비교해도 오류가 나지 않게 하고,
  NaN 은 NULL 로, NUMERIC 결과는 Decimal 대신 int/float 로 돌려준다.
- Postgres 는 트랜잭션 도중 오류가 한 번 나면 이후 명령을 모두 거부하지만, 앱은 SQLite 처럼
  "중복은 건너뛰고 계속"하는 코드가 많다. 명령마다 SAVEPOINT 를 걸어 오류 난 명령만 되돌린다.
- 연결은 매번 새로 맺지 않고 재사용한다(Supabase 는 연결 한 번에 수백 ms 가 걸린다).
"""
import datetime
import math
import os
import re
import threading
import time
import warnings

import numpy as np
import psycopg2
import psycopg2.extensions as ext
import psycopg2.extras

# pandas 는 sqlite3/SQLAlchemy 가 아닌 연결에 매번 경고를 띄우지만, 이 어댑터는 sqlite3 처럼 동작하도록 맞춰 두었다.
warnings.filterwarnings("ignore", message=r"pandas only supports SQLAlchemy")


# ========== sqlite3 호환 예외 ==========
class Error(Exception):
    pass


class DatabaseError(Error):
    pass


class IntegrityError(DatabaseError):
    pass


class OperationalError(DatabaseError):
    pass


ProgrammingError = OperationalError

# SQLite 오류 문구에 맞춰 둔다. 앱은 "UNIQUE" 가 들어 있으면 '이미 등록되어 건너뜀'으로 센다.
_INTEGRITY_PREFIX = {
    "23505": "UNIQUE constraint failed",
    "23514": "CHECK constraint failed",
    "23503": "FOREIGN KEY constraint failed",
    "23502": "NOT NULL constraint failed",
}


def _wrap_error(e):
    code = getattr(e, "pgcode", None) or ""
    msg = (getattr(e, "pgerror", None) or str(e)).strip()
    if code.startswith("23"):
        return IntegrityError("%s: %s" % (_INTEGRITY_PREFIX.get(code, "constraint failed"), msg))
    return OperationalError(msg)


# ========== 값 변환 (파이썬 → SQL) ==========
# 숫자를 '3' 처럼 타입 미정 리터럴로 보내면 Postgres 가 비교·저장 대상 컬럼 타입에 맞춰 해석한다.
# (SQLite 는 TEXT 컬럼 building 과 숫자 3 을 비교해도 되지만, Postgres 는 text = integer 오류를 낸다.)
def _adapt_int(v):
    return ext.AsIs("'%d'" % int(v))


def _adapt_float(v):
    v = float(v)
    if math.isnan(v) or math.isinf(v):
        return ext.AsIs("NULL")  # SQLite 도 NaN 은 NULL 로 저장한다
    return ext.AsIs("'%r'" % v)


def _adapt_bool(v):
    return ext.AsIs("'1'" if v else "'0'")  # SQLite 처럼 1/0 으로 저장


def _adapt_date(v):
    return ext.QuotedString(v.isoformat())


def _adapt_datetime(v):
    return ext.QuotedString(v.isoformat(" "))


ext.register_adapter(bool, _adapt_bool)
ext.register_adapter(np.bool_, _adapt_bool)
ext.register_adapter(int, _adapt_int)
for _t in (np.int8, np.int16, np.int32, np.int64, np.uint8, np.uint16, np.uint32, np.uint64):
    ext.register_adapter(_t, _adapt_int)
for _t in (float, np.float16, np.float32, np.float64):
    ext.register_adapter(_t, _adapt_float)
ext.register_adapter(datetime.date, _adapt_date)
ext.register_adapter(datetime.datetime, _adapt_datetime)
try:
    import pandas as _pd

    ext.register_adapter(_pd.Timestamp, lambda v: _adapt_datetime(v.to_pydatetime()))
    ext.register_adapter(type(_pd.NaT), lambda v: ext.AsIs("NULL"))
except ImportError:
    pass


# ========== 값 변환 (SQL → 파이썬) ==========
# NUMERIC 은 기본적으로 Decimal 로 오는데, 앱은 float 와 섞어 계산한다(Decimal * float 는 오류).
# SQLite 의 NUMERIC 처럼 정수로 떨어지면 int, 아니면 float 로 돌려준다.
def _cast_numeric(value, cur):
    if value is None:
        return None
    if value in ("NaN", "Infinity", "-Infinity"):
        return float(value)
    if "." not in value and "e" not in value.lower():
        return int(value)
    f = float(value)
    return int(f) if f.is_integer() else f


ext.register_type(ext.new_type(ext.DECIMAL.values, "SQLITE_LIKE_NUMERIC", _cast_numeric))


# ========== SQL 변환 ==========
_LITERAL_RE = re.compile(r"('(?:[^']|'')*')")
_TYPE_WORDS = {"TEXT", "INTEGER", "INT", "BIGINT", "REAL", "NUMERIC", "DATE", "TIMESTAMP",
               "VARCHAR", "FLOAT", "DOUBLE", "BLOB", "BOOLEAN", "SERIAL"}
_SQLITE_MASTER = (
    "(SELECT table_name AS name, "
    "CASE table_type WHEN 'VIEW' THEN 'view' ELSE 'table' END AS type, "
    "NULL::text AS sql FROM information_schema.tables "
    "WHERE table_schema = current_schema()) AS sqlite_master"
)
# 날짜는 'YYYY-MM-DD' 문자열로 저장하므로 strftime 은 앞부분 자르기와 같다.
_STRFTIME_LEN = {"%Y-%m-%d": 10, "%Y-%m": 7, "%Y": 4}


def _strftime(m):
    fmt, expr = m.group(1), m.group(2)
    if fmt in _STRFTIME_LEN:
        return "substr(%s, 1, %d)" % (expr, _STRFTIME_LEN[fmt])
    pg_fmt = fmt.replace("%Y", "YYYY").replace("%m", "MM").replace("%d", "DD") \
                .replace("%H", "HH24").replace("%M", "MI").replace("%S", "SS")
    return "to_char((%s)::timestamp, '%s')" % (expr, pg_fmt)


def _quote_alias(m):
    word = m.group(2)
    # Postgres 는 따옴표 없는 별칭을 소문자로 바꾼다(ID → id). 대문자가 있는 별칭은 그대로 두도록 감싼다.
    if word.upper() in _TYPE_WORDS or not re.search(r"[A-Z]", word):
        return m.group(0)
    return '%s "%s"' % (m.group(1), word)


def _translate(sql):
    stripped = sql.strip()
    m = re.match(r"PRAGMA\s+table_info\((.+?)\)", stripped, re.I)
    if m:
        table = m.group(1).strip("'\" ")
        return ("SELECT ordinal_position - 1, column_name, data_type, 0, NULL, 0 "
                "FROM information_schema.columns WHERE table_schema = current_schema() "
                "AND table_name = '%s' ORDER BY ordinal_position" % table.replace("'", "''"))
    if re.match(r"PRAGMA\b", stripped, re.I):
        return "SELECT 1"  # busy_timeout / journal_mode 등 SQLite 설정은 Postgres 에 해당 없음

    sql = re.sub(r"\bAS\s+'([^']*)'", r'AS "\1"', sql, flags=re.I)
    sql = re.sub(r"strftime\(\s*'([^']*)'\s*,\s*([^()]+?)\s*\)", _strftime, sql, flags=re.I)
    parts = _LITERAL_RE.split(sql)
    for i in range(0, len(parts), 2):  # 짝수 번째 = 문자열 리터럴 바깥
        p = parts[i]
        p = re.sub(r"\bsqlite_master\b", _SQLITE_MASTER, p)
        p = re.sub(r"INTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT", "SERIAL PRIMARY KEY", p, flags=re.I)
        p = re.sub(r"\b(AS)\s+([^\W\d]\w*)\b(?!\s*\()", _quote_alias, p, flags=re.I)
        parts[i] = p
    return "".join(parts)


def _bind(sql):
    """? 자리표시자를 %s 로 바꾸고, psycopg2 가 서식 문자로 오해하지 않게 % 를 %% 로 이스케이프한다."""
    parts = _LITERAL_RE.split(sql)
    for i, p in enumerate(parts):
        p = p.replace("%", "%%")
        if i % 2 == 0:
            p = p.replace("?", "%s")
        parts[i] = p
    return "".join(parts)


# ========== 연결 설정 ==========
def database_url():
    url = os.environ.get("DATABASE_URL")
    if url:
        return url
    try:
        import streamlit as st

        return st.secrets.get("DATABASE_URL")
    except Exception:
        return None


def enabled():
    """DATABASE_URL 이 설정돼 있으면 Supabase 를 쓴다. 없으면 앱은 로컬 SQLite 파일을 쓴다."""
    return bool(database_url())


def schema_for(db_file):
    """'.../erp_sunsan.db' → 'sunsan'. 농장 하나가 스키마 하나다."""
    name = os.path.splitext(os.path.basename(str(db_file)))[0]
    if name.startswith("erp_"):
        name = name[len("erp_"):]
    name = name.replace(" ", "_").replace('"', "")
    return name or "public"


def _ident(name):
    return '"%s"' % name.replace('"', '""')


class _Pool:
    """Streamlit 은 버튼을 누를 때마다 스크립트 전체를 다시 실행하고 그때마다 연결을 여러 번 연다.
    Supabase 연결은 한 번 맺는 데 수백 ms 가 걸리므로, 닫힌 연결을 버리지 않고 모아 두었다 다시 쓴다."""

    MAX_IDLE = 4
    HEALTH_CHECK_AFTER = 30  # 초. 이보다 오래 놀던 연결은 끊겼을 수 있으니 확인 후 쓴다.

    def __init__(self):
        self._idle = []
        self._lock = threading.Lock()

    def get(self):
        while True:
            with self._lock:
                if not self._idle:
                    break
                raw, schema, since = self._idle.pop()
            if raw.closed:
                continue
            if time.time() - since > self.HEALTH_CHECK_AFTER:
                try:
                    with raw.cursor() as c:
                        c.execute("SELECT 1")
                    raw.rollback()
                except psycopg2.Error:
                    _close_quietly(raw)
                    continue
            return raw, schema
        url = database_url()
        if not url:
            raise OperationalError("DATABASE_URL 이 설정되지 않았습니다 (.streamlit/secrets.toml 또는 Streamlit Cloud Secrets).")
        try:
            raw = psycopg2.connect(url, connect_timeout=15, application_name="testfarm-erp",
                                   keepalives=1, keepalives_idle=30)
        except psycopg2.Error as e:
            raise OperationalError("Supabase 연결 실패: %s" % e) from None
        return raw, None

    def put(self, raw, schema):
        if raw.closed:
            return
        try:
            if raw.get_transaction_status() != ext.TRANSACTION_STATUS_IDLE:
                raw.rollback()
        except psycopg2.Error:
            _close_quietly(raw)
            return
        with self._lock:
            if len(self._idle) < self.MAX_IDLE:
                self._idle.append((raw, schema, time.time()))
                return
        _close_quietly(raw)


def _close_quietly(raw):
    try:
        raw.close()
    except Exception:
        pass


_POOL = _Pool()


def _set_schema(raw, schema):
    raw.autocommit = True
    try:
        with raw.cursor() as c:
            c.execute("SET search_path TO %s" % _ident(schema))
    finally:
        raw.autocommit = False


# ========== sqlite3 호환 연결 / 커서 ==========
class Cursor:
    arraysize = 1
    lastrowid = None

    def __init__(self, conn):
        self._conn = conn
        self._cur = conn._raw.cursor()
        self.description = None
        self.rowcount = -1

    def execute(self, sql, params=None):
        q = _translate(sql)
        if params is not None:
            if isinstance(params, dict):
                raise ProgrammingError("이름 붙은 파라미터(:name)는 지원하지 않습니다.")
            params = tuple(params)
            q = _bind(q)
        self._conn._run(self._cur, q, params)
        self.description = self._cur.description
        self.rowcount = self._cur.rowcount
        return self

    def executemany(self, sql, seq_of_parameters):
        total = 0
        for params in seq_of_parameters:
            self.execute(sql, params)
            total += max(self.rowcount, 0)
        self.rowcount = total
        return self

    def executescript(self, script):
        self._conn._run(self._cur, _translate(script), None)
        return self

    def fetchone(self):
        return self._cur.fetchone() if self._cur.description else None

    def fetchall(self):
        return self._cur.fetchall() if self._cur.description else []

    def fetchmany(self, size=None):
        if not self._cur.description:
            return []
        return self._cur.fetchmany(size or self.arraysize)

    def __iter__(self):
        return iter(self.fetchall())

    def close(self):
        try:
            self._cur.close()
        except Exception:
            pass


class Connection:
    def __init__(self, raw, schema):
        self._raw = raw
        self._schema = schema
        self._savepoint = False  # 현재 트랜잭션 안에 SAVEPOINT 가 걸려 있는지
        self._closed = False

    def _check(self):
        if self._closed:
            raise ProgrammingError("Cannot operate on a closed database.")

    def _run(self, cur, sql, params):
        self._check()
        # 명령마다 SAVEPOINT 를 새로 건다(같은 왕복에 실어 보내므로 추가 지연 없음).
        # 실패하면 그 명령만 되돌려, 같은 트랜잭션의 이전 작업과 이후 명령이 살아남는다 (SQLite 와 같은 동작).
        prefix = "RELEASE SAVEPOINT sqlite_stmt; SAVEPOINT sqlite_stmt; " if self._savepoint else "SAVEPOINT sqlite_stmt; "
        try:
            cur.execute(prefix + sql, params)
            self._savepoint = True
        except psycopg2.Error as e:
            self._recover()
            raise _wrap_error(e) from None

    def _recover(self):
        raw = self._raw
        if raw.closed:
            return
        try:
            if raw.get_transaction_status() == ext.TRANSACTION_STATUS_INERROR:
                # 새 SAVEPOINT 가 걸리기 전에 실패했을 수도 있으니, 되돌릴 지점이 있을 때만 부분 롤백한다.
                with raw.cursor() as c:
                    c.execute("ROLLBACK TO SAVEPOINT sqlite_stmt")
                self._savepoint = True
        except psycopg2.Error:
            try:
                raw.rollback()
            except psycopg2.Error:
                _close_quietly(raw)
            self._savepoint = False

    def cursor(self):
        self._check()
        return Cursor(self)

    def execute(self, sql, params=None):
        return self.cursor().execute(sql, params)

    def executemany(self, sql, seq_of_parameters):
        return self.cursor().executemany(sql, seq_of_parameters)

    def executescript(self, script):
        self.commit()  # sqlite3 의 executescript 도 먼저 커밋한다
        cur = self.cursor().executescript(script)
        self.commit()
        return cur

    def commit(self):
        self._check()
        try:
            self._raw.commit()
        except psycopg2.Error as e:
            raise _wrap_error(e) from None
        finally:
            self._savepoint = False

    def rollback(self):
        if self._closed or self._raw.closed:
            return
        try:
            self._raw.rollback()
        except psycopg2.Error as e:
            raise _wrap_error(e) from None
        finally:
            self._savepoint = False

    def close(self):
        """커밋하지 않은 변경은 버리고(sqlite3 와 같음) 연결은 재사용을 위해 풀로 돌려준다."""
        if self._closed:
            return
        self._closed = True
        _POOL.put(self._raw, self._schema)

    # sqlite3 와 같은 의미: 블록이 끝나면 커밋/롤백만 하고 연결은 닫지 않는다.
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is None:
            self.commit()
        else:
            self.rollback()
        return False


def connect(db_file, timeout=30, check_same_thread=False, **kwargs):
    """sqlite3.connect() 대신 쓴다. 농장 DB 파일 이름에 해당하는 스키마를 기본 검색 경로로 잡는다."""
    schema = schema_for(db_file)
    raw, current = _POOL.get()
    if current != schema:
        try:
            _set_schema(raw, schema)
        except psycopg2.Error as e:
            _close_quietly(raw)
            raise _wrap_error(e) from None
    return Connection(raw, schema)


# ========== 스키마 관리 ==========
def _raw_connection(schema=None):
    raw, _ = _POOL.get()
    if schema:
        _set_schema(raw, schema)
    return raw


def database_exists(db_file):
    """농장 스키마에 cattle 표가 있으면 '이 농장 DB가 있다'로 본다 (os.path.exists 대용)."""
    raw = _raw_connection()
    try:
        with raw.cursor() as c:
            c.execute("SELECT to_regclass(%s)", (_ident(schema_for(db_file)) + ".cattle",))
            return c.fetchone()[0] is not None
    except psycopg2.Error:
        return False
    finally:
        _POOL.put(raw, None)


def _run_ddl(c, schema, ddl, drop_first):
    if drop_first:
        c.execute("DROP SCHEMA IF EXISTS %s CASCADE" % _ident(schema))
    c.execute("CREATE SCHEMA IF NOT EXISTS %s" % _ident(schema))
    c.execute("SET LOCAL search_path TO %s" % _ident(schema))
    c.execute(ddl)


def ensure_schema(db_file, ddl, extra_sql=()):
    """스키마와 표를 만든다(이미 있으면 그대로). extra_sql 은 나중에 추가된 컬럼 등 보완용 문장."""
    schema = schema_for(db_file)
    raw = _raw_connection()
    try:
        with raw.cursor() as c:
            _run_ddl(c, schema, ddl, drop_first=False)
            for stmt in extra_sql:
                c.execute(stmt)
        raw.commit()
    except psycopg2.Error as e:
        raw.rollback()
        raise _wrap_error(e) from None
    finally:
        _POOL.put(raw, None)


def reset_database(db_file, ddl):
    """농장 스키마를 통째로 지우고 빈 표로 다시 만든다 (SQLite 에서 파일 삭제 후 재생성하던 것과 같다)."""
    schema = schema_for(db_file)
    raw = _raw_connection()
    try:
        with raw.cursor() as c:
            _run_ddl(c, schema, ddl, drop_first=True)
        raw.commit()
    except psycopg2.Error as e:
        raw.rollback()
        raise _wrap_error(e) from None
    finally:
        _POOL.put(raw, None)


# ========== SQLite 파일 ↔ Supabase 복사 (백업 / 복원 / 이관) ==========
_NUMERIC_TYPES = {"integer", "bigint", "smallint", "numeric", "double precision", "real"}


def _pg_columns(c, table):
    c.execute("SELECT column_name, data_type FROM information_schema.columns "
              "WHERE table_schema = current_schema() AND table_name = %s ORDER BY ordinal_position", (table,))
    return c.fetchall()


def _clean_value(v, pg_type):
    if isinstance(v, (bytes, memoryview)):
        raw = bytes(v)
        # 예전 버그로 numpy.int64 가 8바이트 BLOB 으로 저장된 값 (erp_ui._repair_numpy_int_blobs 참고)
        return int.from_bytes(raw, "little", signed=True) if len(raw) == 8 else raw.decode("utf-8", "replace")
    if pg_type in _NUMERIC_TYPES and isinstance(v, str):
        s = v.replace(",", "").strip()
        if not s:
            return None
        try:
            f = float(s)
        except ValueError:
            return None
        return int(f) if pg_type in ("integer", "bigint", "smallint") or f.is_integer() else f
    if pg_type in ("integer", "bigint", "smallint") and isinstance(v, float):
        return None if math.isnan(v) else int(round(v))
    if pg_type == "text" and v is not None and not isinstance(v, str):
        return str(v)
    return v


def import_from_sqlite(db_file, sqlite_path, ddl, tables, trigger_tables=()):
    """SQLite 파일의 데이터로 농장 스키마를 통째로 교체한다. 한 트랜잭션이라 실패하면 아무것도 바뀌지 않는다.

    tables: 참조 순서(부모 → 자식)대로 나열한 표 이름.
    trigger_tables: 옮기는 동안 트리거를 꺼 둘 표 (매입 트리거가 재고를 이중으로 더하지 않도록).
    반환값: {표 이름: 옮긴 행 수}
    """
    import sqlite3 as real_sqlite3

    schema = schema_for(db_file)
    src = real_sqlite3.connect(sqlite_path)
    raw = _raw_connection()
    counts = {}
    try:
        src_tables = {r[0] for r in src.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        with raw.cursor() as c:
            _run_ddl(c, schema, ddl, drop_first=True)
            for t in trigger_tables:
                c.execute("ALTER TABLE %s DISABLE TRIGGER USER" % _ident(t))
            for table in tables:
                if table not in src_tables:
                    counts[table] = 0
                    continue
                src_cols = {r[1] for r in src.execute("PRAGMA table_info(%s)" % table)}
                cols = [(n, t) for n, t in _pg_columns(c, table) if n in src_cols]
                names = [n for n, _ in cols]
                rows = src.execute("SELECT %s FROM %s" % (", ".join(names), table)).fetchall()
                rows = [tuple(_clean_value(v, t) for v, (_, t) in zip(row, cols)) for row in rows]
                if rows:
                    psycopg2.extras.execute_values(
                        c, "INSERT INTO %s (%s) VALUES %%s" % (_ident(table), ", ".join(_ident(n) for n in names)),
                        rows, page_size=500,
                    )
                counts[table] = len(rows)
            for t in trigger_tables:
                c.execute("ALTER TABLE %s ENABLE TRIGGER USER" % _ident(t))
            # SERIAL 컬럼의 다음 번호를 옮긴 데이터의 최댓값 다음으로 맞춘다.
            c.execute("SELECT table_name, column_name, pg_get_serial_sequence(quote_ident(table_name), column_name) "
                      "FROM information_schema.columns WHERE table_schema = current_schema() "
                      "AND column_default LIKE 'nextval(%%'")
            for table, col, seq in c.fetchall():
                c.execute("SELECT setval(%%s, COALESCE((SELECT MAX(%s) FROM %s), 0) + 1, false)"
                          % (_ident(col), _ident(table)), (seq,))
        raw.commit()
        return counts
    except psycopg2.Error as e:
        raw.rollback()
        raise _wrap_error(e) from None
    finally:
        src.close()
        _POOL.put(raw, None)


def export_to_sqlite(db_file, out_path, sqlite_ddl, tables, trigger_names=()):
    """농장 스키마를 SQLite 파일로 내보낸다 (백업). 이 파일은 '백업 파일로 복원'에 그대로 쓸 수 있다."""
    import sqlite3 as real_sqlite3

    schema = schema_for(db_file)
    raw = _raw_connection(schema)
    dst = real_sqlite3.connect(out_path)
    try:
        dst.executescript(sqlite_ddl)
        # 트리거가 켜진 채 매입 내역을 넣으면 재고·단가가 한 번 더 더해진다. 옮긴 뒤 다시 만든다.
        for name in trigger_names:
            dst.execute("DROP TRIGGER IF EXISTS %s" % name)
        with raw.cursor() as c:
            for table in tables:
                cols = [n for n, _ in _pg_columns(c, table)]
                if not cols:
                    continue
                dst_cols = {r[1] for r in dst.execute("PRAGMA table_info(%s)" % table)}
                cols = [n for n in cols if n in dst_cols]
                c.execute("SELECT %s FROM %s" % (", ".join(_ident(n) for n in cols), _ident(table)))
                dst.executemany(
                    "INSERT INTO %s (%s) VALUES (%s)" % (table, ", ".join(cols), ", ".join("?" * len(cols))),
                    c.fetchall(),
                )
        raw.rollback()
        dst.commit()
        dst.executescript(sqlite_ddl)  # 지웠던 트리거를 다시 만든다 (나머지는 IF NOT EXISTS 라 그대로)
        dst.commit()
    finally:
        dst.close()
        _POOL.put(raw, schema)
    return out_path
