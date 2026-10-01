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


IDLE_IN_TRANSACTION_TIMEOUT = 60  # 초. 화면 한 번 그리는 시간보다 충분히 길게.
DDL_LOCK_TIMEOUT = "3s"           # 스키마 점검 DDL 이 잠금을 기다리는 최대 시간


class _Pool:
    """Streamlit 은 버튼을 누를 때마다 스크립트 전체를 다시 실행하고 그때마다 연결을 여러 번 연다.
    Supabase 연결은 한 번 맺는 데 수백 ms 가 걸리므로, 닫힌 연결을 버리지 않고 모아 두었다 다시 쓴다."""

    MAX_IDLE = 4

    def __init__(self):
        self._idle = []
        self._lock = threading.Lock()

    def get(self, fresh=False):
        """쉬고 있던 연결을 돌려준다. 미리 'SELECT 1' 로 살아 있는지 확인하면 그만큼 왕복이 늘어나므로,
        확인하지 않고 쓰다가 끊긴 연결이면 Connection._run 이 새 연결로 한 번 다시 시도한다."""
        while not fresh:
            with self._lock:
                if not self._idle:
                    break
                raw, schema, _ = self._idle.pop()
            if not raw.closed:
                return raw, schema
        url = database_url()
        if not url:
            raise OperationalError("DATABASE_URL 이 설정되지 않았습니다 (.streamlit/secrets.toml 또는 Streamlit Cloud Secrets).")
        try:
            raw = psycopg2.connect(url, connect_timeout=15, application_name="testfarm-erp",
                                   keepalives=1, keepalives_idle=30)
            # Streamlit 은 화면을 그리는 도중 브라우저가 닫히면 스크립트를 멈추는데, 그때 열려 있던 연결은
            # 조회 트랜잭션이 열린 채('idle in transaction') 남아 표 잠금을 쥐고 있게 된다. 그러면 스키마 점검·
            # 리셋·복원 같은 DDL 이 잠금을 기다리다 2분 뒤 시간 초과로 실패한다.
            # 이런 연결은 서버가 일정 시간 뒤 끊도록 한다 (끊긴 연결은 다음에 쓸 때 Connection._run 이 다시 연결).
            raw.autocommit = True
            with raw.cursor() as c:
                c.execute("SET idle_in_transaction_session_timeout = '%ds'" % IDLE_IN_TRANSACTION_TIMEOUT)
            raw.autocommit = False
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


# ========== 조회 결과 캐시 ==========
# Streamlit 은 클릭할 때마다 화면 전체를 다시 그리며 조회를 30번 넘게 보낸다. Streamlit Cloud(미국) ↔
# Supabase(서울) 왕복이 길어 클릭마다 수 초가 걸리므로, SELECT 결과를 농장(스키마)별로 기억해 두고 재사용한다.
# 그 농장에 쓰기(INSERT/UPDATE/DELETE 등)가 일어나면 즉시 그 농장 캐시를 비운다 → 이 앱에서 한 변경은 바로 보인다.
# 다른 곳(로컬 PC 의 앱, Supabase 대시보드 등)에서 바꾼 내용은 CACHE_TTL 이 지나거나 clear_cache() 를 부르면 보인다.
CACHE_TTL = 600  # 초
_CACHE_MAX = 2000
_cache = {}      # (schema, sql, params) -> (만료시각, description, rows)
_cache_ver = {}  # schema -> 쓰기 때마다 1씩 증가 (조회 도중 쓰기가 끼면 그 결과는 저장하지 않는다)
_cache_lock = threading.Lock()
_exists = set()  # cattle 표가 있다고 확인된 스키마
# 이 프로세스에서 이미 표 구조를 점검한 스키마. Streamlit 은 클릭마다 erp_ui.py 를 처음부터 다시 실행하므로
# 이 기록은 앱 쪽이 아니라 (다시 실행되지 않는) 이 모듈에 둬야 한다.
_ensured = set()


def _is_read(sql):
    return sql.lstrip().lstrip("(").lstrip()[:6].upper() == "SELECT"


def _cache_get(key):
    with _cache_lock:
        hit = _cache.get(key)
        if hit and hit[0] > time.time():
            return hit[1], hit[2]
        if hit:
            del _cache[key]
    return None


def _cache_put(key, version, description, rows):
    with _cache_lock:
        if _cache_ver.get(key[0], 0) != version:
            return
        if len(_cache) >= _CACHE_MAX:
            _cache.clear()
        _cache[key] = (time.time() + CACHE_TTL, description, rows)


def _cache_version(schema):
    with _cache_lock:
        return _cache_ver.get(schema, 0)


def invalidate(schema):
    """해당 농장 스키마의 조회 캐시를 비운다."""
    with _cache_lock:
        _cache_ver[schema] = _cache_ver.get(schema, 0) + 1
        for k in [k for k in _cache if k[0] == schema]:
            del _cache[k]


def clear_cache():
    """모든 조회 캐시를 비운다 (다른 곳에서 바꾼 데이터를 바로 불러오고 싶을 때)."""
    with _cache_lock:
        for schema in set(_cache_ver) | {k[0] for k in _cache}:
            _cache_ver[schema] = _cache_ver.get(schema, 0) + 1
        _cache.clear()
        _exists.clear()


# ========== sqlite3 호환 연결 / 커서 ==========
class Cursor:
    arraysize = 1
    lastrowid = None

    def __init__(self, conn):
        self._conn = conn
        self._cur = conn._raw.cursor()
        self.description = None
        self.rowcount = -1
        self._rows = None  # 결과는 한 번에 받아 두고 여기서 꺼내 준다 (캐시에 넣기 위해)
        self._pos = 0

    def execute(self, sql, params=None):
        q = _translate(sql)
        if params is not None:
            if isinstance(params, dict):
                raise ProgrammingError("이름 붙은 파라미터(:name)는 지원하지 않습니다.")
            params = tuple(params)
            q = _bind(q)
        conn = self._conn
        self._rows, self._pos = None, 0
        is_read = _is_read(q)
        # 이 연결에 아직 커밋하지 않은 쓰기가 있으면, 그 변경이 보여야 하므로 캐시를 쓰지 않는다.
        key = None
        if is_read and not conn._dirty:
            key = (conn._schema, q, params)
            try:
                hit = _cache_get(key)
            except TypeError:  # 해시할 수 없는 파라미터
                key, hit = None, None
            if hit:
                self.description, rows = hit
                self._rows, self.rowcount = rows, len(rows)
                return self
            version = _cache_version(conn._schema)
        conn._run(self, q, params)
        self.description = self._cur.description
        self.rowcount = self._cur.rowcount
        if self.description is not None:
            self._rows = self._cur.fetchall()
            if key is not None:
                _cache_put(key, version, self.description, self._rows)
        if not is_read:
            conn._mark_dirty()
        return self

    def executemany(self, sql, seq_of_parameters):
        total = 0
        for params in seq_of_parameters:
            self.execute(sql, params)
            total += max(self.rowcount, 0)
        self.rowcount = total
        return self

    def executescript(self, script):
        self._rows, self._pos = None, 0
        self._conn._run(self, _translate(script), None)
        self._conn._mark_dirty()
        return self

    def fetchone(self):
        if not self._rows or self._pos >= len(self._rows):
            return None
        self._pos += 1
        return self._rows[self._pos - 1]

    def fetchall(self):
        if not self._rows:
            return []
        rest = self._rows[self._pos:]
        self._pos = len(self._rows)
        return list(rest)

    def fetchmany(self, size=None):
        if not self._rows:
            return []
        size = size or self.arraysize
        part = self._rows[self._pos:self._pos + size]
        self._pos += len(part)
        return list(part)

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
        self._dirty = False      # 현재 트랜잭션 안에 쓰기가 있었는지
        self._closed = False

    def _mark_dirty(self):
        self._dirty = True
        invalidate(self._schema)

    def _check(self):
        if self._closed:
            raise ProgrammingError("Cannot operate on a closed database.")

    def _run(self, cursor, sql, params, retry=True):
        self._check()
        # 명령마다 SAVEPOINT 를 새로 건다(같은 왕복에 실어 보내므로 추가 지연 없음).
        # 실패하면 그 명령만 되돌려, 같은 트랜잭션의 이전 작업과 이후 명령이 살아남는다 (SQLite 와 같은 동작).
        prefix = "RELEASE SAVEPOINT sqlite_stmt; SAVEPOINT sqlite_stmt; " if self._savepoint else "SAVEPOINT sqlite_stmt; "
        try:
            cursor._cur.execute(prefix + sql, params)
            self._savepoint = True
        except (psycopg2.OperationalError, psycopg2.InterfaceError) as e:
            # 오래 쉬던 연결이 서버 쪽에서 끊긴 경우: 이 트랜잭션에 쓰기가 없었다면 새 연결로 한 번 다시 시도한다.
            if retry and self._raw.closed and not self._dirty:
                self._reconnect()
                cursor._cur = self._raw.cursor()
                return self._run(cursor, sql, params, retry=False)
            self._recover()
            raise _wrap_error(e) from None
        except psycopg2.Error as e:
            self._recover()
            raise _wrap_error(e) from None

    def _reconnect(self):
        _close_quietly(self._raw)
        raw, _ = _POOL.get(fresh=True)
        try:
            _set_schema(raw, self._schema)
        except psycopg2.Error as e:
            _close_quietly(raw)
            raise _wrap_error(e) from None
        self._raw, self._savepoint = raw, False

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
            if self._dirty:
                # 커밋 전에 다른 연결이 옛 값을 읽어 캐시에 넣었을 수 있으니 커밋 시점에 한 번 더 비운다.
                invalidate(self._schema)
            self._savepoint = self._dirty = False

    def rollback(self):
        if self._closed or self._raw.closed:
            return
        try:
            self._raw.rollback()
        except psycopg2.Error as e:
            raise _wrap_error(e) from None
        finally:
            self._savepoint = self._dirty = False

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
        except psycopg2.Error:
            # 쉬던 연결이 끊겨 있었을 수 있으니 새 연결로 한 번 더 시도한다.
            _close_quietly(raw)
            raw, _ = _POOL.get(fresh=True)
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
    """농장 스키마에 cattle 표가 있으면 '이 농장 DB가 있다'로 본다 (os.path.exists 대용).
    매 화면마다 묻지 않도록, 있다고 확인된 스키마는 기억해 둔다 (리셋·복원·clear_cache 때 다시 확인)."""
    schema = schema_for(db_file)
    if schema in _exists:
        return True
    raw = _raw_connection()
    try:
        with raw.cursor() as c:
            c.execute("SELECT to_regclass(%s)", (_ident(schema) + ".cattle",))
            found = c.fetchone()[0] is not None
        if found:
            _exists.add(schema)
        return found
    except psycopg2.Error:
        _close_quietly(raw)
        return False
    finally:
        _POOL.put(raw, None)


def _run_ddl(c, schema, ddl, drop_first):
    if drop_first:
        c.execute("DROP SCHEMA IF EXISTS %s CASCADE" % _ident(schema))
    c.execute("CREATE SCHEMA IF NOT EXISTS %s" % _ident(schema))
    c.execute("SET LOCAL search_path TO %s" % _ident(schema))
    c.execute(ddl)


def _schema_complete(c, schema, required_tables, required_columns, required_triggers):
    """필요한 표·컬럼·트리거가 이미 다 있는지 카탈로그만 읽어 확인한다 (표 잠금을 잡지 않는다)."""
    c.execute("SELECT table_name FROM information_schema.tables WHERE table_schema = %s", (schema,))
    tables = {r[0] for r in c.fetchall()}
    if not set(required_tables) <= tables:
        return False
    c.execute("SELECT table_name, column_name FROM information_schema.columns WHERE table_schema = %s", (schema,))
    columns = {(r[0], r[1]) for r in c.fetchall()}
    if not set(required_columns) <= columns:
        return False
    c.execute("SELECT t.tgname FROM pg_trigger t JOIN pg_class r ON r.oid = t.tgrelid "
              "JOIN pg_namespace n ON n.oid = r.relnamespace WHERE n.nspname = %s AND NOT t.tgisinternal", (schema,))
    return set(required_triggers) <= {r[0] for r in c.fetchall()}


def ensure_schema(db_file, ddl, extra_sql=(), required_tables=(), required_columns=(), required_triggers=()):
    """스키마와 표를 만든다(이미 있으면 그대로). extra_sql 은 나중에 추가된 컬럼 등 보완용 문장.
    프로세스당 스키마별로 한 번만 실제로 실행한다.

    CREATE TRIGGER / ALTER TABLE 은 '이미 있어도' 표에 배타 잠금을 잡으므로, 다른 연결이 그 표를 읽는 중이면
    기다리게 된다. 그래서 required_* 로 넘긴 것이 이미 다 있으면 DDL 을 아예 실행하지 않고,
    실행할 때도 잠금은 DDL_LOCK_TIMEOUT 까지만 기다린다. 기다리다 실패하면(표는 이미 있으니) 이번에는
    건너뛰고 다음 실행 때 다시 점검한다 — 화면이 멈추거나 오류로 끝나지 않게."""
    schema = schema_for(db_file)
    if schema in _ensured:
        return
    raw = _raw_connection()
    try:
        with raw.cursor() as c:
            if required_tables and _schema_complete(c, schema, required_tables, required_columns, required_triggers):
                raw.rollback()
                _exists.add(schema)
                _ensured.add(schema)
                return
            c.execute("SET LOCAL lock_timeout = '%s'" % DDL_LOCK_TIMEOUT)
            _run_ddl(c, schema, ddl, drop_first=False)
            for stmt in extra_sql:
                c.execute(stmt)
        raw.commit()
        invalidate(schema)
        _exists.add(schema)
        _ensured.add(schema)
    except psycopg2.Error as e:
        raw.rollback()
        if getattr(e, "pgcode", None) == "55P03" and database_exists(db_file):  # lock_not_available
            return
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
        invalidate(schema)
        _exists.add(schema)
        _ensured.add(schema)
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
        invalidate(schema)
        _exists.add(schema)
        _ensured.add(schema)
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
