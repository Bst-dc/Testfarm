import psycopg2
import psycopg2.extras
import streamlit as st

def _convert_sql(sql):
    sql = sql.replace('INTEGER PRIMARY KEY AUTOINCREMENT', 'SERIAL PRIMARY KEY')
    sql = sql.replace('AUTOINCREMENT', 'SERIAL')
    # Pandas checking table existence
    import re
    if 'sqlite_master' in sql:
        sql = re.sub(
            r"SELECT\s+name\s+FROM\s+sqlite_master\s+WHERE\s+type\s+IN\s+\('table',\s*'view'\)\s+AND\s+name=",
            "SELECT tablename as name FROM pg_tables WHERE schemaname = current_schema() AND tablename=",
            sql, flags=re.IGNORECASE|re.MULTILINE
        )
        sql = re.sub(
            r"SELECT\s+name\s+FROM\s+sqlite_master\s+WHERE\s+type='table'\s+AND\s+name=",
            "SELECT tablename as name FROM pg_tables WHERE schemaname = current_schema() AND tablename=",
            sql, flags=re.IGNORECASE|re.MULTILINE
        )
        
    if 'PRAGMA' in sql:
        # Handle PRAGMA table_info(table_name)
        match = re.search(r"PRAGMA\s+table_info\((.+?)\)", sql, re.IGNORECASE)
        if match:
            table_name = match.group(1).strip("'\"")
            sql = f"SELECT 0, column_name FROM information_schema.columns WHERE table_name = '{table_name}' AND table_schema = current_schema() ORDER BY ordinal_position"
        else:
            # Replace other PRAGMAs with a no-op
            sql = re.sub(r"PRAGMA\s+.*", "SELECT 1", sql, flags=re.IGNORECASE)
            
    return sql

def _replace_placeholders(sql):
    return sql.replace('?', '%s')

class PostgresCursorWrapper:
    def __init__(self, cursor):
        self.cursor = cursor
        self.description = None
        self.rowcount = -1
        
    def execute(self, sql, params=None):
        sql = _convert_sql(sql)
        sql = _replace_placeholders(sql)
        try:
            if params is None:
                self.cursor.execute(sql)
            else:
                self.cursor.execute(sql, params)
            self.description = self.cursor.description
            self.rowcount = self.cursor.rowcount
        except psycopg2.Error as e:
            if e.pgcode and e.pgcode.startswith('23'):
                raise IntegrityError(str(e))
            raise OperationalError(str(e))
        return self
        
    def executemany(self, sql, seq_of_parameters):
        sql = _convert_sql(sql)
        sql = _replace_placeholders(sql)
        try:
            self.cursor.executemany(sql, seq_of_parameters)
            self.rowcount = self.cursor.rowcount
        except psycopg2.Error as e:
            if e.pgcode and e.pgcode.startswith('23'):
                raise IntegrityError(str(e))
            raise OperationalError(str(e))
        return self
        
    def executescript(self, sql_script):
        sql_script = _convert_sql(sql_script)
        try:
            self.cursor.execute(sql_script)
        except psycopg2.Error as e:
            raise OperationalError(str(e))
        return self

    def fetchone(self):
        return self.cursor.fetchone()
        
    def fetchall(self):
        return self.cursor.fetchall()
        
    def fetchmany(self, size):
        return self.cursor.fetchmany(size)
        
    @property
    def lastrowid(self):
        # lastrowid is not natively supported without RETURNING in Postgres
        return None
        
    def close(self):
        self.cursor.close()

class PostgresConnectionWrapper:
    def __init__(self, conn):
        self.conn = conn
        
    def cursor(self):
        return PostgresCursorWrapper(self.conn.cursor(cursor_factory=psycopg2.extras.DictCursor))
        
    def execute(self, sql, params=None):
        cur = self.cursor()
        return cur.execute(sql, params)
        
    def executescript(self, sql_script):
        cur = self.cursor()
        return cur.executescript(sql_script)
        
    def commit(self):
        self.conn.commit()
        
    def rollback(self):
        self.conn.rollback()
        
    def close(self):
        self.conn.close()
        
    def __enter__(self):
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is None:
            self.commit()
        else:
            self.rollback()
        self.close()

# Fake sqlite3 exceptions
OperationalError = psycopg2.OperationalError
IntegrityError = psycopg2.IntegrityError
Error = psycopg2.Error

def connect(farm_name, timeout=30, check_same_thread=False):
    url = st.secrets["DATABASE_URL"]
    conn = psycopg2.connect(url)
    
    # Extract schema name from farm_name (e.g., 'erp_sunsan.db' -> 'sunsan')
    import os
    schema_name = os.path.splitext(os.path.basename(farm_name))[0].replace("erp_", "").replace(" ", "_")
    if not schema_name or schema_name == "test_db":
        schema_name = "public"
        
    with conn.cursor() as cur:
        cur.execute(f'CREATE SCHEMA IF NOT EXISTS "{schema_name}"')
        cur.execute(f'SET search_path TO "{schema_name}"')
    
    conn.commit()
    return PostgresConnectionWrapper(conn)

