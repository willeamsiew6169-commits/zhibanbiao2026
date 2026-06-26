# db.py

import os
import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2.pool import SimpleConnectionPool
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL 没有设置，请检查 .env 文件")

pool = SimpleConnectionPool(
    1,
    20,
    dsn=DATABASE_URL,
    cursor_factory=RealDictCursor,
    connect_timeout=10,
    sslmode="require",
    keepalives=1,
    keepalives_idle=30,
    keepalives_interval=10,
    keepalives_count=5
)

def get_db():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)

def get_conn():
    return psycopg2.connect(DATABASE_URL) 

def db_query(sql, params=None, fetchone=False, fetchall=False):
    import psycopg2

    conn = None

    try:
        conn = pool.getconn()

        with conn.cursor() as cur:
            cur.execute(sql, params or ())

            if fetchone:
                result = cur.fetchone()
            elif fetchall:
                result = cur.fetchall()
            else:
                result = None

            conn.commit()
            return result

    except psycopg2.OperationalError:

        # 坏连接不要放回 pool
        try:
            if conn:
                pool.putconn(conn, close=True)
        except:
            pass

        # 重新拿新连接
        conn = pool.getconn()

        with conn.cursor() as cur:
            cur.execute(sql, params or ())

            if fetchone:
                result = cur.fetchone()
            elif fetchall:
                result = cur.fetchall()
            else:
                result = None

            conn.commit()
            return result

    finally:
        try:
            if conn and not conn.closed:
                pool.putconn(conn)
        except:
            pass