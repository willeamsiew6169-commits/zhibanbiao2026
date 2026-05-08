# db.py

import os
import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2.pool import SimpleConnectionPool

DATABASE_URL = os.environ.get("DATABASE_URL")

def get_db():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL 没有设置")
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)

pool = SimpleConnectionPool(
    1,
    20,
    dsn=DATABASE_URL,
    cursor_factory=RealDictCursor
)

def db_query(sql, params=None, fetchone=False, fetchall=False):
    import psycopg2

    conn = pool.getconn()
    try:
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
        # 🔥 连接坏了 → 重试一次
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
        pool.putconn(conn)
