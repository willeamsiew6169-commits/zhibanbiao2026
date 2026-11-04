# db.py

import os
import time
import logging
import psycopg2
from contextlib import contextmanager
from psycopg2.extras import RealDictCursor
from psycopg2.pool import ThreadedConnectionPool
# ThreadedConnectionPool 可安全用于多线程请求；保持原有 pool 接口。
from dotenv import load_dotenv


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL 没有设置，请检查 .env 文件")


# ============================================================
# PostgreSQL connection pool
# ============================================================
pool = ThreadedConnectionPool(
    1,
    20,
    dsn=DATABASE_URL,
    cursor_factory=RealDictCursor,
    connect_timeout=10,
    sslmode="require",
    keepalives=1,
    keepalives_idle=30,
    keepalives_interval=10,
    keepalives_count=5,
)


def get_db():
    """
    保留旧接口：
    某些旧模块会直接调用 get_db()，所以这里不改变原本行为。
    每次调用建立一个新的 PostgreSQL connection。
    """
    return psycopg2.connect(
        DATABASE_URL,
        cursor_factory=RealDictCursor,
        connect_timeout=10,
        sslmode="require",
        keepalives=1,
        keepalives_idle=30,
        keepalives_interval=10,
        keepalives_count=5,
    )


def _discard_connection(conn):
    """把失效连接从 connection pool 中彻底丢弃。"""
    if conn is None:
        return

    try:
        pool.putconn(conn, close=True)
    except Exception:
        try:
            conn.close()
        except Exception:
            pass


def _get_healthy_connection():
    """
    从 pool 取得一个可用连接。

    重点：
    psycopg2 的 conn.closed == 0 并不保证服务器端连接仍然存在。
    因此从 pool 拿出旧连接后，用 SELECT 1 实际测试一次。

    如果测试失败：
    - 丢弃坏连接
    - 再取一个连接
    - 最多尝试 2 次
    """
    last_error = None

    for _ in range(2):
        conn = None

        try:
            conn = pool.getconn()

            if conn is None:
                raise psycopg2.OperationalError(
                    "无法从 PostgreSQL connection pool 取得连接"
                )

            if conn.closed:
                _discard_connection(conn)
                conn = None
                continue

            # 如果上一次使用后 transaction 状态异常，
            # 先 rollback，避免 SELECT 1 因 aborted transaction 失败。
            try:
                conn.rollback()
            except (psycopg2.InterfaceError, psycopg2.OperationalError):
                _discard_connection(conn)
                conn = None
                continue

            # 真正向 PostgreSQL 发出请求，确认服务器端连接仍然活着。
            with conn.cursor() as cur:
                cur.execute("SELECT 1")
                cur.fetchone()

            # 清掉健康检查产生的 transaction。
            conn.rollback()
            return conn

        except (psycopg2.OperationalError, psycopg2.InterfaceError) as exc:
            last_error = exc

            if conn is not None:
                _discard_connection(conn)
                conn = None

    if last_error is not None:
        raise last_error

    raise psycopg2.OperationalError("无法取得可用的 PostgreSQL connection")


@contextmanager
def get_conn():
    """
    推荐的新式数据库连接方式：

        with get_conn() as conn:
            with conn.cursor() as cur:
                cur.execute(...)

    功能：
    - 从 connection pool 取得连接
    - 使用前检查连接是否仍然有效
    - 成功后自动 commit
    - 出错自动 rollback
    - 已断开的连接不会重新放回 pool
    """
    conn = None
    broken = False

    try:
        conn = _get_healthy_connection()
        yield conn

        if conn and not conn.closed:
            conn.commit()
        else:
            broken = True

    except (psycopg2.OperationalError, psycopg2.InterfaceError):
        broken = True

        try:
            if conn and not conn.closed:
                conn.rollback()
        except (psycopg2.InterfaceError, psycopg2.OperationalError):
            pass

        raise

    except Exception:
        try:
            if conn and not conn.closed:
                conn.rollback()
            else:
                broken = True
        except (psycopg2.InterfaceError, psycopg2.OperationalError):
            broken = True

        raise

    finally:
        if conn is not None:
            try:
                pool.putconn(
                    conn,
                    close=(broken or bool(conn.closed)),
                )
            except Exception:
                try:
                    conn.close()
                except Exception:
                    pass


def _execute_query(conn, sql, params=None, fetchone=False, fetchall=False):
    """db_query() 内部共用的 SQL 执行函数。"""
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


def db_query(sql, params=None, fetchone=False, fetchall=False):
    """执行单次 SQL，并安全管理连接及事务。

    连接借用阶段由 _get_healthy_connection() 负责检测断线并重试。
    SQL 开始执行后不自动重试：网络可能在数据库提交成功、
    但客户端尚未收到确认时断开；盲目重试 INSERT/UPDATE
    可能导致重复签到、重复付款等数据问题。

    调用方如需对幂等的只读查询重试，应在业务层明确控制。
    """
    conn = None
    broken = False
    started = time.perf_counter()
    acquired = started
    started = time.perf_counter()
    acquired = started
    try:
        conn = _get_healthy_connection()
        acquired = time.perf_counter()
        return _execute_query(
            conn,
            sql,
            params=params,
            fetchone=fetchone,
            fetchall=fetchall,
        )
    except (psycopg2.OperationalError, psycopg2.InterfaceError):
        broken = True
        if conn is not None and not conn.closed:
            try:
                conn.rollback()
            except (psycopg2.OperationalError, psycopg2.InterfaceError):
                pass
        raise
    except Exception:
        if conn is not None and not conn.closed:
            try:
                conn.rollback()
            except (psycopg2.OperationalError, psycopg2.InterfaceError):
                broken = True
        else:
            broken = True
        raise
    finally:
        if os.getenv("ATTENDANCE_PERF_LOG", "0") == "1":
            logging.warning("[DB-PERF] op=%s checkout_ms=%.0f query_commit_ms=%.0f total_ms=%.0f",
                str(sql).strip().split(None, 1)[0].upper() if str(sql).strip() else "?",
                (acquired-started)*1000, (time.perf_counter()-acquired)*1000,
                (time.perf_counter()-started)*1000)
        if conn is not None:
            try:
                pool.putconn(conn, close=(broken or bool(conn.closed)))
            except Exception:
                try:
                    conn.close()
                except Exception:
                    pass
