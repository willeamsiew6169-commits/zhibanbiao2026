"""阶段一：通知队列。绝不直接操作 WhatsApp。"""
from db import get_conn
from psycopg2.extras import RealDictCursor


def enqueue_whatsapp_test(message, requested_by):
    """测试消息：仅入队；将来发送端必须校验 test_mode 和测试群组。"""
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO whatsapp_notification_jobs
                (message_type, target_date, destination, body, status, created_by)
                VALUES ('test', CURRENT_DATE, 'TEST_GROUP_ONLY', %s, 'pending', %s)
                RETURNING id
            """, (message, requested_by))
            job_id = cur.fetchone()["id"]
        conn.commit()
    return job_id


def recent_whatsapp_jobs(limit=20):
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT id, message_type, status, created_at
                FROM whatsapp_notification_jobs
                ORDER BY id DESC LIMIT %s
            """, (min(max(int(limit), 1), 100),))
            return cur.fetchall()
