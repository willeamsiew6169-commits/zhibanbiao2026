"""Stage 5B: concurrent PostgreSQL row-lock test, lab schema ONLY.
Run A and B in separate terminals. Requires psycopg2 and DATABASE_URL.
"""
import argparse
import os
import sys
import time
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()
import psycopg2

TEST_DATE = '2026-11-08'
TEST_SIGNUP = 900003


def log(msg):
    print(f'[{datetime.now():%H:%M:%S}] {msg}', flush=True)


def connect():
    url = os.environ.get('WA_2200_LAB_DATABASE_URL') or os.environ.get('DATABASE_URL')
    if not url:
        sys.exit('缺少 DATABASE_URL。请在本地终端设置数据库连接变量（不要发密码给我）。')
    return psycopg2.connect(url, connect_timeout=12, sslmode='require', application_name='wa_2200_stage5b_lab')


def worker(which, hold):
    conn = connect()
    try:
        conn.autocommit = False
        with conn.cursor() as cur:
            cur.execute('SET LOCAL lock_timeout = %s', ('25000ms',))
            cur.execute('SET LOCAL statement_timeout = %s', ('35000ms',))
            log(f'{which}: 准备取得测试日期 {TEST_DATE} 的行锁')
            t0 = time.monotonic()
            cur.execute('SELECT target_date FROM wa_2200_lab.day_locks WHERE target_date = %s FOR UPDATE', (TEST_DATE,))
            if cur.fetchone() is None:
                raise RuntimeError('测试锁定行不存在；先运行前一步的 SQL。')
            waited = time.monotonic() - t0
            log(f'{which}: 已取得行锁；等待 {waited:.2f} 秒')
            cur.execute('SELECT COUNT(*) FROM wa_2200_lab.assignments WHERE target_date=%s AND signup_id=%s', (TEST_DATE, TEST_SIGNUP))
            count = cur.fetchone()[0]
            if count == 0:
                if hold:
                    log(f'{which}: 持锁 {hold} 秒。现在立即启动终端 B！')
                    time.sleep(hold)
                cur.execute('''INSERT INTO wa_2200_lab.assignments
                    (target_date, signup_id, volunteer_id, role, start_time, end_time, assigned_place)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (target_date, signup_id, role, start_time, end_time) DO NOTHING''',
                    (TEST_DATE, TEST_SIGNUP, 'TEST-003', '值班', '10:00', '12:00', '观音堂'))
                log(f'{which}: 尝试新增；实际新增 {cur.rowcount} 笔')
            else:
                log(f'{which}: 发现已有 {count} 笔，跳过新增')
        conn.commit()
        log(f'{which}: COMMIT 完成')
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def check():
    conn = connect()
    try:
        with conn.cursor() as cur:
            cur.execute('SELECT COUNT(*) FROM wa_2200_lab.assignments WHERE target_date=%s AND signup_id=%s', (TEST_DATE, TEST_SIGNUP))
            count = cur.fetchone()[0]
            log(f'最终测试记录数: {count}（预期 1）')
            if count != 1:
                sys.exit(1)
    finally:
        conn.close()


def main():
    p = argparse.ArgumentParser()
    p.add_argument('mode', choices=['A', 'B', 'check'])
    args = p.parse_args()
    if args.mode == 'check':
        check()
    else:
        worker(args.mode, 12 if args.mode == 'A' else 0)

if __name__ == '__main__':
    main()
