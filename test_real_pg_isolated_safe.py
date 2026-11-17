"""真实 PostgreSQL 事务测试；仅连接独立 TEST_DATABASE_URL；不访问正式表。"""
import ast
import os
import sys
import types
from pathlib import Path
from urllib.parse import urlsplit

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / 'schedule' / 'builders' / 'schedule_builder_V2_2200_merged_REVIEW.py'

def fail(msg):
    raise SystemExit('STOP | ' + msg)

# 安全保护：正式库仅从本机 .env 读取，不连接正式库。
from dotenv import dotenv_values

EXPECTED_TEST_PROJECT_REF = 'kztjycyxawrqrfvlgddh'
test_url = os.getenv('TEST_DATABASE_URL', '').strip()
prod_url = (dotenv_values(ROOT / '.env').get('DATABASE_URL') or '').strip()
if not test_url:
    fail('未设置 TEST_DATABASE_URL。')
if not SOURCE.exists():
    fail(f'找不到整合版：{SOURCE}')
if not prod_url:
    fail('本机 .env 未找到正式 DATABASE_URL；禁止运行。')
if os.getenv('CONFIRM_ISOLATED_TEST_DB') != 'YES':
    fail('未设置 CONFIRM_ISOLATED_TEST_DB=YES。')

def project_ref(url):
    try:
        p = urlsplit(url.replace('postgresql+psycopg2://', 'postgresql://', 1))
        if p.scheme not in ('postgresql', 'postgres') or not p.hostname or not p.username:
            return None
        # Supabase Session/Transaction pooler: postgres.<project_ref>
        if p.username.startswith('postgres.'):
            return p.username.split('.', 1)[1]
        # Supabase direct: db.<project_ref>.supabase.co
        parts = p.hostname.split('.')
        if len(parts) >= 4 and parts[0] == 'db' and parts[2:] == ['supabase', 'co']:
            return parts[1]
        return None
    except ValueError:
        return None

if project_ref(test_url) != EXPECTED_TEST_PROJECT_REF:
    fail('测试连接地址的 Supabase 项目编号不符合预期；禁止运行。')
prod_ref = project_ref(prod_url)
if not prod_ref:
    fail('无法从正式 DATABASE_URL 识别 Supabase 项目编号；禁止运行。')
if prod_ref == EXPECTED_TEST_PROJECT_REF:
    fail('正式库与测试库项目编号相同；禁止运行。')
print('PASS | 隔离检查：测试项目编号符合预期，且与正式项目不同（未连接正式库）')

import psycopg2
from psycopg2.extras import RealDictCursor

# 仅提取实际整合版中的快照函数，不 import 整个排班系统。
source = SOURCE.read_text(encoding='utf-8')
tree = ast.parse(source)
fn = next((n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_capture_2200_baseline_in_transaction'), None)
if fn is None:
    fail('整合版缺少快照辅助函数')
code = compile(ast.Module(body=[fn], type_ignores=[]), str(SOURCE), 'exec')

# 提供固定测试数据；快照 SQL 和锁逻辑均来自真实整合版。
services = types.ModuleType('schedule.services.whatsapp_morning_summary')
services._assignments = lambda cur, target: [
    {'volunteer_id': 'TEST-801', 'role': '值班'},
    {'volunteer_id': 'TEST-802', 'role': '卫生'},
]
services._key = lambda r: (r['volunteer_id'], r['role'])
for name in ('schedule', 'schedule.services'):
    if name not in sys.modules:
        m = types.ModuleType(name)
        m.__path__ = []
        sys.modules[name] = m
sys.modules['schedule.services.whatsapp_morning_summary'] = services
ns = {}
exec(code, ns)
capture = ns['_capture_2200_baseline_in_transaction']

conn = psycopg2.connect(test_url, connect_timeout=10, sslmode='require')
conn.autocommit = False
try:
    with conn.cursor() as cur:
        cur.execute('CREATE TEMP TABLE schedule_day_flags (flag_date date PRIMARY KEY, is_published boolean NOT NULL DEFAULT false) ON COMMIT PRESERVE ROWS')
        cur.execute('CREATE TEMP TABLE whatsapp_morning_snapshots (target_date date PRIMARY KEY, baseline jsonb NOT NULL) ON COMMIT PRESERVE ROWS')
    conn.commit()
    date = '2099-12-31'

    # 失败注入：发布状态 + 快照写入后强制失败，要求两者全部回滚。
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute('INSERT INTO schedule_day_flags(flag_date, is_published) VALUES (%s, true)', (date,))
            capture(cur, date)
            raise RuntimeError('INJECTED_FAILURE_AFTER_SNAPSHOT')
    except RuntimeError as exc:
        assert str(exc) == 'INJECTED_FAILURE_AFTER_SNAPSHOT'
        conn.rollback()
    with conn.cursor() as cur:
        cur.execute('SELECT count(*) FROM schedule_day_flags')
        assert cur.fetchone()[0] == 0, '发布状态未回滚'
        cur.execute('SELECT count(*) FROM whatsapp_morning_snapshots')
        assert cur.fetchone()[0] == 0, '快照未回滚'
    print('PASS | 真实 PostgreSQL：发布状态与快照一起回滚')

    with conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute('INSERT INTO schedule_day_flags(flag_date, is_published) VALUES (%s, true)', (date,))
        capture(cur, date)
    conn.commit()
    with conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute('SELECT baseline FROM whatsapp_morning_snapshots WHERE target_date=%s', (date,))
        original = cur.fetchone()['baseline']
        assert ['TEST-801', '值班'] in original and ['TEST-802', '卫生'] in original
    print('PASS | 真实 PostgreSQL：成功提交快照，包含值班及卫生')

    with conn.cursor(cursor_factory=RealDictCursor) as cur:
        capture(cur, date)
    conn.commit()
    with conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute('SELECT baseline FROM whatsapp_morning_snapshots WHERE target_date=%s', (date,))
        assert cur.fetchone()['baseline'] == original
        cur.execute('SELECT count(*) FROM whatsapp_morning_snapshots')
        assert cur.fetchone()['count'] == 1
    print('PASS | 真实 PostgreSQL：重复调用不覆盖、不重复')
    print('RESULT | 3/3 PASS；仅验证真实快照辅助函数与临时表事务，不代表整条 22:00 流程或 WhatsApp 已实测。')
finally:
    conn.rollback()
    conn.close()  # 临时表随连接消失，测试库不留持久数据
