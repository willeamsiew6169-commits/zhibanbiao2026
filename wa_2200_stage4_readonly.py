"""阶段四：读取真实 Supabase 排班，生成只读补排建议。
放在项目根目录，与 wa_2200_safe_merge_stage3_test.py 同级。
用法: python wa_2200_stage4_readonly.py 2026-11-07
不会调用写入/发布/WhatsApp 函数。理论排班依赖 fixed_schedule.xlsx。
"""
import json
import sys
from datetime import date, timedelta
from zoneinfo import ZoneInfo
from datetime import datetime
from psycopg2.extras import RealDictCursor
from db import get_conn
from wa_2200_safe_merge_stage3_test import plan


def minutes(value):
    if value is None:
        return None
    raw = str(value).strip().lower().replace(' ', '')
    if raw.endswith(('am', 'pm')) and ':' not in raw:
        raw = raw[:-2] + ':00' + raw[-2:]
    for fmt in ('%H:%M:%S', '%H:%M', '%I:%M%p'):
        try:
            t = datetime.strptime(raw, fmt)
            return t.hour * 60 + t.minute
        except ValueError:
            pass
    return None


def interval(row):
    a, b = minutes(row.get('start_time')), minutes(row.get('end_time'))
    return (a, b) if a is not None and b is not None and a < b else None


def run(target):
    date.fromisoformat(target)
    # 只读 SQL 查询；未执行 INSERT/UPDATE/DELETE。
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""SELECT id, name, volunteer_id, role, start_time, end_time
                FROM volunteer_schedule_signups WHERE signup_date=%s AND role='值班'
                AND COALESCE(status,'pending') <> 'cancelled' ORDER BY id""", (target,))
            signups = [dict(x) for x in cur.fetchall()]
            cur.execute("""SELECT id, signup_id, volunteer_id, name, role, shift_label,
                assigned_place, start_time, end_time, COALESCE(locked_by_admin,false) AS locked_by_admin
                FROM volunteer_schedule_assignments WHERE assignment_date=%s
                AND COALESCE(status,'assigned') <> 'cancelled' ORDER BY id""", (target,))
            existing = [dict(x) for x in cur.fetchall()]
            cur.execute("""SELECT COALESCE(is_published,false) AS is_published
                FROM schedule_day_flags WHERE flag_date=%s""", (target,))
            flag = cur.fetchone()

    from schedule.builders.schedule_builder import calculate_schedule_for_date
    _, _, theory_rows, _, _ = calculate_schedule_for_date(target)
    proposed = []
    warnings = []
    for r in theory_rows:
        if r.get('role') != '值班' or r.get('signup_id') is None:
            continue
        iv = interval(r)
        if iv is None:
            warnings.append('理论排班时间无效: signup_id=' + str(r.get('signup_id')))
            continue
        proposed.append({'signup_id': int(r['signup_id']), 'role': '值班',
                         'interval': iv, 'place': r.get('assigned_place')})
    current = []
    for r in existing:
        if r.get('signup_id') is None or r.get('role') != '值班':
            continue
        iv = interval(r)
        if iv is None:
            warnings.append('已有值班时间无效: assignment_id=' + str(r['id']))
            continue
        current.append({'signup_id': int(r['signup_id']), 'role': '值班',
                        'interval': iv, 'place': r.get('assigned_place'),
                        'locked': bool(r['locked_by_admin'])})
    checks = []
    for s in signups:
        iv = interval(s)
        if iv is None:
            result = {'status': 'invalid_signup_time', 'add': [], 'missing': []}
        else:
            result = plan({'signup_id': int(s['id']), 'interval': iv}, current, proposed)
        checks.append({'signup_id': s['id'], 'name': s['name'], 'status': result['status'],
                       'missing_intervals': result['missing'], 'would_insert': result['add']})
    return {'stage': 'STAGE4_REAL_DB_READ_ONLY', 'target_date': target,
            'published': bool(flag and flag['is_published']),
            'signup_count': len(signups), 'assignment_count': len(existing),
            'locked_assignment_count': sum(bool(x['locked_by_admin']) for x in existing),
            'checks': checks, 'warnings': warnings,
            'writes_performed_by_this_script': 0, 'published_by_this_script': False,
            'whatsapp_sent_by_this_script': False, 'safe_to_publish': False,
            'note': '仅生成建议；未验证卫生、佛台、并发事务或真实写入安全性。'}


if __name__ == '__main__':
    try:
        target = sys.argv[1] if len(sys.argv) > 1 else (
            datetime.now(ZoneInfo('Asia/Kuala_Lumpur')).date() + timedelta(days=1)).isoformat()
        print(json.dumps(run(target), ensure_ascii=False, indent=2, default=str))
    except Exception as exc:
        print(json.dumps({'stage': 'STAGE4_REAL_DB_READ_ONLY', 'error': str(exc),
                          'safe_to_publish': False}, ensure_ascii=False, indent=2))
        sys.exit(1)
