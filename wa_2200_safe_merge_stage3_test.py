"""22:00 第三阶段：完全离线的安全补排规划测试。
不导入项目代码、不连接数据库、不写入、不发布、不发送。
运行: python wa_2200_safe_merge_stage3_test.py
"""
import json


def gaps(required, occupied):
    start, end = required
    if not (0 <= start < end <= 1440):
        raise ValueError('报名时间无效')
    cursor, result = start, []
    for a, b in sorted(occupied):
        if not (0 <= a < b <= 1440):
            raise ValueError('已有排班时间无效')
        if b <= cursor or a >= end:
            continue
        if a > cursor:
            result.append((cursor, min(a, end)))
        cursor = max(cursor, min(b, end))
        if cursor >= end:
            break
    if cursor < end:
        result.append((cursor, end))
    return result


def overlap(a, b):
    return max(a[0], b[0]) < min(a[1], b[1])


def fmt(pair):
    def t(m):
        return f'{m // 60:02d}:{m % 60:02d}'
    return t(pair[0]) + '-' + t(pair[1])


def plan(signup, existing, theoretical):
    """只计算 INSERT 候选；理论段必须恰好覆盖缺口，不能修改任何旧记录。"""
    sid = signup['signup_id']
    current = [r for r in existing if r.get('signup_id') == sid and r.get('role') == '值班']
    occupied = [r['interval'] for r in current]
    missing = gaps(signup['interval'], occupied)
    if not missing:
        return {'status': 'already_covered', 'add': [], 'missing': []}
    if any(r.get('locked') for r in current):
        return {'status': 'manual_review_locked_partial', 'add': [], 'missing': list(map(fmt, missing))}
    candidates = [r for r in theoretical if r.get('signup_id') == sid and r.get('role') == '值班']
    additions = []
    for gap in missing:
        for row in candidates:
            segment = row['interval']
            a, b = max(gap[0], segment[0]), min(gap[1], segment[1])
            if a >= b:
                continue
            if row.get('place') not in ('观音堂', '活动中心'):
                return {'status': 'invalid_place', 'add': [], 'missing': list(map(fmt, missing))}
            if any(overlap((a, b), x) for x in occupied):
                return {'status': 'overlap_existing', 'add': [], 'missing': list(map(fmt, missing))}
            additions.append({'signup_id': sid, 'interval': (a, b), 'place': row['place']})
    if gaps(signup['interval'], occupied + [r['interval'] for r in additions]):
        return {'status': 'theory_incomplete', 'add': [], 'missing': list(map(fmt, missing))}
    for i, a in enumerate(additions):
        if any(overlap(a['interval'], b['interval']) for b in additions[i + 1:]):
            return {'status': 'theory_overlapping', 'add': [], 'missing': list(map(fmt, missing))}
    return {'status': 'plan_only', 'add': [{'time': fmt(r['interval']), 'place': r['place']} for r in additions], 'missing': list(map(fmt, missing))}


def row(sid, start, end, place='观音堂', locked=False, role='值班'):
    return {'signup_id': sid, 'interval': (start, end), 'place': place, 'locked': locked, 'role': role}


def run():
    signup = {'signup_id': 100, 'interval': (600, 840)}
    theory = [row(100, 600, 720), row(100, 720, 840, '活动中心')]
    cases = [
        ('完整排班不新增', [row(100, 600, 720), row(100, 720, 840, '活动中心')], theory, 'already_covered', 0),
        ('前半段已有只补后半段', [row(100, 600, 720)], theory, 'plan_only', 1),
        ('后半段已有只补前半段', [row(100, 720, 840, '活动中心')], theory, 'plan_only', 1),
        ('完全空白补两段', [], theory, 'plan_only', 2),
        ('锁定但完整不改变', [row(100, 600, 840, locked=True)], theory, 'already_covered', 0),
        ('锁定且部分缺漏交人工', [row(100, 600, 720, locked=True)], theory, 'manual_review_locked_partial', 0),
        ('理论不完整禁止补排', [], [row(100, 600, 720)], 'theory_incomplete', 0),
        ('理论重叠禁止补排', [], [row(100, 600, 780), row(100, 720, 840)], 'theory_overlapping', 0),
        ('其他岗位原记录不参与值班覆盖', [row(100, 600, 840, role='卫生')], theory, 'plan_only', 2),
        ('其他报名原记录不参与覆盖', [row(101, 600, 840)], theory, 'plan_only', 2),
    ]
    checks = []
    for name, existing, proposed, expected, expected_count in cases:
        result = plan(signup, existing, proposed)
        checks.append({'scenario': name, 'passed': result['status'] == expected and len(result['add']) == expected_count,
                       'status': result['status'], 'would_insert': result['add']})
    return {'stage': 'STAGE3_OFFLINE_PLAN_ONLY', 'all_passed': all(c['passed'] for c in checks),
            'checks': checks, 'database_access': False, 'writes_performed': 0,
            'published': False, 'whatsapp_sent': False, 'safe_to_publish': False,
            'note': '仅模拟值班补排规划；尚未验证实际数据库事务、卫生和佛台完整性。'}


if __name__ == '__main__':
    print(json.dumps(run(), ensure_ascii=False, indent=2))
