# -*- coding: utf-8 -*-
"""白话佛法历史记录一次性导入：默认只读预览；输入 yes 才写入。

放在项目根目录，连同 reading_history_payload.json 和原始 reading.xlsx。
不修改 reading_web.py、现有报表代码或 Supabase 表结构。
"""
import hashlib
import json
from collections import Counter
from datetime import date, datetime
from pathlib import Path

from db import get_conn

BASE = Path(__file__).resolve().parent
PAYLOAD = BASE / 'reading_history_payload.json'
EXCEL = BASE / 'reading.xlsx'
FIELDS = ('date', 'name', 'identity', 'topic', 'session', 'time')


def canonical(value):
    if value is None:
        return ''
    if isinstance(value, (date, datetime)):
        return value.strftime('%Y-%m-%d')
    return str(value).strip()


def key(row):
    return tuple(canonical(row.get(f)) for f in FIELDS)


def load_payload():
    data = json.loads(PAYLOAD.read_text(encoding='utf-8'))
    if not EXCEL.exists():
        raise RuntimeError('找不到 reading.xlsx；请将原始 Excel 放在同一目录。')
    actual_sha = hashlib.sha256(EXCEL.read_bytes()).hexdigest()
    if actual_sha != data['source_sha256']:
        raise RuntimeError('reading.xlsx 与已核对的原始文件不同。为避免误导入，已停止。')
    rows = data['rows']
    if len(rows) != data['row_count']:
        raise RuntimeError('导入数据行数与核对结果不一致。')
    for i, r in enumerate(rows, 2):
        if not isinstance(r, dict) or any(f not in r for f in FIELDS):
            raise RuntimeError(f'第 {i} 行字段不完整。')
        if not r['date'] or not r['name'] or not r['topic']:
            raise RuntimeError(f'第 {i} 行缺少日期、姓名或主题。')
        try:
            date.fromisoformat(r['date'])
        except ValueError as exc:
            raise RuntimeError(f'第 {i} 行日期格式错误。') from exc
    return rows


def plan(rows, existing):
    # 多重集合核对：同一人同一天同主题的原始重复行也予以保留；
    # 已有 N 笔则只补足到 Excel 中的笔数，避免重复运行再次插入。
    current = Counter(key(r) for r in existing)
    pending = []
    for r in rows:
        k = key(r)
        if current[k]:
            current[k] -= 1
        else:
            pending.append(r)
    return pending


def main():
    rows = load_payload()
    print('=== 白话佛法历史数据导入核对 ===')
    print('Excel 原始记录：', len(rows))
    for month, count in sorted(Counter(r['date'][:7] for r in rows).items()):
        print(f'  {month}: {count} 笔')
    duplicates = sum(n - 1 for n in Counter(key(r) for r in rows).values() if n > 1)
    print('原始资料内完全相同的额外行：', duplicates, '笔（保留，不擅自删除）')

    # 所有数据库操作均在一个事务内；用户确认前不进行 INSERT。
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT date, name, identity, topic, session, time FROM public.reading')
            existing = cur.fetchall()
            missing = plan(rows, existing)
            print('云端现有 reading 记录：', len(existing))
            print('准备补入：', len(missing), '笔')
            print('已经存在而跳过：', len(rows) - len(missing), '笔')
            if not missing:
                print('无需导入；没有修改数据库。')
                return
            print('\n注意：确认后只补入缺少的历史记录，不删除、不覆盖云端任何记录。')
            answer = input('确认写入 Supabase？输入 yes 并按 Enter；其他输入取消：').strip()
            if answer != 'yes':
                print('已取消，没有修改数据库。')
                return
            # 事务中重新检查，防止等待确认期间已有其他导入。
            cur.execute('SELECT date, name, identity, topic, session, time FROM public.reading')
            refreshed = cur.fetchall()
            missing = plan(rows, refreshed)
            for r in missing:
                cur.execute('''
                    INSERT INTO public.reading (date, name, identity, topic, session, time)
                    VALUES (%s, %s, %s, %s, %s, %s)
                ''', tuple(r[f] if r[f] != '' else None for f in FIELDS))
            print('已提交导入：', len(missing), '笔。')
    print('请运行：python reports_audit.py 2026 10，再检查白话佛法记录数。')


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print('导入中止（数据库事务将回滚）：', type(exc).__name__, str(exc))
        raise SystemExit(1)
