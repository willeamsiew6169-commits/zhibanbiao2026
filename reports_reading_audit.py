# -*- coding: utf-8 -*-
"""只读核对 reading.xlsx 与 Supabase reading；不修改报表、数据库或原始 Excel。
运行: python reports_reading_audit.py 2026 10
"""
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path
import pandas as pd
from reports_service import load_reading_rows


def clean_text(v):
    if v is None or pd.isna(v):
        return ''
    return str(v).strip()


def clean_day(v):
    if v is None or pd.isna(v):
        return ''
    d = pd.to_datetime(v, errors='coerce')
    return '' if pd.isna(d) else d.strftime('%Y-%m-%d')


def normalized(rows, cloud=False):
    result = []
    for row in rows:
        result.append((
            clean_day(row.get('date') if cloud else row.get('日期')),
            clean_text(row.get('name') if cloud else row.get('姓名')),
            clean_text(row.get('topic') if cloud else row.get('主题')),
            clean_text(row.get('session') if cloud else row.get('场次')),
            clean_text(row.get('time') if cloud else row.get('时间')),
        ))
    return result


def audit(year, month):
    path = Path('reading.xlsx')
    if not path.is_file():
        raise FileNotFoundError('请在包含 reading.xlsx 的项目目录运行')
    local_df = pd.read_excel(path, dtype=object)
    required = {'日期', '姓名', '主题', '场次', '时间'}
    missing = required - set(local_df.columns)
    if missing:
        raise ValueError('本地 reading.xlsx 缺少字段：' + ', '.join(sorted(missing)))
    local = [r for r in normalized(local_df.to_dict('records')) if r[0].startswith(f'{year}-{month:02d}-')]
    cloud = normalized(load_reading_rows(year, month), cloud=True)
    local_counter, cloud_counter = Counter(local), Counter(cloud)
    only_local = local_counter - cloud_counter
    only_cloud = cloud_counter - local_counter
    print(f'核对月份：{year}-{month:02d}')
    print(f'本地 reading.xlsx：{len(local)} 笔；Supabase reading：{len(cloud)} 笔')
    print(f'逐笔完全匹配：{sum((local_counter & cloud_counter).values())} 笔')
    print(f'仅本地：{sum(only_local.values())} 笔；仅云端：{sum(only_cloud.values())} 笔')

    local_names, cloud_names = Counter(r[1] for r in local), Counter(r[1] for r in cloud)
    people = sorted(set(local_names) | set(cloud_names))
    summary = pd.DataFrame([{'姓名': n, '本地次数': local_names[n], '云端次数': cloud_names[n],
                             '次数差异': cloud_names[n] - local_names[n]} for n in people])
    detail_rows = []
    for label, counter in [('仅本地', only_local), ('仅云端', only_cloud)]:
        for key, count in sorted(counter.items()):
            detail_rows.append(dict(zip(['差异来源', '日期', '姓名', '主题', '场次', '时间', '笔数'],
                                        [label, *key, count])))
    stem = f'{year}_{month:02d}_白话佛法核对'
    summary.to_csv(stem + '_逐人次数.csv', index=False, encoding='utf-8-sig')
    pd.DataFrame(detail_rows, columns=['差异来源', '日期', '姓名', '主题', '场次', '时间', '笔数']).to_csv(
        stem + '_逐笔差异.csv', index=False, encoding='utf-8-sig')
    mismatch = summary[summary['次数差异'] != 0]
    print(f'逐人次数不同：{len(mismatch)} 人')
    print('已生成：' + stem + '_逐人次数.csv、' + stem + '_逐笔差异.csv')
    if not only_local and not only_cloud:
        print('结果：日期、姓名、主题、场次、时间逐笔一致（未核对身份字段）。')
    else:
        print('结果：存在差异；请上传两个 CSV 核对，不要切换正式数据来源。')


if __name__ == '__main__':
    if len(sys.argv) != 3:
        raise SystemExit('用法：python reports_reading_audit.py 2026 10')
    audit(int(sys.argv[1]), int(sys.argv[2]))
