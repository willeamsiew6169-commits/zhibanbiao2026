# -*- coding: utf-8 -*-
"""2026-11-02~14 白话佛法精准补录；默认只预览，输入 yes 才写入。"""
from collections import Counter
from datetime import datetime, date
from pathlib import Path
import pandas as pd
from db import get_conn

ROOT = Path(__file__).resolve().parent
FILE = ROOT / 'reading.xlsx'
START, END = date(2026, 11, 2), date(2026, 11, 14)

def clean(v):
    return '' if pd.isna(v) else str(v).strip()

def main():
    if not FILE.exists():
        raise SystemExit(f'找不到 {FILE}，请将最新 Excel 命名 reading.xlsx 放在此程序同一目录')
    df = pd.read_excel(FILE, dtype=object)
    needed = {'日期', '姓名', '主题', '场次', '时间'}
    if not needed.issubset(df.columns):
        raise SystemExit('Excel 缺少栏位：' + ', '.join(sorted(needed - set(df.columns))))
    records = []
    for i, row in df.iterrows():
        raw_date = row['日期']

        if isinstance(raw_date, (int, float)):
            dt = pd.to_datetime(raw_date, unit='D', origin='1899-12-30', errors='coerce')
        else:
            dt = pd.to_datetime(raw_date, errors='coerce')
        if pd.isna(dt):
            continue
        day = dt.date()
        if START <= day <= END:
            name = clean(row['姓名'])
            if not name:
                raise SystemExit(f'第 {i+2} 行缺少姓名，已停止')
            records.append((day, name, clean(row.get('身份')) or '义工', clean(row['主题']), clean(row['场次']), clean(row['时间'])))
    # 既有管理员补点名以 (date,name) 判断重复；这里采用同样的保守规则。
    unique = {}
    for r in records:
        key = r[:2]
        if key in unique and unique[key] != r:
            raise SystemExit(f'Excel 同日同名存在不同内容：{key}，请人工核对后再运行')
        unique[key] = r
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT date::date,name,identity,topic,session,time FROM public.reading WHERE date::date BETWEEN %s AND %s', (START, END))
            cloud = cur.fetchall()
            existing = {(r[0], clean(r[1])): r for r in cloud}
            pending = [r for key, r in unique.items() if key not in existing]
            print(f'本地目标日期：{len(records)} 行；唯一日期+姓名：{len(unique)}；云端已有：{len(cloud)}；预计新增：{len(pending)}')
            for r in pending:
                print('待新增', r[0], r[1], r[3])
            if not pending:
                print('没有待补录记录；未修改数据库。')
                return
            answer = input('核对以上名单无误后，输入 yes 正式补录（其他输入均取消）：').strip()
            if answer != 'yes':
                print('已取消；未修改数据库。')
                return
            inserted = 0
            for r in pending:
                cur.execute('SELECT pg_advisory_xact_lock(hashtext(%s),hashtext(%s))', (str(r[0]), r[1]))
                cur.execute('SELECT 1 FROM public.reading WHERE date::date=%s AND name=%s LIMIT 1', r[:2])
                if cur.fetchone():
                    print('跳过云端已有', r[0], r[1])
                    continue
                cur.execute('INSERT INTO public.reading (date,name,identity,topic,session,time) VALUES (%s,%s,%s,%s,%s,%s)', r)
                inserted += 1
            print(f'事务准备提交：新增 {inserted} 笔。')
    print('提交完成。请回 Supabase 核对日期及总数。')

if __name__ == '__main__':
    main()
