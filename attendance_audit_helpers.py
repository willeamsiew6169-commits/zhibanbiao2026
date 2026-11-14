# -*- coding: utf-8 -*-
"""独立导入历史实际出席记录。默认只预览；输入 yes 才写入 Supabase。
放在项目根目录，与 db.py 同目录；Excel 从指定的值班报表系统目录读取。
"""
from __future__ import annotations

import hashlib
import json
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook
from psycopg2.extras import execute_values
from db import get_conn

BASE = Path(__file__).resolve().parent
EXCEL_DIR = Path(r"C:\gyt_system\值班报表系统")

ATTENDANCE = EXCEL_DIR / "attendance.xlsx"
VOLUNTEERS = EXCEL_DIR / "volunteers.xlsx"
CUTOFF = date(2026, 11, 30)
SOURCE = 'local_excel'
BATCH_SIZE = 500
# 已人工确认的旧/新编号对应；仅供本次历史导入关联使用。
ID_ALIASES = {'CHE-243': '837'}
# 仅允许指定编号的历史姓名字形差异，不放宽其他人的姓名核对。
NAME_ALIASES = {('CHE-226', '钟镒壕', '锺镒壕')}
ROLES = {'值班', '卫生', '佛台', '供台', '供花', '供果', '膳食', '佛学班', '整理佛台'}
# 只用于匹配姓名，不修改姓名或数据库
TRAD = '蕭倫麗儀蘇葉賜陳劉黃吳張楊趙鄭謝鍾馮許羅鄧盧龔龍鳳華國學會義體關愛歡樂禮佛經號聯絡電話門東萬與為來時後裡麼這個們從還過見長開發現實應當讓點無處間問題資料檔案審核轉會費續狀態舊新寫報名號碼備註職業區歲歷師專屬請鏈結個資緊急親係單雙總數個人檔案'
SIMP = '萧伦丽仪苏叶赐陈刘黄吴张杨赵郑谢钟冯许罗邓卢龚龙凤华国学会义体关爱欢乐礼佛经号联络电话门东万与为来时后里么这个们从还过见长开发现实应当让点无处间问题资料档案审核转会费续状态旧新写报名号码备注职业区岁历师专属请链结个资紧急亲系单双总数个人档案'
TRANS = str.maketrans(dict(zip(TRAD, SIMP)))


def norm_name(value):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(value or '')).strip()).translate(TRANS).casefold()


def norm_id(value):
    return str(value or '').strip().upper()


def flag(value):
    if value is True or value == 1:
        return True
    return str(value or '').strip().casefold() in {'1', '1.0', 'true', 'yes', 'y', '是', '有', '已签到', '签到', '√', '✓'}


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value or '').strip()
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%d-%m-%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f'无法识别日期：{value!r}')


def parse_time(value):
    if value is None or str(value).strip() == '':
        return None
    if isinstance(value, datetime):
        return value.time().replace(microsecond=0)
    if isinstance(value, time):
        return value.replace(microsecond=0)
    if isinstance(value, (int, float)) and 0 <= value < 1:
        seconds = round(value * 86400) % 86400
        return time(seconds // 3600, (seconds % 3600) // 60, seconds % 60)
    s = re.sub(r'\s+', '', str(value).strip()).lower().replace('a.m.', 'am').replace('p.m.', 'pm')
    s = s.rstrip(':')
    s = re.sub(r'^(\d{1,2})\.(\d{2})(am|pm)?$', r'\1:\2\3', s)
    for fmt in ('%I:%M%p', '%I%p', '%H:%M', '%H:%M:%S'):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    raise ValueError(f'无法识别时间：{value!r}')


def parse_hours(value):
    if value is None or str(value).strip() == '':
        return None
    try:
        number = Decimal(str(value).strip())
    except InvalidOperation:
        raise ValueError(f'无法识别时数：{value!r}')
    if not number.is_finite() or number < 0 or number > 24:
        raise ValueError(f'时数不合理：{value!r}')
    return number.quantize(Decimal('0.01'))


def read_rows(path, sheet):
    wb = load_workbook(path, read_only=True, data_only=True)
    if sheet not in wb.sheetnames:
        raise ValueError(f'{path.name} 缺少工作表 {sheet!r}；实际有 {wb.sheetnames}')
    ws = wb[sheet]
    iterator = ws.iter_rows(values_only=True)
    headers = [str(v or '').strip() for v in next(iterator)]
    for row_no, values in enumerate(iterator, 2):
        if not any(v is not None and str(v).strip() for v in values):
            continue
        yield row_no, dict(zip(headers, values))
    wb.close()


def main():
    for path in (ATTENDANCE, VOLUNTEERS):
        if not path.is_file():
            raise FileNotFoundError(f'找不到 {path}；请检查 EXCEL_DIR 指向的值班报表系统目录。')

    local_by_name = defaultdict(set)
    local_by_id = {}
    for row_no, r in read_rows(VOLUNTEERS, 'volunteers'):
        vid, name = norm_id(r.get('编号')), str(r.get('姓名') or '').strip()
        if vid and name:
            local_by_name[norm_name(name)].add(vid)
            if vid in local_by_id and norm_name(local_by_id[vid]) != norm_name(name):
                raise ValueError(f'本地义工编号重复且姓名不同：{vid}')
            local_by_id[vid] = name

    # 只读查询云端实际义工编号；不能仅凭姓名或月费编号擅自合并。
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT id, name, member_id FROM public.volunteers')
            cloud_rows = cur.fetchall()
    cloud_by_id = {
        norm_id(r["id"]): str(r["name"] or "").strip()
        for r in cloud_rows
    }
    cloud_by_name = defaultdict(set)
    for vid, name in cloud_by_id.items():
        cloud_by_name[norm_name(name)].add(vid)

    pending = []
    issues = []
    signed = 0
    skipped_unsigned = 0
    skipped_future = 0
    seen_payloads = Counter()
    seen_ids = Counter()

    for row_no, r in read_rows(ATTENDANCE, 'records'):
        if not flag(r.get('签到')):
            skipped_unsigned += 1
            continue
        signed += 1
        try:
            day = parse_date(r.get('日期'))
            if day > CUTOFF:
                skipped_future += 1
                continue
            name = str(r.get('姓名') or '').strip()
            role = str(r.get('岗位') or '').strip()
            if not name or role not in ROLES:
                raise ValueError(f'姓名为空或未知岗位：{name!r} / {role!r}')
            normalized = norm_name(name)
            local_ids = local_by_name.get(normalized, set())
            if len(local_ids) != 1:
                raise ValueError(f'本地姓名对应编号不是唯一：{name} → {sorted(local_ids)}')
            local_id = next(iter(local_ids))
            # 优先按本地编号查云端，并强制核对姓名；若本地编号与云端不一致，停止人工确认。
            cloud_id = ID_ALIASES.get(local_id, local_id)
            if cloud_id not in cloud_by_id:
                raise ValueError(f'本地编号 {local_id} 对应云端编号 {cloud_id} 不存在；云端同名候选：{sorted(cloud_by_name.get(normalized, set()))}')
            cloud_name = cloud_by_id[cloud_id]
            if norm_name(cloud_name) != normalized and (cloud_id, name, cloud_name) not in NAME_ALIASES:
                raise ValueError(f'编号 {cloud_id} 的云端姓名 {cloud_name!r} 与历史姓名 {name!r} 不符')
            start = parse_time(r.get('开始时间'))
            end = parse_time(r.get('结束时间'))
            hours = parse_hours(r.get('时数'))
            remarks = str(r.get('备注') or '').strip()
            # 内容稳定指纹，不依赖 Excel 行号；完全相同的多条记录用出现序号区分。
            signature = json.dumps([local_id, day.isoformat(), role,
                                    start.isoformat() if start else None,
                                    end.isoformat() if end else None,
                                    str(hours) if hours is not None else None,
                                    remarks], ensure_ascii=False, separators=(',', ':'))
            seen_payloads[signature] += 1
            key = hashlib.sha256(f'{signature}|{seen_payloads[signature]}'.encode('utf-8')).hexdigest()
            pending.append((cloud_id, cloud_name, day, role, start, end,
                            hours, SOURCE, key, remarks))
            seen_ids[cloud_id] += 1
        except (ValueError, TypeError) as exc:
            issues.append(f'第 {row_no} 行：{exc}')

    print('\n========== 历史出席导入预览（尚未写入）==========')
    print(f'本地义工：{len(local_by_id)}；云端义工：{len(cloud_by_id)}')
    print(f'已签到原始记录：{signed}；待导入历史记录：{len(pending)}')
    print(f'未签到跳过：{skipped_unsigned}；超过 {CUTOFF} 跳过：{skipped_future}')
    print(f'涉及义工：{len(seen_ids)}；资料异常：{len(issues)}')
    print(f'完全相同内容的额外记录：{sum(v-1 for v in seen_payloads.values() if v > 1)}（保留原始笔数）')
    if issues:
        print('\n❌ 有资料异常，禁止上传。前50项：')
        for issue in issues[:50]:
            print(' -', issue)
        if len(issues) > 50:
            print(f'... 另外 {len(issues)-50} 项')
        print('请核对 Excel 和云端义工资料后重试。')
        return
    if not pending:
        print('没有可上传的历史出席记录。')
        return

    # 再查已有键，提供重复导入预览；只读。
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT source_row_key FROM public.volunteer_attendance_history WHERE source=%s', (SOURCE,))
            existing = {r["source_row_key"] for r in cur.fetchall()}
    new_count = sum(1 for row in pending if row[8] not in existing)
    print(f'云端已存在相同记录：{len(pending)-new_count}；预计新增：{new_count}')
    if new_count == 0:
        print('没有新增资料，已结束。')
        return
    print('\n⚠️ 本次只导入签到=1、日期不晚于2026-11-30的记录。')
    print('⚠️ 不会删除、覆盖旧记录；原 Excel 内容变更可能产生新指纹，请先人工核对。')
    answer = input('确认正式上传 Supabase？请输入 yes：').strip()
    if answer != 'yes':
        print('⏭ 已取消，云端资料未修改。')
        return

    sql = '''INSERT INTO public.volunteer_attendance_history
        (volunteer_id, volunteer_name, attendance_date, role, start_time, end_time,
         hours, source, source_row_key, remarks) VALUES %s
        ON CONFLICT (source, source_row_key) DO NOTHING'''
    inserted = 0
    with get_conn() as conn:
        with conn.cursor() as cur:
            for offset in range(0, len(pending), BATCH_SIZE):
                chunk = pending[offset:offset+BATCH_SIZE]
                execute_values(cur, sql, chunk, page_size=BATCH_SIZE)
                inserted += cur.rowcount
    print(f'✅ 导入完成，新增 {inserted} 笔；重复记录自动跳过。')


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(f'❌ 导入中止：{exc}')
        raise
