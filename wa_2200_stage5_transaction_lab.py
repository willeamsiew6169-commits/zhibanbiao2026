"""第五阶段 A：SQLite 内存数据库事务实验。绝不连接 Supabase/Render。
运行：python wa_2200_stage5_transaction_lab.py
这不是 PostgreSQL 生产事务验证，也不执行发布或 WhatsApp。
"""
import json
import sqlite3
from wa_2200_safe_merge_stage3_test import plan, row

SCHEMA = '''
CREATE TABLE assignments (
 id INTEGER PRIMARY KEY, signup_id INTEGER, role TEXT NOT NULL,
 start_min INTEGER, end_min INTEGER, place TEXT, locked INTEGER DEFAULT 0,
 CHECK(start_min < end_min),
 UNIQUE(signup_id,role,start_min,end_min,place)
);
'''
SIGNUP = {'signup_id': 100, 'interval': (600, 840)}
THEORY = [row(100,600,720),row(100,720,840,'活动中心')]

def open_lab(existing):
    db = sqlite3.connect(':memory:', isolation_level=None)
    db.executescript(SCHEMA)
    for x in existing:
        db.execute('INSERT INTO assignments(signup_id,role,start_min,end_min,place,locked) VALUES(?,?,?,?,?,?)',
                   (x['signup_id'],x['role'],*x['interval'],x['place'],int(x.get('locked',False))))
    return db

def snapshot(db):
    return db.execute('SELECT id,signup_id,role,start_min,end_min,place,locked FROM assignments ORDER BY id').fetchall()

def read_rows(db):
    return [row(sid,s,e,place,bool(locked),role) for _,sid,role,s,e,place,locked in snapshot(db)]

def apply_in_transaction(db, theoretical=THEORY, *, force_error=False):
    # 事务内重新读取、重新规划；真实 PostgreSQL 版本仍需单独设计行锁与并发控制。
    db.execute('BEGIN IMMEDIATE')
    try:
        decision = plan(SIGNUP, read_rows(db), theoretical)
        if decision['status'] not in ('already_covered','plan_only'):
            raise ValueError(decision['status'])
        before = snapshot(db)
        for candidate in decision['add']:
            s_txt,e_txt = candidate['time'].split('-')
            def minutes(txt):
                h,m = map(int,txt.split(':'))
                return h*60+m
            s,e = minutes(s_txt),minutes(e_txt)
            # 不允许同一报名、同一岗位出现任何重叠（即使地点不同）。
            conflict = db.execute('''SELECT COUNT(*) FROM assignments WHERE signup_id=? AND role='值班'
                AND start_min<? AND end_min>?''',(SIGNUP['signup_id'],e,s)).fetchone()[0]
            if conflict:
                raise ValueError('overlap_conflict')
            db.execute('''INSERT INTO assignments(signup_id,role,start_min,end_min,place,locked)
                VALUES(?,'值班',?,?,?,0)''',(SIGNUP['signup_id'],s,e,candidate['place']))
        if force_error:
            raise RuntimeError('injected_failure')
        after = snapshot(db)
        # 既有记录必须逐行不变；只允许尾部增加记录。
        if after[:len(before)] != before:
            raise AssertionError('existing_record_changed')
        db.execute('COMMIT')
        return decision['status'], len(decision['add'])
    except Exception:
        db.execute('ROLLBACK')
        raise

def check(label, existing, theory, expected_status, expected_new, *, fail=False, twice=False):
    db = open_lab(existing)
    before = snapshot(db)
    error = None
    added = None
    try:
        status, added = apply_in_transaction(db, theory, force_error=fail)
        if twice:
            second_status, second_added = apply_in_transaction(db, theory)
            if (second_status, second_added) != ('already_covered',0):
                raise AssertionError('not_idempotent')
    except Exception as exc:
        status = type(exc).__name__
        error = str(exc)
    after = snapshot(db)
    passed = (status == expected_status and added == expected_new and
              (after == before if expected_new is None else after[:len(before)] == before))
    db.close()
    return {'scenario':label,'passed':passed,'status':status,'inserted':added,'error':error,
            'original_rows_unchanged':after[:len(before)] == before}

def main():
    cases = [
      check('空白日期插入两段',[],THEORY,'plan_only',2),
      check('已有前段只补后段',[row(100,600,720)],THEORY,'plan_only',1),
      check('已有后段只补前段',[row(100,720,840,'活动中心')],THEORY,'plan_only',1),
      check('完整排班不插入',[row(100,600,720),row(100,720,840,'活动中心')],THEORY,'already_covered',0),
      check('锁定完整记录不变',[row(100,600,840,locked=True)],THEORY,'already_covered',0),
      check('锁定部分记录整笔回滚',[row(100,600,720,locked=True)],THEORY,'ValueError',None),
      check('理论缺漏整笔回滚',[],[row(100,600,720)],'ValueError',None),
      check('理论重叠整笔回滚',[],[row(100,600,780),row(100,720,840)],'ValueError',None),
      check('模拟写入失败自动回滚',[row(100,600,720)],THEORY,'RuntimeError',None,fail=True),
      check('重复执行第二次不新增',[],THEORY,'plan_only',2,twice=True),
      check('其他岗位原记录保留',[row(100,600,840,role='卫生')],THEORY,'plan_only',2),
      check('其他报名原记录保留',[row(101,600,840)],THEORY,'plan_only',2),
    ]
    return {'stage':'STAGE5A_SQLITE_IN_MEMORY_TRANSACTION_LAB',
            'all_passed':all(x['passed'] for x in cases), 'checks':cases,
            'production_database_access':False,'supabase_writes':0,
            'published':False,'whatsapp_sent':False,'safe_to_publish':False,
            'next_step':'独立测试数据库上的 PostgreSQL 事务、锁与并发验证；不能据此开启生产写入。'}

if __name__=='__main__':
    print(json.dumps(main(),ensure_ascii=False,indent=2))
