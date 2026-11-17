"""观音堂一键离线综合测试。仅读取本地代码；绝不连接数据库、Render 或 WhatsApp。"""
import ast, os, subprocess, sys, json, difflib
from pathlib import Path
ROOT = Path(__file__).resolve().parent
BUILDER = ROOT/'schedule'/'builders'
passed=failed=0

def check(label, ok, detail=''):
    global passed,failed
    passed += bool(ok); failed += not bool(ok)
    print(('PASS' if ok else 'FAIL')+' | '+label+((' | '+detail) if detail else ''),flush=True)

def source(path):
    return path.read_text(encoding='utf-8-sig')

def function(text, name):
    tree=ast.parse(text)
    return next((ast.get_source_segment(text,n) for n in tree.body if isinstance(n,ast.FunctionDef) and n.name==name),None)

print('=== 观音堂综合安全测试：只读离线 ===',flush=True)
for file in ['test_duty_v2.py','test_duty_v2_integration.py','test_duty_v2_safety.py']:
    path=ROOT/file
    if not path.exists():
        check(file+' 存在',False,'请将旧测试文件放在项目根目录');continue
    # Explicitly remove connection credentials from child tests.
    env=os.environ.copy()
    for key in ('DATABASE_URL','SUPABASE_URL','SUPABASE_KEY','SUPABASE_SERVICE_ROLE_KEY','WA_2200_ALLOW_DB_WRITE'):
        env.pop(key,None)
    result=subprocess.run([sys.executable,str(path)],cwd=ROOT,env=env,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,timeout=60)
    lines=result.stdout.strip().splitlines()
    print('\n'.join(lines[-7:]),flush=True)
    check(file+' 完整测试',result.returncode==0)

original=BUILDER/'schedule_builder.py'
patched=BUILDER/'schedule_builder_2200_snapshot_safe.py'
for path in (original,patched):
    check(str(path.relative_to(ROOT))+' 存在',path.exists())
if original.exists() and patched.exists():
    a,b=source(original),source(patched)
    try:
        ast.parse(a);ast.parse(b)
        check('正式版与快照测试版 Python 语法',True)
        diff=list(difflib.unified_diff(a.splitlines(),b.splitlines()))
        removed=[x for x in diff if x.startswith('-') and not x.startswith('---') and x[1:].strip()]
        check('快照版未删除正式版非空代码行',not removed,str(len(removed))+' removed')
        helper=function(b,'_capture_2200_baseline_in_transaction')
        check('事务快照辅助函数存在',bool(helper))
        if helper:
            check('快照不覆盖既有基准', 'ON CONFLICT(target_date) DO NOTHING' in helper)
            check('快照复用同一事务游标', 'cur.execute' in helper and 'get_conn(' not in helper and 'commit(' not in helper)
        for name in ('publish_full_schedule_2200_atomic','_publish_existing_2200_atomic'):
            f=function(b,name)
            check(name+' 发布提交前建立快照',bool(f) and '_capture_2200_baseline_in_transaction(cur, target)' in f and f.index('_capture_2200_baseline_in_transaction(cur, target)')<f.rfind('conn.commit()'))
        check('快照函数只被预期路径调用',b.count('_capture_2200_baseline_in_transaction(cur, target)')==3)
    except Exception as e:check('解析快照版本',False,repr(e))

# The following checks inspect source only, never import service modules.
services=ROOT/'schedule'/'services'
for filename, alternative in [('whatsapp_publish_2200.py',ROOT/'whatsapp_publish_2200(3).py'),('whatsapp_morning_summary.py',ROOT/'whatsapp_morning_summary(2).py'),('publish_service.py',ROOT/'publish_service(4).py')]:
    path=services/filename
    if not path.exists() and alternative.exists():path=alternative
    if not path.exists():
        check(filename+' 文件可读取',False);continue
    try:
        t=source(path);ast.parse(t);check(filename+' 语法',True)
        if filename=='whatsapp_publish_2200.py':
            check('22点流程有快照调用', 'capture_2200_snapshot(target)' in t)
            check('22点发送存在防重复查询', 'auto_roster_2200:' in t and 'whatsapp_notification_jobs' in t)
            check('22点正式发布有写入保护', 'WA_2200_ALLOW_DB_WRITE' in t)
        elif filename=='whatsapp_morning_summary.py':
            check('早间缺快照停止发送', 'missing_2200_snapshot' in t)
            check('早间批次有去重账本', 'whatsapp_morning_batches' in t and 'already_queued' in t)
            check('早间快照不可覆盖', 'ON CONFLICT(target_date) DO NOTHING' in t)
        else:
            check('手动发布建立快照', 'capture_2200_snapshot' in t)
            check('手动发布正式表入队', 'finalize_published_roster' in t)
    except Exception as e:check(filename+' 解析',False,repr(e))

print(f'\n=== 总结果：PASS {passed} / FAIL {failed} ===',flush=True)
print('注意：这是离线源码及模拟测试，不代表真实 PostgreSQL 回滚或 WhatsApp 实际发送已验证。',flush=True)
if failed:print('有失败项：不要部署；将完整输出发回检查。',flush=True)
else:print('离线检查通过；部署前仍需确认实际数据库事务与通知流程。',flush=True)
sys.exit(1 if failed else 0)
