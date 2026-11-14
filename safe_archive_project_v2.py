# -*- coding: utf-8 -*-
"""安全归档 v2：排除自身误判；候选文件组内引用不阻止整组归档。"""
from pathlib import Path
from datetime import datetime
import subprocess
import shutil

ROOT = Path(__file__).resolve().parent
DEST = ROOT.parent / '系统历史归档' / ('检查工具_' + datetime.now().strftime('%Y%m%d_%H%M%S'))
CANDIDATES = [
    'attendance_audit_helpers.py', 'audit_attendance_sync.py',
    'audit_october_remaining.py', 'audit_october_v2.py',
    'build_october_integrated.py', 'build_october_v2.py',
    'build_october_v3.py', 'build_october_v4_audit.py',
    'compare_october_reports.py', 'test_2200_now.py',
    'wa_2200_safe_merge_stage3_test.py', 'wa_2200_stage4_readonly.py',
    'wa_2200_stage5b_concurrency.py', 'wa_2200_stage5_transaction_lab.py',
]


def tracked(path):
    try:
        r = subprocess.run(['git', 'ls-files', '--error-unmatch', '--', path.name],
                           cwd=ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           timeout=10)
        return r.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return None


def references(path, ignored_names):
    stem, name = path.stem, path.name
    found = []
    for p in ROOT.rglob('*.py'):
        if p == path or p.name in ignored_names or any(x in p.parts for x in
                ('.git', '__pycache__', '.venv', 'venv')):
            continue
        try:
            s = p.read_text(encoding='utf-8-sig')
        except (UnicodeError, OSError):
            found.append(str(p.relative_to(ROOT)) + '（无法读取，保守保留）')
            continue
        if stem in s or name in s:
            found.append(str(p.relative_to(ROOT)))
    return found


if not (ROOT / 'app.py').is_file():
    raise SystemExit('未找到 app.py：请把脚本放在项目根目录运行。')

# 只忽略本脚本，以及同一批候选工具之间的相互引用；
# 不忽略任何正式运行模块的引用。
ignored = set(CANDIDATES) | {Path(__file__).name, 'safe_archive_project.py'}
print('项目：', ROOT)
print('归档目标（项目目录外）：', DEST)
print('本次仅检查14个审计/测试候选文件，不删除任何文件。\n')
plan = []
for name in CANDIDATES:
    p = ROOT / name
    if not p.is_file():
        print('不存在：', name)
        continue
    refs = references(p, ignored)
    git = tracked(p)
    if refs:
        status = '保留：候选组外文件存在引用或读取异常'
    elif git is True:
        status = '只复制备份：Git 已跟踪，原文件保留'
        plan.append((p, 'copy'))
    elif git is None:
        status = '只复制备份：无法确认 Git 状态，原文件保留'
        plan.append((p, 'copy'))
    else:
        status = '可移动：Git 未跟踪，未发现组外 Python 文本引用'
        plan.append((p, 'move'))
    print(f'{name} -> {status}')
    if refs:
        print('  相关文件：', ', '.join(refs[:12]))

print(f'\n计划处理 {len(plan)} 个：复制 {sum(a == "copy" for _, a in plan)}，移动 {sum(a == "move" for _, a in plan)}。')
print('注意：静态文本扫描无法覆盖动态导入、命令行调用或部署配置。')
if input('确认执行请输入 yes（其他输入取消）：').strip() != 'yes':
    raise SystemExit('已取消，未修改文件。')

DEST.mkdir(parents=True, exist_ok=True)
for p, action in plan:
    target = DEST / p.name
    if target.exists():
        print('跳过（目标已存在）：', p.name)
        continue
    if action == 'copy':
        shutil.copy2(p, target)
        print('已复制：', p.name)
    else:
        shutil.move(str(p), str(target))
        print('已移动：', p.name)
print('完成。请在部署前检查 git status，并测试项目启动。')
