"""手动立即触发明天的 22:00 测试流程（仅当前 Python 进程模拟22:05）。
放在项目根目录运行。不会修改电脑/Render系统时间或数据库设置。
"""
import os
from datetime import datetime
from zoneinfo import ZoneInfo

from schedule.services import whatsapp_publish_2200 as job


def main():
    now = datetime.now(ZoneInfo('Asia/Kuala_Lumpur'))
    target = now.date().toordinal() + 1
    from datetime import date
    target_date = date.fromordinal(target).isoformat()
    print('马来西亚现在：', now.strftime('%Y-%m-%d %H:%M:%S'))
    print('本次排班目标：', target_date, '（明天）')
    flags = job._settings(('wa_auto_enabled', 'wa_test_mode', 'wa_roster_2200', 'wa_auto_publish_2200'))
    print('数据库开关：', flags)
    if not all(flags.values()):
        print('停止：四个开关必须都是 true 才进行完整流程。')
        return
    answer = input('将真实写入明天的排班、发布并创建TEST_GROUP_ONLY任务。输入 TEST 执行：').strip()
    if answer != 'TEST':
        print('已取消。')
        return
    os.environ['WA_2200_ALLOW_DB_WRITE'] = '1'  # 仅本次本地进程
    real_clock = job.malaysia_now
    try:
        job.malaysia_now = lambda: now.replace(hour=22, minute=5, second=0, microsecond=0)
        print('执行结果：', job.run_2200_tick())
    finally:
        job.malaysia_now = real_clock
        os.environ.pop('WA_2200_ALLOW_DB_WRITE', None)


if __name__ == '__main__':
    main()
