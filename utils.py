# utils.py

from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from flask import request

MY_TZ = ZoneInfo("Asia/Kuala_Lumpur")

TODAY_CODE_LIST = [
    "2580", "7312", "4901", "8625", "1047",
    "3698", "5206", "9174", "6842", "0359",
    "2468", "1357", "8080", "1122", "5566",
    "7788", "9090", "3145", "6721", "4826",
]

def get_today_code():
    today = datetime.now(MY_TZ)
    day_index = today.toordinal() % len(TODAY_CODE_LIST)
    return TODAY_CODE_LIST[day_index]


def now_date_str():
    return datetime.now(MY_TZ).strftime("%Y-%m-%d")

def parse_time(value):
    s = str(value or "").strip().lower().replace(" ", "")
    if not s:
        return None

    for fmt in ["%I:%M%p", "%I%p", "%H:%M"]:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass

    return None

def calc_hours(start_time, end_time):
    st = parse_time(start_time)
    et = parse_time(end_time)

    if not st or not et:
        return 0.0

    diff = (et - st).total_seconds() / 3600

    if diff < 0:
        return 0.0

    return round(diff, 2)

def get_lang() -> str:
    lang = request.cookies.get("lang", "zh")
    return lang if lang in TEXT else "zh"


def get_text() -> dict:
    return TEXT[get_lang()]

TEXT = {
    "zh": {
        "html_lang": "zh-Hans",
        "system_title": "蕉赖观音堂义工签到系统",
        "check_in": "签到",
        "enter_id": "输入编号",
        "id_placeholder": "例如：123",
        "pin": "PIN",
        "pin_placeholder": "输入PIN",
        "find_volunteer": "查找义工",
        "name": "姓名",
        "phone": "电话",
        "status": "状态",
        "role": "选择岗位",
        "open_records": "今日进行中（未签退）",
        "all_today_records": "今日记录",
        "show_today_records": "查看今日记录",
        "latest_records_note": "只显示最新 20 条，Excel 资料不会被删除。",
        "no_open": "现在没有未签退的义工。",
        "no_today": "今天还没有记录。",
        "start": "开始",
        "time": "时间",
        "hours": "时数",
        "action": "操作",
        "sign_out": "签退",
        "edit": "修改",
        "row_id": "编号",
        "admin_tools": "管理员工具",
        "admin_pin": "请输入管理员 PIN",
        "generate_report": "管理员登录",
        "change_pin": "修改 PIN",
        "language": "语言",
        "chinese": "中文",
        "english": "English",
        "not_found_id": "找不到编号",
        "enter_id_first": "请先输入编号。",
        "lookup_first": "请先按【查找义工】确认姓名。",
        "enter_pin": "请输入 PIN。",
        "signout_prompt": "请输入 PIN 才能签退：",
        "pin_empty": "PIN 不能为空。",
        "edit_title": "修改今日记录",
        "back_home": "返回首页",
        "date": "日期",
        "start_time": "开始时间",
        "end_time": "结束时间",
        "remark": "备注",
        "save_edit": "保存修改",
        "delete_record": "删除这笔误按记录",
        "delete_confirm": "确定删除这笔记录吗？删除前系统会自动备份。",
        "change_pin_title": "修改 PIN",
        "old_pin": "旧 PIN",
        "new_pin": "新 PIN",
        "confirm_new_pin": "确认新 PIN",
        "save": "保存",
        "bhff_title": "📖 白话佛法共修",
        "bhff_enter": "📖 进入共修记录",
        "bhff_desc": "👉 记录今日共修（最少2人）",
        "bhff_record_title": "📖 白话佛法共修记录",
        "today_code": "今日签到码",
        "today_code_placeholder": "请输入现场今日码",
        "admin_title": "🔐 管理员工具",
        "today_code_big": "今日签到码",
        "today_code_warning": "⚠ 请只写在观音堂现场，不要发群",
        "download_data": "📥 下载签到数据",
        "admin_add_record": "🛠 补录签到",
        "admin_records": "✏️ 修改 / 删除今日记录",
        "today_stats": "📊 今日统计",
        "today_checkin": "今日签到",
        "today_not_checkout": "目前未签退",
        "today_checkout_done": "已完成签退",
        "people_count": "人次",
        "people": "人",
        "paid_until": "月费已供养至",
        "pin_wrong": "PIN 不正确，无法显示个人资料",
        "today_topic": "今日主题",
        "topic_placeholder": "例如：佛陀的大智慧",
        "session_remark": "场次 / 备注",
        "session_placeholder": "例如：早上共修 / 晚上共修",
        "volunteer_list": "义工名单",
        "select_all_volunteers": "✅ 全选全部义工",
        "record_study": "✅ 记录共修",
        "extra_friend_name": "非义工佛友姓名",
        "extra_friend_placeholder": "例如：王小明、李美玲、陈先生",
        "extra_friend_tip": "多个名字用逗号、空格或顿号隔开",
        "today_recorded": "今日已记录",
        "identity": "身份",
        "topic": "主题",
        "operation": "操作",
        "today_study_count": "📊 今日共修次数",
        "count": "次数",
        "name": "姓名",
        "status": "状态",
        "phone": "电话",
        "not_registered": "未登记",
        "no_contribution": "暂无月费记录",


    },
    "en": {
        "html_lang": "en",
        "system_title": "Cheras Guan Yin Citta Volunteer Check-in System",
        "check_in": "Check In",
        "enter_id": "Volunteer ID",
        "id_placeholder": "Example: 123",
        "pin": "PIN",
        "pin_placeholder": "Enter PIN",
        "find_volunteer": "Find Volunteer",
        "name": "Name",
        "phone": "Phone",
        "status": "Status",
        "role": "Select Role",
        "open_records": "Currently Checked In",
        "all_today_records": "Today’s Records",
        "show_today_records": "View Today’s Records",
        "latest_records_note": "Only the latest 20 records are shown. Excel data is not deleted.",
        "no_open": "No volunteer is currently checked in.",
        "no_today": "No records yet today.",
        "start": "Start",
        "time": "Time",
        "hours": "Hours",
        "action": "Action",
        "sign_out": "Check Out",
        "edit": "Edit",
        "row_id": "ID",
        "admin_tools": "Admin Tools",
        "admin_pin": "Enter Admin PIN",
        "generate_report": "🔐 Enter Admin Panel",
        "change_pin": "Change PIN",
        "language": "Language",
        "chinese": "中文",
        "english": "English",
        "not_found_id": "Volunteer ID not found",
        "enter_id_first": "Please enter the volunteer ID first.",
        "lookup_first": "Please find the volunteer first.",
        "enter_pin": "Please enter PIN.",
        "signout_prompt": "Enter PIN to check out:",
        "pin_empty": "PIN cannot be empty.",
        "edit_title": "Edit Today’s Record",
        "back_home": "Back to Home",
        "date": "Date",
        "start_time": "Start Time",
        "end_time": "End Time",
        "remark": "Remark",
        "save_edit": "Save Changes",
        "delete_record": "Delete This Record",
        "delete_confirm": "Are you sure you want to delete this record? A backup will be created first.",
        "change_pin_title": "Change PIN",
        "old_pin": "Old PIN",
        "new_pin": "New PIN",
        "confirm_new_pin": "Confirm New PIN",
        "save": "Save",
        "bhff_title": "📖 Bai Hua Fo Fa (BHFF) Group Study",
        "bhff_enter": "📖 Enter Study Records",
        "bhff_desc": "👉 Record today’s group study (minimum 2 participants)",
        "bhff_record_title": "📖 Bai Hua Fo Fa (BHFF) Study Records",
        "today_code": "Today Code",
        "today_code_placeholder": "Enter today's code (on-site)",
        "admin_title": "🔐 Admin Tools",
        "today_code_big": "Today Code",
        "today_code_warning": "⚠ Display this only on-site. Do not share it in group chats.",
        "download_data": "📥 Download Check-in Data",
        "admin_add_record": "🛠 Add Record",
        "admin_records": "✏️ Edit / Delete Today’s Records",
        "today_stats": "📊 Today Summary",
        "today_checkin": "Check-ins",
        "today_not_checkout": "Not Signed Out",
        "today_checkout_done": "Completed",
        "people_count": "times",
        "people": "people",
        "paid_until": "Paid Until",
        "pin_wrong": "PIN incorrect. Cannot display personal info.",
        "today_topic": "Today’s Topic",
        "topic_placeholder": "Example: The Buddha’s Great Wisdom",
        "session_remark": "Session / Remark",
        "session_placeholder": "Example: Morning Study / Evening Study",
        "volunteer_list": "Volunteer List",
        "select_all_volunteers": "✅ Select All Volunteers",
        "record_study": "✅ Save Study Record",
        "extra_friend_name": "Non-volunteer Names",
        "extra_friend_placeholder": "Example: Wang Xiao Ming, Li Mei Ling",
        "extra_friend_tip": "Separate multiple names with commas, spaces, or new lines.",
        "today_recorded": "Today’s Records",
        "identity": "Identity",
        "topic": "Topic",
        "operation": "Action",
        "today_study_count": "📊 Today’s Study Count",
        "count": "Count",
        "name": "Name",
        "status": "Status",
        "phone": "Phone",
        "not_registered": "Not Registered",
        "no_contribution": "No Contribution Record",

    }
}