# 第五阶段：负责人管理菜单路由（从 schedule_web.py 原样搬迁）
from datetime import date
from flask import request, session, redirect, url_for, render_template_string
from schedule.blueprint import schedule_bp
from schedule.builders.time_utils import malaysia_today
from schedule.schedule_page_templates import _MANAGEMENT_MENU_HTML


@schedule_bp.route("/schedule/management/fixed", methods=["GET"])
def schedule_management_fixed():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))
    selected_date = request.args.get("date", "").strip()
    try:
        selected_date = date.fromisoformat(selected_date).isoformat()
    except ValueError:
        selected_date = malaysia_today().isoformat()
    return render_template_string(_MANAGEMENT_MENU_HTML,
        title="🪷 佛台及固定岗位管理",
        links=[
            {"label": "🌸 供台及佛台特别安排", "url": f"/schedule?mode=special&override_date={selected_date}"},
            {"label": "👥 固定岗位人员管理", "url": "/schedule/fixed-roles"},
        ])



@schedule_bp.route("/schedule/management/settings", methods=["GET"])
def schedule_management_settings():
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))
    selected_date = request.args.get("date", "").strip()
    try:
        selected_date = date.fromisoformat(selected_date).isoformat()
    except ValueError:
        selected_date = malaysia_today().isoformat()
    return render_template_string(_MANAGEMENT_MENU_HTML,
        title="⚙️ 系统设置",
        links=[
            {"label": "🔧 报名开放控制", "url": f"/schedule?mode=controls&override_date={selected_date}"},
            {"label": "⚙️ 排班设置", "url": "/schedule/settings"},
            {"label": "📲 WhatsApp 自动通知", "url": "/schedule/whatsapp-automation"},
        ])

