# 第六阶段：独立负责人辅助路由，原样迁移，不改变原路由或业务逻辑。
from datetime import date, datetime
from flask import request, session, redirect, url_for, jsonify, render_template_string
from db import get_conn
from schedule.blueprint import schedule_bp


@schedule_bp.route("/schedule/logout")
def schedule_logout():

    session.pop("schedule_login", None)

    return redirect("/volunteer")


@schedule_bp.route("/schedule/api/admin-signup-version", methods=["GET"])
def admin_signup_version():
    """轻量检查报名和排班变化；负责人页面只在数据改变时刷新。"""
    if not session.get("schedule_login"):
        return jsonify({"error": "unauthorized"}), 401
    day = (request.args.get("date") or "").strip()
    try:
        day = date.fromisoformat(day).isoformat()
    except (ValueError, TypeError):
        return jsonify({"error": "invalid date"}), 400
    import hashlib
    import json
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT id, volunteer_id, name, role, start_time, end_time,
                       status, assigned_place, remarks
                FROM volunteer_schedule_signups
                WHERE signup_date = %s ORDER BY id
            """, (day,))
            signups = cur.fetchall()
            cur.execute("""
                SELECT volunteer_id, role, shift_label, assigned_place,
                       start_time, end_time, status, locked_by_admin
                FROM volunteer_schedule_assignments
                WHERE assignment_date = %s
                ORDER BY volunteer_id, role, shift_label, assigned_place, start_time, end_time
            """, (day,))
            assignments = cur.fetchall()
            cur.execute("""
                SELECT is_published, need_republish, last_schedule_change
                FROM schedule_day_flags WHERE flag_date = %s
            """, (day,))
            flags = cur.fetchall()
    payload = json.dumps([signups, assignments, flags], default=str, ensure_ascii=False)
    version = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    return jsonify({"version": version})


@schedule_bp.route("/schedule/dry_run_2200", methods=["GET"])
def schedule_dry_run_2200():
    # 负责人专用，只读预演
    if not session.get("schedule_login"):
        return redirect(url_for("schedule.schedule_admin"))

    target = (request.args.get("date") or "").strip()
    try:
        target = datetime.strptime(target, "%Y-%m-%d").date().isoformat()
    except ValueError:
        return "日期格式错误，请返回负责人首页重新选择日期。", 400

    report = None
    error = None
    try:
        from schedule.services.whatsapp_publish_2200 import dry_run_2200
        report = dry_run_2200(target)
    except Exception:
        import logging
        logging.getLogger(__name__).exception("22:00 read-only dry-run failed for %s", target)
        error = "预演计算失败，请检查 Render 日志（没有执行发布或发送）。"

    return render_template_string('''<!doctype html><html lang="zh"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>22:00 自动排班安全预演</title>
<style>body{font-family:Arial,"Microsoft YaHei",sans-serif;background:#f2f6f8;color:#172a39;margin:0;padding:20px}
main{max-width:920px;margin:auto;background:white;padding:24px;border-radius:16px;box-shadow:0 4px 20px #0001}
.badge{display:inline-block;background:#d9f4e8;color:#075a35;padding:8px 12px;border-radius:9px;font-weight:bold}
.warn{background:#fff3d8;padding:12px;border-radius:9px;margin:14px 0}table{width:100%;border-collapse:collapse}th,td{padding:10px;border-bottom:1px solid #e1e8ed;text-align:left}
th{background:#eef4f7}a{display:inline-block;margin:12px 0;color:#0755b6}ul{line-height:1.9}
</style></head><body><main>
<h1>🧪 22:00 自动排班安全预演</h1><div class="badge">只读测试 · 不保存 · 不发布 · 不发送 WhatsApp</div>
<p><b>预演日期：</b>{{ target }}</p>
<a href="/schedule?mode=day&amp;override_date={{ target }}">← 返回负责人页面</a>
{% if error %}<div class="warn">❌ {{ error }}</div>{% elif report %}
<p><b>发布状态：</b>{{ '已发布' if report.published else '未发布' }}　｜　
<b>现有排班记录：</b>{{ report.existing_assignment_count }}　｜　
<b>理论排班记录：</b>{{ report.theoretical_assignment_count }}　｜　
<b>值班报名：</b>{{ report.duty_signup_count }}</p>
<div class="warn">⚠️ 当前预演只核对<strong>值班报名</strong>的时间覆盖；不代表卫生、佛台及其他岗位已通过完整验证。<br>
“未安排”指数据库当前尚无对应排班，不代表正式自动排班一定失败；请参考右侧理论排班。</div>
<h2>值班报名检查</h2>
{% if report.coverage_checks %}<div style="overflow-x:auto"><table><thead><tr><th>义工</th><th>报名时间</th><th>当前排班状态</th><th>理论安排</th></tr></thead><tbody>
{% for c in report.coverage_checks %}<tr><td>{{ c.name }}</td><td>{{ c.signup_time|join(' ~ ') }}</td>
<td>{% if c.status == 'covered' %}✅ 已覆盖{% elif c.status == 'missing' %}⚪ 当前未安排{% elif c.status == 'invalid_signup_time' %}❌ 报名时间异常{% else %}⚠️ 当前排班未完全覆盖{% endif %}</td>
<td>{% if c.theoretical %}{% for a in c.theoretical %}{{ a.assigned_place }} {{ a.start_time }}–{{ a.end_time }}<br>{% endfor %}{% else %}—{% endif %}</td></tr>{% endfor %}
</tbody></table></div>{% else %}<p>此日期没有有效值班报名。</p>{% endif %}
<h2>预演注意事项</h2><ul>{% for w in report.warnings %}<li>{{ w }}</li>{% endfor %}</ul>
{% endif %}<a href="/schedule?mode=day&amp;override_date={{ target }}">← 返回负责人页面</a>
</main></body></html>''', target=target, report=report, error=error)

