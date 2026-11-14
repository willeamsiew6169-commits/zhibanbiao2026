"""年月报表中心：下载已上传至 Supabase Storage 的原版 Excel。

本页面不重新生成报表，不修改数据库；不存在的文件会明确提示。
"""
from datetime import datetime
from io import BytesIO
import os
from zoneinfo import ZoneInfo

from flask import Blueprint, abort, redirect, render_template_string, request, session, url_for, send_file

reports_bp = Blueprint('reports', __name__, url_prefix='/reports')

PAGE = '''<!doctype html>
<html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>年月报表中心 · 蕉赖观音堂</title>
<link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">
<style>
:root{--wine:#752d3d;--wine2:#9b5260;--ink:#2f3542;--muted:#78818e;--line:#e9e6e5;--bg:#f8f7f5}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:system-ui,-apple-system,'Microsoft YaHei','Noto Sans CJK SC',sans-serif}
.report-page{max-width:960px;margin:0 auto;padding:28px 18px 64px}.hero{background:linear-gradient(125deg,#632638,#a35a61);border-radius:24px;padding:36px 38px;color:#fff;box-shadow:0 14px 32px #672f3b25;position:relative;overflow:hidden}.hero:after{content:'✿';position:absolute;right:26px;top:-50px;font-size:190px;color:#ffffff12;pointer-events:none}.eyebrow{letter-spacing:2px;font-size:13px;opacity:.85}.hero h1{font-size:34px;letter-spacing:1px;margin:12px 0 8px}.hero p{margin:0;opacity:.9;font-size:16px}.status{display:inline-block;margin-top:20px;background:#ffffff24;border:1px solid #ffffff4a;border-radius:100px;padding:7px 13px;font-size:13px}
.report-card{background:white;border:1px solid var(--line);border-radius:20px;padding:27px 29px;margin-top:20px;box-shadow:0 5px 22px #20232b08}.section-top{display:flex;align-items:center;gap:13px;margin-bottom:21px}.icon{display:grid;place-items:center;width:44px;height:44px;border-radius:13px;background:#f7edef;font-size:22px}.section-top h2{margin:0;font-size:21px}.section-top p{margin:5px 0 0;font-size:14px;color:var(--muted)}.report-row{display:flex;gap:18px;flex-wrap:wrap}.field{flex:1;min-width:180px}.field label{display:block;font-size:14px;font-weight:700;margin:0 0 9px}.field select,.field input{width:100%;height:51px;padding:0 14px;border:1px solid #dedfe3;border-radius:12px;background:#fff;color:var(--ink);font-size:17px;outline:none}.field select:focus,.field input:focus{border-color:var(--wine2);box-shadow:0 0 0 3px #9b526020}.button-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}.report-button{display:flex;align-items:center;justify-content:center;gap:10px;min-height:64px;padding:14px 20px;border-radius:14px;background:var(--wine);color:white!important;text-decoration:none;border:0;font-size:17px;font-weight:700;cursor:pointer;transition:transform .15s,box-shadow .15s}.report-button:hover{transform:translateY(-2px);box-shadow:0 8px 18px #752d3d28}.report-button.alt{background:#316d68}.report-note{margin-top:20px;background:#fff8e9;border:1px solid #f0e3c9;border-radius:13px;padding:15px 18px;line-height:1.7;font-size:14px;color:#775c2f}.footer{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;align-items:center}.footer a{color:var(--wine);font-weight:700;text-decoration:none;padding:8px 0}.footer a:hover{text-decoration:underline}.error{color:#b42318;background:#fff1f1;border-radius:10px;padding:12px}.login-wrap{max-width:550px;margin:20px auto}.login-wrap .report-button{width:100%;margin-top:18px}.mini{color:var(--muted);font-size:13px;margin:0}
@media(max-width:620px){.report-page{padding:14px 12px 45px}.hero{padding:28px 22px;border-radius:19px}.hero h1{font-size:27px}.report-card{padding:21px 18px}.button-grid{grid-template-columns:1fr}.footer{flex-direction:column;align-items:flex-start}.report-row{gap:14px}}
</style></head><body><main class="report-page">
<header class="hero"><div class="eyebrow">CHERAS GUANYIN HALL · REPORT CENTER</div><h1>📊 年月报表中心</h1><p>蕉赖观音堂 · 义工出席统计与 Excel 报表</p><span class="status">☁️ 云端报表服务</span></header>
{% if not authorized %}
<section class="report-card login-wrap"><div class="section-top"><span class="icon">🔐</span><div><h2>负责人验证</h2><p>请输入管理员 PIN 进入报表中心</p></div></div><form method="post" action="{{ url_for('reports.login') }}"><div class="field"><label for="admin_pin">管理员 PIN</label><input id="admin_pin" name="admin_pin" type="password" required autocomplete="off" placeholder="请输入管理员 PIN"></div><button class="report-button" type="submit">🔓 进入报表中心</button></form>{% if error %}<p class="error" role="alert">PIN 不正确，请重试。</p>{% endif %}</section>
{% else %}
<section class="report-card"><div class="section-top"><span class="icon">📅</span><div><h2>选择报表期间</h2><p>选择需要下载的年份与月份</p></div></div><div class="report-row"><div class="field"><label for="year">报表年份</label><select id="year">{% for y in years %}<option value="{{y}}" {% if y==current_year %}selected{% endif %}>{{y}} 年</option>{% endfor %}</select></div><div class="field"><label for="month">报表月份</label><select id="month">{% for m in range(1,13) %}<option value="{{m}}" {% if m==current_month %}selected{% endif %}>{{m}} 月</option>{% endfor %}</select></div></div></section>
<section class="report-card"><div class="section-top"><span class="icon">📥</span><div><h2>下载 Excel 报表</h2><p>保留原有 Excel 报表格式</p></div></div><div class="button-grid"><a class="report-button" id="month-download" href="#">📄 下载月报 Excel ↓</a><a class="report-button alt" id="year-download" href="#">📚 下载年报 Excel ↓</a></div><div class="report-note">ℹ️ 当前下载的是 Supabase Storage 中已上传的 Excel 文件，<strong>不会自动重新计算最新历史记录</strong>。若该月份或年份的报表尚未上传，系统会提示找不到文件。</div></section>
<section class="report-card footer"><a href="{{ url_for('admin.admin_home') }}">← 返回管理员中心</a><a href="{{ url_for('reports.logout') }}">退出报表中心 ↗</a></section>
<script>
function updateReportLinks(){
 const y=document.getElementById('year').value;
 const m=document.getElementById('month').value;
 document.getElementById('month-download').href={{ url_for('reports.download',kind='month')|tojson }}+'?year='+encodeURIComponent(y)+'&month='+encodeURIComponent(m);
 document.getElementById('year-download').href={{ url_for('reports.download',kind='year')|tojson }}+'?year='+encodeURIComponent(y);
}
document.getElementById('year').addEventListener('change',updateReportLinks);
document.getElementById('month').addEventListener('change',updateReportLinks);
updateReportLinks();
</script>
{% endif %}
</main></body></html>'''


def _authorized():
    return session.get('reports_admin_authorized') is True


@reports_bp.route('/', methods=['GET'])
def home():
    now = datetime.now(ZoneInfo('Asia/Kuala_Lumpur'))
    return render_template_string(
        PAGE, authorized=_authorized(), error=request.args.get('error') == '1',
        years=range(2026, now.year + 3), current_year=now.year, current_month=now.month,
    )


@reports_bp.route('/login', methods=['POST'])
def login():
    # 复用既有管理员 PIN；不改变原有登录流程或其他模块。
    from admin_web import ADMIN_PIN
    from hmac import compare_digest
    supplied = request.form.get('admin_pin', '')
    if not supplied or not compare_digest(supplied, ADMIN_PIN):
        session.pop('reports_admin_authorized', None)
        return redirect(url_for('reports.home', error='1'))
    session['reports_admin_authorized'] = True
    return redirect(url_for('reports.home'))


@reports_bp.route('/logout', methods=['GET'])
def logout():
    session.pop('reports_admin_authorized', None)
    return redirect(url_for('reports.home'))


@reports_bp.route('/download/<kind>', methods=['GET'])
def download(kind):
    if not _authorized():
        abort(403)
    if kind not in ('month', 'year'):
        abort(404)
    try:
        year = int(request.args.get('year', ''))
        month = int(request.args.get('month', '')) if kind == 'month' else None
    except (TypeError, ValueError):
        abort(400, description='请选择正确的年份和月份')
    if not 2020 <= year <= 2100 or (kind == 'month' and not 1 <= month <= 12):
        abort(400, description='年份或月份不正确')

    # 与 upload_reports.py 的 make_remote_name() 保持一致。
    remote_name = (f'{year}_{month:02d}_month_report.xlsx'
                   if kind == 'month' else f'{year}_year_report.xlsx')
    supabase_url = os.getenv('SUPABASE_URL')
    service_key = os.getenv('SUPABASE_SERVICE_ROLE_KEY')
    if not supabase_url or not service_key:
        abort(503, description='云端报表储存尚未配置')
    from supabase import create_client
    try:
        payload = create_client(supabase_url, service_key).storage.from_('reports').download(remote_name)
    except Exception as exc:
        # 不把云端异常/密钥内容暴露给网页。
        from flask import current_app
        current_app.logger.warning('报表下载失败 %s: %s', remote_name, type(exc).__name__)
        abort(404, description=f'尚未找到 {remote_name}，请先从本地上传对应报表。')
    return send_file(BytesIO(payload), as_attachment=True, download_name=remote_name,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
