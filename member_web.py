# member_web.py

import os
import json
import psycopg2
import pandas as pd

from io import BytesIO
from opencc import OpenCC
from flask import send_file
from db import db_query, get_conn
from psycopg2.extras import RealDictCursor
from openpyxl.utils import get_column_letter
from datetime import datetime, date, timedelta
from schedule.builders.time_utils import malaysia_now
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from flask import Blueprint, request, render_template_string, redirect, url_for, flash, session, send_file
import hmac


cc = OpenCC('t2s')  # 繁 → 简

member_bp = Blueprint("member", __name__, url_prefix="/member")

@member_bp.app_template_filter("member_phone_display")
def member_phone_display(value):
    """只格式化网页显示；不修改数据库、查询、PIN 或 WhatsApp 电话号码。"""
    raw = str(value or "").strip()
    digits = "".join(ch for ch in raw if ch.isdigit())
    if digits.startswith("0") and 9 <= len(digits) <= 11:
        return digits[:3] + "-" + digits[3:]
    return raw or "-"


DATABASE_URL = os.environ.get("DATABASE_URL")

MEMBER_ADMIN_PIN = os.environ.get("MEMBER_ADMIN_PIN", "123789")
FINANCE_PIN = os.environ.get("FINANCE_PIN", "1234")

PAYMENT_YEAR = 2026
MONTHS = [f"{PAYMENT_YEAR}-{m:02d}" for m in range(1, 13)]

def to_simple(text):
    if not text:
        return ""
    return cc.convert(str(text).strip())


def search_volunteer_names(keyword, volunteers):
    keyword_simple = to_simple(keyword)

    if not keyword_simple:
        return []

    results = []

    for v in volunteers:
        name = str(v.get("name") or v.get("姓名") or "").strip()
        phone = str(v.get("phone") or v.get("电话") or v.get("电话号码") or "").strip()

        name_simple = to_simple(name)

        if keyword_simple in name_simple or keyword_simple in phone:
            results.append(v)

    return results

def check_missing_months(paid_months):
    if not paid_months:
        return []

    paid_set = set(paid_months)
    sorted_months = sorted(paid_months)

    first = sorted_months[0]
    last = sorted_months[-1]

    expected = []
    start_m = int(first.split("-")[1])
    end_m = int(last.split("-")[1])
    year = first.split("-")[0]

    for m in range(start_m, end_m + 1):
        month = f"{year}-{m:02d}"
        if month not in paid_set:
            expected.append(month)

    return expected

def normalize_volunteer_id(keyword):
    keyword = str(keyword).strip().upper()

    if keyword.startswith("CHE-") or keyword.startswith("STW-"):
        return keyword

    if keyword.isdigit():
        return f"CHE-{int(keyword)}"

    return keyword

def get_member_query_volunteer():
    return session.get("member_query_volunteer")


def normalize_member_id(raw):
    s = str(raw or "").strip().upper()

    if not s:
        return ""

    # 208 -> CHE-208
    if s.isdigit():
        if s.startswith("0"):
            # 0208 / 0160 这种先保留给 STW 逻辑
            return "STW-" + s.lstrip("0")
        return "CHE-" + s

    return s

def find_volunteer_for_member(cur, member):
    member_id = str(member.get("member_id") or "").strip()
    name = str(member.get("name") or "").strip()
    phone = str(member.get("phone") or "").strip()

    # 1. 先用编号找
    cur.execute("""
        select *
        from volunteers
        where id = %s
        limit 1
    """, (member_id,))
    vol = cur.fetchone()
    if vol:
        return vol

    # 2. 再用电话找
    if phone:
        cur.execute("""
            select *
            from volunteers
            where phone = %s
            limit 1
        """, (phone,))
        vol = cur.fetchone()
        if vol:
            return vol

    # 3. 最后用姓名找
    if name:
        cur.execute("""
            select *
            from volunteers
            where name = %s
            limit 1
        """, (name,))
        vol = cur.fetchone()
        if vol:
            return vol

    return None


def verify_member_pin(member, pin):
    pin = str(pin or "").strip()

    db_pin = str(member.get("pin") or "").strip()
    phone = str(member.get("phone") or "").strip()
    default_pin = phone[-4:] if len(phone) >= 4 else ""

    if db_pin:
        return pin == db_pin

    return pin == default_pin


from datetime import datetime
import pandas as pd

def parse_month(value):

    if value is None or pd.isna(value):
        return None

    # Excel 日期
    if isinstance(value, datetime):
        return value.date().replace(day=1)

    text = str(value).strip()

    # Nov-24
    for fmt in (
        "%b-%y",
        "%B-%y",
        "%Y-%m",
        "%Y/%m",
    ):
        try:
            return datetime.strptime(text, fmt).date().replace(day=1)
        except:
            pass

    # 最后再交给 pandas
    return pd.to_datetime(text).date().replace(day=1)


@member_bp.route("/query-logout")
def member_query_logout():
    session.pop("member_query_volunteer", None)
    return redirect(url_for("member.member_query_login"))


@member_bp.route("/query-login", methods=["GET", "POST"])
def member_query_login():

    error = None

    if request.method == "POST":
        keyword = request.form.get("volunteer_id", "").strip()
        branch = request.form.get("branch", "CHE").strip().upper()

        if keyword.isdigit():
            volunteer_id = f"{branch}-{int(keyword)}"
        else:
            volunteer_id = normalize_volunteer_id(keyword)

        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    select id, name, status
                    from volunteers
                    where id = %s
                    limit 1
                """, (volunteer_id,))

                vol = cur.fetchone()

        if not vol:
            error = "找不到这个义工编号"
        elif vol.get("status") != "在册":
            error = "这个义工不是在册状态"
        else:
            session["member_query_volunteer"] = {
                "id": str(vol["id"]),
                "name": vol["name"]
            }
            return redirect(url_for("member.member_home"))

    return render_template_string("""
    <!doctype html>
    <html>
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="theme-color" content="#1976d2">
    <title>月费查询系统</title>

    <link rel="manifest" href="/member-manifest.json?v=3">
    <link rel="icon" href="/static/member_icon.png?v=3">
    <link rel="apple-touch-icon" href="/static/member_icon.png?v=3">

    <style>
    body {
        margin: 0;
        font-family: Arial, "Microsoft YaHei", sans-serif;
        background: #f4f6f8;
    }

    .card {
        max-width: 520px;
        margin: 60px auto;
        background: white;
        padding: 35px;
        border-radius: 18px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.12);
    }

    h1 {
        margin-top: 0;
        color: #7a0000;
        font-size: 34px;
    }

    h2 {
        font-size: 22px;
        color: #333;
    }

    .branch-row {
        display: flex;
        gap: 12px;
        margin: 20px 0;
    }

    .branch-btn{
        flex:1;
        padding:14px;
        font-size:22px;
        font-weight:bold;
        border:2px solid #ddd;
        border-radius:12px;
        background:white;
        color:#666;
        cursor:pointer;
    }

    /* CHE */
    #btnCHE.active{
        background:#16a34a;
        border-color:#16a34a;
        color:white;
    }

    /* STW */
    #btnSTW.active{
        background:#dc2626;
        border-color:#dc2626;
        color:white;
    }

    input {
        width: 100%;
        box-sizing: border-box;
        font-size: 26px;
        padding: 16px;
        border-radius: 12px;
        border: 1px solid #aaa;
        margin-bottom: 18px;
    }

    button.submit-btn {
        width: 100%;
        font-size: 26px;
        padding: 18px;
        border: none;
        border-radius: 14px;
        background: #198754;
        color: white;
        font-weight: bold;
        cursor: pointer;
    }

    .error {
        background: #ffe5e5;
        color: #b00020;
        padding: 14px;
        border-radius: 10px;
        font-size: 20px;
        margin-bottom: 15px;
    }

    .note {
        color: #666;
        font-size: 16px;
    }
    </style>
    </head>

    <body>

    <div class="card">
        <h1>月费查询系统</h1>
        <h2>请先输入值班义工编号</h2>

        {% if error %}
            <div class="error">{{ error }}</div>
        {% endif %}

        <form method="post">
            <input type="hidden" id="branch" name="branch" value="CHE">

            <div class="branch-row">
                <button type="button" id="btnCHE" class="branch-btn active" onclick="setBranch('CHE')">
                    CHE
                </button>
                <button type="button" id="btnSTW" class="branch-btn" onclick="setBranch('STW')">
                    STW
                </button>
            </div>

            <input name="volunteer_id"
                placeholder="例如：108"
                required
                autofocus>

            <button class="submit-btn" type="submit">
                进入查询
            </button>

            <p class="note">
                系统会记录查询义工与查询次数，方便财政查看。
            </p>
        </form>
    </div>

    <script>
    function setBranch(branch) {
        document.getElementById("branch").value = branch;

        document.getElementById("btnCHE").classList.remove("active");
        document.getElementById("btnSTW").classList.remove("active");

        document.getElementById("btn" + branch).classList.add("active");
    }
    </script>

    </body>
    </html>
    """, error=error)


@member_bp.route("/query-logs")
def member_query_logs():

    if not session.get("member_admin"):
        return redirect(url_for("member.member_admin"))

    mode = request.args.get("mode", "today")

    where_sql = ""
    params = []

    if mode == "today":
        where_sql = "where queried_at::date = current_date"
    elif mode == "month":
        where_sql = """
            where date_trunc('month', queried_at)
            = date_trunc('month', current_date)
        """

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(f"""
                select *
                from member_query_logs
                {where_sql}
                order by queried_at desc
                limit 300
            """, params)

            logs = cur.fetchall()

            cur.execute(f"""
                select
                    volunteer_id,
                    volunteer_name,
                    count(*) as total
                from member_query_logs
                {where_sql}
                group by volunteer_id, volunteer_name
                order by total desc
            """, params)

            summary = cur.fetchall()

        return render_template_string("""
    <!doctype html>
    <html lang="zh">
    <head>
    <meta charset="utf-8">
    <title>月费查询记录</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <link rel="stylesheet"
        href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>
    .log-filter-row{
        display:grid;
        grid-template-columns:repeat(4, 1fr);
        gap:12px;
    }

    .query-log-page .section-title{
        font-size:26px;
        margin-bottom:16px;
    }

    .query-log-page .card{
        padding:22px;
    }

    .query-log-page .btn-tool{
        min-height:54px;
        font-size:21px;
    }

    @media(max-width:700px){
        .log-filter-row{
            grid-template-columns:1fr;
        }
    }
    </style>
    </head>

    <body>

    <div class="page query-log-page">

        <h1 class="page-title">🔍 月费查询记录</h1>
        <p class="page-subtitle">查看义工查询佛友月费资料的记录</p>

        <div class="card">
            <div class="section-title">📅 筛选</div>

            <div class="log-filter-row">
                <a class="btn-tool {% if mode == 'today' %}btn-primary{% else %}btn-secondary{% endif %}"
                href="/member/query-logs?mode=today">
                    今天
                </a>

                <a class="btn-tool {% if mode == 'month' %}btn-primary{% else %}btn-secondary{% endif %}"
                href="/member/query-logs?mode=month">
                    本月
                </a>

                <a class="btn-tool {% if mode == 'all' %}btn-primary{% else %}btn-secondary{% endif %}"
                href="/member/query-logs?mode=all">
                    全部
                </a>

                <a class="btn-tool btn-warning"
                href="/member/finance-upload">
                    ← 返回管理中心
                </a>
            </div>
        </div>

        <div class="card">
            <div class="section-title">👥 义工查询统计</div>

            {% if summary %}
            <div class="table-responsive">
                <table class="record-table">
                    <tr>
                        <th>义工编号</th>
                        <th>义工姓名</th>
                        <th>查询次数</th>
                    </tr>

                    {% for s in summary %}
                    <tr>
                        <td>{{ s.volunteer_id }}</td>
                        <td>{{ s.volunteer_name }}</td>
                        <td>{{ s.total }}</td>
                    </tr>
                    {% endfor %}
                </table>
            </div>
            {% else %}
            <div class="empty-state">
                <div class="empty-title">暂无查询统计</div>
            </div>
            {% endif %}
        </div>

        <div class="card">
            <div class="section-title">📋 详细记录</div>

            {% if logs %}
            <div class="table-responsive">
                <table class="record-table">
                    <tr>
                        <th>时间</th>
                        <th>查询义工</th>
                        <th>佛友编号</th>
                        <th>佛友姓名</th>
                        <th>输入关键词</th>
                        <th>IP</th>
                    </tr>

                    {% for r in logs %}
                    <tr>
                        <td>
                        {{ (r.queried_at + timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S") if r.queried_at else "-" }}
                        </td>
                        <td>{{ r.volunteer_id }} {{ r.volunteer_name }}</td>
                        <td>{{ r.member_id }}</td>
                        <td>{{ r.member_name }}</td>
                        <td>{{ r.keyword }}</td>
                        <td>{{ r.ip_address }}</td>
                    </tr>
                    {% endfor %}
                </table>
            </div>
            {% else %}
            <div class="empty-state">
                <div class="empty-title">暂无查询记录</div>
            </div>
            {% endif %}
        </div>

    </div>

    </body>
    </html>
    """, logs=logs, summary=summary, mode=mode, timedelta=timedelta)


@member_bp.route("/", methods=["GET", "POST"])
def member_home():

    query_volunteer = get_member_query_volunteer()

    if not query_volunteer:
        return redirect(url_for("member.member_query_login"))
    
    member = None
    error = None
    paid_until = None
    summary = None
    payments = []
    candidates = []

    def load_member_payment_data(cur, real_member_id):
        cur.execute("""
            select
                coalesce(sum(amount), 0) as total_payment,
                coalesce(sum(month_count), 0) as total_months,
                max(end_month) as paid_until,
                max(payment_date) as last_payment_date
            from member_payments
            where member_id = %s
        """, (real_member_id,))
        summary_data = cur.fetchone()

        paid_until_text = None
        if summary_data and summary_data["paid_until"]:
            paid_until_text = summary_data["paid_until"].strftime("%Y年%m月")

        cur.execute("""
            select amount
            from member_payments
            where member_id = %s
            order by payment_date desc, id desc
            limit 1
        """, (real_member_id,))
        last_payment = cur.fetchone()

        if summary_data:
            summary_data["last_payment_amount"] = (
                last_payment["amount"] if last_payment else 0
            )

        cur.execute("""
            select
                payment_date,
                receipt_no,
                start_month,
                end_month,
                month_count,
                amount
            from member_payments
            where member_id = %s
            order by payment_date desc, receipt_no desc
        """, (real_member_id,))
        payment_rows = cur.fetchall()

        return summary_data, paid_until_text, payment_rows

    if request.method == "POST":
        raw_member_id = request.form.get("member_id", "").strip()
        selected_member_id = request.form.get("selected_member_id", "").strip()
        branch = request.form.get("branch", "CHE").strip().upper()
        

        try:
            with get_conn() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:

                    # 第二次提交：用户已经从候选名单选择会员
                    if selected_member_id:
                        cur.execute("""
                            select *
                            from members
                            where member_id = %s
                            limit 1
                        """, (selected_member_id,))
                        member = cur.fetchone()

                    else:
                        keyword = raw_member_id.strip()
                        digits_only = "".join(ch for ch in keyword if ch.isdigit())

                        # 情况 1：纯数字，而且 6 位或以下，当会员编号
                        if keyword.isdigit() and len(keyword) <= 6:
                            if branch == "STW":
                                member_id = normalize_member_id(f"STW-{keyword}")
                            else:
                                member_id = normalize_member_id(keyword)

                            cur.execute("""
                                select *
                                from members
                                where member_id = %s
                                limit 1
                            """, (member_id,))
                            member = cur.fetchone()

                        # 情况 2：纯数字，而且超过 6 位，当电话号码
                        elif keyword.isdigit() and len(keyword) > 6:
                            cur.execute("""
                                select *
                                from members
                                where replace(replace(coalesce(phone,''), '-', ''), ' ', '') = %s
                                order by member_id
                            """, (digits_only,))
                            candidates = cur.fetchall()

                            if len(candidates) == 1:
                                member = candidates[0]
                                candidates = []

                            elif len(candidates) > 1:
                                member = None

                            else:
                                member = None

                        # 情况 3：非纯数字，当会员编号 / 中文名 / 英文名 搜索
                        else:
                            member_id = normalize_member_id(keyword)

                            cur.execute("""
                                select *
                                from members
                                where member_id = %s
                                   or name ilike %s
                                   or english_name ilike %s
                                order by member_id
                            """, (
                                member_id,
                                f"%{keyword}%",
                                f"%{keyword}%"
                            ))
                            candidates = cur.fetchall()

                            if len(candidates) == 1:
                                member = candidates[0]
                                candidates = []

                            elif len(candidates) > 1:
                                member = None

                            else:
                                member = None

                    if candidates:
                        error = None

                    elif not member:
                        error = "找不到这个会员，请检查编号 / 姓名 / 电话"

                    else:
                        real_member_id = member["member_id"]

                        # ✅ 成功找到会员后，记录是谁查询了这个会员
                        query_volunteer = get_member_query_volunteer()

                        if query_volunteer:
                            cur.execute("""
                                insert into member_query_logs (
                                    queried_at,
                                    volunteer_id,
                                    volunteer_name,
                                    member_id,
                                    member_name,
                                    keyword,
                                    ip_address,
                                    user_agent
                                )
                                values (%s, %s, %s, %s, %s, %s, %s, %s)
                            """, (
                                malaysia_now(),
                                query_volunteer["id"],
                                query_volunteer["name"],
                                member["member_id"],
                                member["name"],
                                raw_member_id or selected_member_id,
                                request.remote_addr,
                                request.headers.get("User-Agent", "")
                            ))
                            conn.commit()

                        summary, paid_until, payments = load_member_payment_data(
                            cur,
                            real_member_id
                        )

                        # 停供不显示欠费；历史付款及最后供养月份照常保留。
                        is_stopped = (member.get("member_status") or member.get("status") or "").strip() == "停供"
                        if is_stopped and summary is not None:
                            summary["late_months"] = None
                            summary["payment_status"] = "⏸️ 已停供"

                        # 新会员在指定月份之前不产生迟付；已有缴费的月份也不回溯到开始月以前。
                        fee_start = member.get("fee_start_month")
                        if isinstance(fee_start, str):
                            fee_start = date.fromisoformat(fee_start[:7] + "-01") if fee_start else None
                        if is_stopped:
                            pass  # 停供不计算历史欠费
                        elif fee_start and date.today().replace(day=1) < fee_start:
                            if summary is not None:
                                summary["late_months"] = 0
                                summary["payment_status"] = f"🟡 尚未开始供养（{fee_start:%Y-%m}开始）"
                        elif summary and summary["paid_until"]:

                            today = date.today()
                            paid_until_date = summary["paid_until"]

                            late_months = (
                                (today.year - paid_until_date.year) * 12
                                + (today.month - paid_until_date.month)
                            )

                            if fee_start:
                                start_index = fee_start.year * 12 + fee_start.month
                                paid_index = paid_until_date.year * 12 + paid_until_date.month
                                late_months = max(0, today.year * 12 + today.month - max(start_index - 1, paid_index))
                            summary["late_months"] = late_months

                            if late_months <= 0:
                                summary["payment_status"] = "🟢 正常"

                            elif late_months <= 3:
                                summary["payment_status"] = f"🟡 已过期 {late_months} 个月"

                            else:
                                summary["payment_status"] = f"🔴 停供参考 {late_months} 个月"

                        else:
                            if summary:
                                if fee_start:
                                    today = date.today()
                                    late_months = max(0, (today.year - fee_start.year) * 12 + today.month - fee_start.month + 1)
                                    summary["late_months"] = late_months
                                    summary["payment_status"] = (f"🟡 待缴 {late_months} 个月（从{fee_start:%Y-%m}起）" if late_months else "🟢 正常")
                                else:
                                    summary["late_months"] = None
                                    summary["payment_status"] = "⚪ 暂无缴费记录"

        except Exception as e:
            error = f"系统错误：{e}"

    return render_template_string(
        MEMBER_HTML,
        member=member,
        error=error,
        paid_until=paid_until,
        summary=summary,
        payments=payments,
        candidates=candidates,
        query_volunteer=query_volunteer
    )

@member_bp.route("/admin", methods=["GET", "POST"])
def member_admin():
    error = None
    member = None
    payments = []
    summary = None
    warnings = []

    page = int(request.args.get("page", 1) or 1)
    per_page = 15
    offset = (page - 1) * per_page
    total_pages = 0

    admin_pin = request.form.get("admin_pin", "").strip()
    raw_member_id = request.values.get("member_id", "").strip()
    branch = request.values.get("branch", "CHE").strip().upper()

    if request.method == "POST":

        if not session.get("member_admin"):

            admin_pin = request.form.get("admin_pin", "")

            if admin_pin == MEMBER_ADMIN_PIN:
                session["member_admin"] = True
                return redirect("/member/admin")

            error = "管理员 PIN 不正确"
            

    if raw_member_id and not error:
        keyword = raw_member_id.strip()

        if keyword.isdigit():
            if branch == "STW":
                member_id = normalize_member_id(f"STW-{keyword}")
            else:
                member_id = normalize_member_id(keyword)
        else:
            member_id = normalize_member_id(keyword)

        try:
            with get_conn() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:

                    # ✅ 数字只查编号，避免 208 查到电话 0123992081
                    if keyword.isdigit():
                        cur.execute("""
                            select *
                            from members
                            where member_id = %s
                            limit 1
                        """, (member_id,))
                    else:
                        cur.execute("""
                            select *
                            from members
                            where member_id = %s
                               or name ilike %s
                               or english_name ilike %s
                               or phone ilike %s
                            limit 1
                        """, (
                            member_id,
                            f"%{keyword}%",
                            f"%{keyword}%",
                            f"%{keyword}%"
                        ))

                    member = cur.fetchone()

                    if not member:
                        error = "找不到这个月费编号"
                    else:
                        real_member_id = member["member_id"]

                        # ✅ 汇总
                        cur.execute("""
                            select
                                coalesce(sum(amount), 0) as total_payment,
                                coalesce(sum(month_count), 0) as total_months,
                                max(end_month) as paid_until,
                                max(payment_date) as last_payment_date
                            from member_payments
                            where member_id = %s
                        """, (real_member_id,))
                        summary = cur.fetchone()

                        # ✅ 最后付款金额
                        if summary:
                            summary["last_payment_amount"] = None

                            if summary.get("last_payment_date"):
                                cur.execute("""
                                    select amount
                                    from member_payments
                                    where member_id = %s
                                      and payment_date = %s
                                    order by id desc
                                    limit 1
                                """, (
                                    real_member_id,
                                    summary["last_payment_date"]
                                ))
                                last_row = cur.fetchone()

                                if last_row:
                                    summary["last_payment_amount"] = last_row["amount"]

                        # ✅ 总页数
                        cur.execute("""
                            select count(*) as cnt
                            from member_payments
                            where member_id = %s
                        """, (real_member_id,))
                        total_rows = cur.fetchone()["cnt"] or 0
                        total_pages = (total_rows + per_page - 1) // per_page

                        # ✅ 当前页，只显示 15 行
                        cur.execute("""
                            select
                                id,
                                payment_date,
                                receipt_no,
                                member_id,
                                start_month,
                                end_month,
                                month_count,
                                amount,
                                name
                            from member_payments
                            where member_id = %s
                            order by payment_date desc, id desc
                            limit %s offset %s
                        """, (
                            real_member_id,
                            per_page,
                            offset
                        ))
                        payments = cur.fetchall()

                        # ✅ 另外查全部记录，用来检查错误
                        cur.execute("""
                            select
                                id,
                                payment_date,
                                receipt_no,
                                start_month,
                                end_month,
                                month_count,
                                amount,
                                name
                            from member_payments
                            where member_id = %s
                            order by payment_date asc, id asc
                        """, (real_member_id,))
                        all_payments = cur.fetchall()

                        seen_months = {}
                        seen_receipts = {}

                        for r in all_payments:
                            receipt_no = r.get("receipt_no") or "-"
                            amount = float(r.get("amount") or 0)
                            month_count = int(r.get("month_count") or 0)

                            start_month = r.get("start_month")
                            end_month = r.get("end_month")
                            payment_date = r.get("payment_date")

                            # 1. 收据重复
                            if receipt_no != "-":
                                if receipt_no in seen_receipts:
                                    warnings.append(
                                        f"收据 {receipt_no}：收据编号重复，之前已经出现过"
                                    )
                                else:
                                    seen_receipts[receipt_no] = True

                            # 2. 金额不是 RM50 倍数
                            if amount % 50 != 0:
                                warnings.append(
                                    f"收据 {receipt_no}：金额 RM {amount:.2f} 不是 RM50 的倍数"
                                )

                            # 3. 月数不正确
                            if month_count <= 0:
                                warnings.append(
                                    f"收据 {receipt_no}：月数不正确"
                                )

                            # 4. 金额和月数不符合
                            if amount > 0 and month_count > 0:
                                expected_amount = month_count * 50
                                if amount != expected_amount:
                                    warnings.append(
                                        f"收据 {receipt_no}：金额 RM {amount:.2f} 和月数 {month_count} 不符合，正常应是 RM {expected_amount:.2f}"
                                    )

                            # 5. 付款日期未来
                            if payment_date and payment_date > date.today():
                                warnings.append(
                                    f"收据 {receipt_no}：付款日期是未来日期"
                                )

                            # 6. 开始结束月份检查
                            if start_month and end_month:
                                if start_month > end_month:
                                    warnings.append(
                                        f"收据 {receipt_no}：开始月份大过结束月份"
                                    )
                                else:
                                    real_month_count = (
                                        (end_month.year - start_month.year) * 12
                                        + (end_month.month - start_month.month)
                                        + 1
                                    )

                                    if real_month_count != month_count:
                                        warnings.append(
                                            f"收据 {receipt_no}：开始月份到结束月份是 {real_month_count} 个月，但记录写 {month_count} 个月"
                                        )

                                    # 7. 重复月份检查
                                    current = start_month

                                    for i in range(real_month_count):
                                        ym = current.strftime("%Y-%m")

                                        if ym in seen_months:
                                            warnings.append(
                                                f"月份重复：{ym} 已在收据 {seen_months[ym]} 记录过，现在又出现在收据 {receipt_no}"
                                            )
                                        else:
                                            seen_months[ym] = receipt_no

                                        if current.month == 12:
                                            current = current.replace(
                                                year=current.year + 1,
                                                month=1
                                            )
                                        else:
                                            current = current.replace(
                                                month=current.month + 1
                                            )

        except Exception as e:
            error = f"系统错误：{e}"

    return render_template_string(
        MEMBER_ADMIN_HTML,
        error=error,
        member=member,
        payments=payments,
        summary=summary,
        warnings=warnings,
        raw_member_id=raw_member_id,
        page=page,
        total_pages=total_pages
    )

@member_bp.route("/admin/late_members")
def admin_late_members():

    if not session.get("member_admin"):
        return redirect(url_for("member.member_admin"))

    rows = db_query("""
        select
            m.member_id,
            m.name,
            m.phone,
            m.remark,
            m.fee_start_month,
            max(p.end_month) as paid_until,
            max(p.payment_date) as last_payment_date
        from members m
        left join member_payments p
            on p.member_id = m.member_id
        where coalesce(m.member_status, m.status, '') not in ('停供', '暂停', '往生', '已往生')
          and (m.fee_start_month is null or m.fee_start_month <= date_trunc('month', current_date)::date)
        group by
            m.member_id,
            m.name,
            m.phone,
            m.remark,
            m.fee_start_month
        having
            (max(p.end_month) is not null and max(p.end_month) < date_trunc('month', current_date))
            or (max(p.end_month) is null and m.fee_start_month is not null
                and m.fee_start_month <= date_trunc('month', current_date)::date)
        order by
            paid_until asc,
            m.member_id
    """, fetchall=True)

    today = date.today()
    current_month_index = today.year * 12 + today.month

    for r in rows:
        paid_until = r["paid_until"]
        fee_start = r.get("fee_start_month")
        if fee_start:
            start_index = fee_start.year * 12 + fee_start.month
            paid_index = paid_until.year * 12 + paid_until.month if paid_until else start_index - 1
            late_months = max(0, current_month_index - max(paid_index, start_index - 1))
        else:
            paid_index = paid_until.year * 12 + paid_until.month
            late_months = current_month_index - paid_index

        r["late_months"] = late_months
        r["reference_amount"] = late_months * 50

        if late_months <= 2:
            r["level"] = "green"
        elif late_months <= 6:
            r["level"] = "yellow"
        else:
            r["level"] = "red"

        phone = (r["phone"] or "").strip()
        phone_digits = "".join(ch for ch in phone if ch.isdigit())

        if phone_digits.startswith("0"):
            phone_digits = "6" + phone_digits

        r["wa_link"] = "https://wa.me/" + phone_digits if phone_digits else ""

    green_count = sum(1 for r in rows if r["level"] == "green")
    yellow_count = sum(1 for r in rows if r["level"] == "yellow")
    red_count = sum(1 for r in rows if r["level"] == "red")
    total_amount = sum(r["reference_amount"] for r in rows)

    return render_template_string("""
    <!doctype html>
    <html lang="zh">
    <head>
    <meta charset="utf-8">
    <title>月费迟付名单</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <link rel="stylesheet"
        href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>
    .late-page .section-title{
        font-size:26px;
        margin-bottom:16px;
    }

    .late-page .card{
        padding:22px;
    }

    .late-page .btn-tool{
        min-height:54px;
        font-size:21px;
    }

    .late-row-green{
        background:#ecfdf5;
    }

    .late-row-yellow{
        background:#fffbeb;
    }

    .late-row-red{
        background:#fef2f2;
    }

    .floating-back{
        position:fixed;
        left:24px;
        bottom:24px;
        z-index:999;
        padding:14px 22px;
        border-radius:14px;
        background:#4b5563;
        color:white;
        text-decoration:none;
        font-size:18px;
        font-weight:700;
        box-shadow:0 8px 20px rgba(0,0,0,.18);
    }

    .floating-back:hover{
        background:#374151;
    }
    </style>
    </head>

    <body>

    <div class="page late-page">

        <h1 class="page-title">🌿 月费迟付名单</h1>
        <p class="page-subtitle">查看已过期月费会员，方便后续关怀联系</p>

        <div class="card">
            {% if member and (member.member_status == '停供' or (not member.member_status and member.status == '停供')) %}
        <div style="padding:14px 18px;margin:12px 0;border-radius:12px;background:#fff1d6;color:#825000;font-weight:700;">
            ⏸️ 已停供｜最后供养月份：{{ summary.paid_until.strftime('%Y-%m') if summary and summary.paid_until else '暂无记录' }}<br>
            <small>停供期间不显示欠缴月份，历史付款记录保留。</small>
        </div>
        {% endif %}

        <div class="info-summary-grid">

                <div class="info-summary-box">
                    <div class="info-summary-title">总人数</div>
                    <div class="info-summary-value">{{ rows|length }} 人</div>
                </div>

                <div class="info-summary-box">
                    <div class="info-summary-title">参考金额</div>
                    <div class="info-summary-value">RM {{ "{:,.2f}".format(total_amount) }}</div>
                </div>

                <div class="info-summary-box">
                    <div class="info-summary-title">🟢 最近缴费</div>
                    <div class="info-summary-value">{{ green_count }} 人</div>
                </div>

                <div class="info-summary-box">
                    <div class="info-summary-title">🟡 一段时间未缴费</div>
                    <div class="info-summary-value">{{ yellow_count }} 人</div>
                </div>

                <div class="info-summary-box">
                    <div class="info-summary-title">🔴 较久未缴费</div>
                    <div class="info-summary-value">{{ red_count }} 人</div>
                </div>

            </div>
        </div>

        <div class="card">
            <div class="section-title">📋 名单</div>

            {% if rows %}
            <div class="table-responsive">
                <table class="record-table">
                    <tr>
                        <th>会员编号</th>
                        <th>姓名</th>
                        <th>电话</th>
                        <th>WhatsApp</th>
                        <th>已缴至</th>
                        <th>缴费间隔</th>
                        <th>参考金额</th>
                        <th>状态</th>
                        <th>最后付款日期</th>
                        <th>备注</th>
                    </tr>

                    {% for r in rows %}
                    <tr class="
                        {% if r.level == 'green' %}
                            late-row-green
                        {% elif r.level == 'yellow' %}
                            late-row-yellow
                        {% elif r.level == 'red' %}
                            late-row-red
                        {% endif %}
                    ">
                        <td>{{ r.member_id }}</td>
                        <td>{{ r.name }}</td>
                        <td>{{ r.phone | member_phone_display }}</td>
                        <td>
                            {% if r.wa_link %}
                                <a class="btn-tool btn-success mini-btn"
                                href="{{ r.wa_link }}"
                                target="_blank">
                                    打开
                                </a>
                            {% else %}
                                -
                            {% endif %}
                        </td>
                        <td>{{ r.paid_until.strftime("%Y-%m") if r.paid_until else "尚未缴费" }}</td>
                        <td>{{ r.late_months }} 个月</td>
                        <td>RM {{ "%.2f"|format(r.reference_amount) }}</td>
                        <td>
                            {% if r.level == "green" %}
                                🟢 最近缴费
                            {% elif r.level == "yellow" %}
                                🟡 一段时间未缴费
                            {% else %}
                                🔴 较久未缴费
                            {% endif %}
                        </td>
                        <td>{{ r.last_payment_date or "-" }}</td>
                        <td>{{ r.remark or "-" }}</td>
                    </tr>
                    {% endfor %}
                </table>
            </div>
            {% else %}
            <div class="empty-state">
                <div class="empty-title">暂无迟付名单</div>
                <div class="empty-text">目前没有需要列入迟付关怀的会员。</div>
            </div>
            {% endif %}
        </div>

    </div>

    <a class="floating-back"
    href="{{ url_for('member.member_admin') }}">
        ⬅ 返回月费管理员
    </a>

    </body>
    </html>
    """,
    rows=rows,
    green_count=green_count,
    yellow_count=yellow_count,
    red_count=red_count,
    total_amount=total_amount
    )

@member_bp.route("/payment/edit/<int:payment_id>", methods=["GET", "POST"])
def edit_member_payment(payment_id):
    if not session.get("member_admin"):
        return redirect(url_for("member.member_admin"))

    error = None

    try:
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:

                cur.execute("""
                    select *
                    from member_payments
                    where id = %s
                """, (payment_id,))
                payment = cur.fetchone()

                if not payment:
                    return "找不到这笔缴费记录", 404

                if request.method == "POST":
                    payment_date = request.form.get("payment_date")
                    receipt_no = request.form.get("receipt_no", "").strip()
                    name = request.form.get("name", "").strip()
                    start_month = request.form.get("start_month")
                    end_month = request.form.get("end_month")
                    month_count = int(request.form.get("month_count") or 0)
                    amount = float(request.form.get("amount") or 0)

                    if month_count <= 0:
                        error = "月数不能小过 1"
                    elif amount <= 0:
                        error = "金额不能小过 0"
                    else:
                        old_data = dict(payment)

                        new_data = {
                            "payment_date": payment_date,
                            "receipt_no": receipt_no,
                            "name": name,
                            "start_month": start_month + "-01",
                            "end_month": end_month + "-01",
                            "month_count": month_count,
                            "amount": amount
                        }

                        cur.execute("""
                            update member_payments
                            set
                                payment_date = %s,
                                receipt_no = %s,
                                name = %s,
                                start_month = %s,
                                end_month = %s,
                                month_count = %s,
                                amount = %s
                            where id = %s
                        """, (
                            payment_date,
                            receipt_no,
                            name,
                            start_month + "-01",
                            end_month + "-01",
                            month_count,
                            amount,
                            payment_id
                        ))

                        cur.execute("""
                            insert into member_payment_history
                            (
                                payment_id,
                                member_id,
                                receipt_no,
                                action,
                                old_data,
                                new_data,
                                changed_by
                            )
                            values (%s, %s, %s, %s, %s::jsonb, %s::jsonb, %s)
                        """, (
                            payment_id,
                            payment["member_id"],
                            receipt_no,
                            "edit",
                            json.dumps(old_data, default=str),
                            json.dumps(new_data, default=str),
                            "member_admin"
                        ))

                        conn.commit()

                        return redirect(url_for("member.member_admin", member_id=payment["member_id"]))

        return render_template_string(
            PAYMENT_EDIT_HTML,
            payment=payment,
            error=error
        )

    except Exception as e:
        return f"系统错误：{e}", 500
    
@member_bp.route("/payment/history/<int:history_id>")
def member_payment_history_detail(history_id):
    if not session.get("member_admin"):
        return redirect(url_for("member.member_admin"))

    try:
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    select *
                    from member_payment_history
                    where id = %s
                """, (history_id,))
                h = cur.fetchone()

                if not h:
                    return "找不到这笔修改历史", 404

        changes = format_history_changes(h.get("old_data"), h.get("new_data"))

        return render_template_string(
            PAYMENT_HISTORY_DETAIL_HTML,
            h=h,
            changes=changes,
            timedelta=timedelta
        )

    except Exception as e:
        return f"系统错误：{e}", 500

@member_bp.route("/change-pin", methods=["GET", "POST"])
def member_change_pin():
    error = None
    ok = None

    if request.method == "POST":
        raw_member_id = request.form.get("member_id", "")
        old_pin = request.form.get("old_pin", "")
        new_pin = request.form.get("new_pin", "")
        confirm_pin = request.form.get("confirm_pin", "")

        member_id = normalize_member_id(raw_member_id)

        if not new_pin or len(new_pin) < 4:
            error = "新密码至少 4 位"
        elif new_pin != confirm_pin:
            error = "两次新密码不一样"
        else:
            try:
                with get_conn() as conn:
                    with conn.cursor(cursor_factory=RealDictCursor) as cur:
                        cur.execute("""
                            select *
                            from members
                            where member_id = %s
                            limit 1
                        """, (member_id,))
                        member = cur.fetchone()

                        if not member:
                            error = "找不到这个月费编号"
                        elif not verify_member_pin(member, old_pin):
                            error = "旧密码不正确"
                        else:
                            cur.execute("""
                                update members
                                set pin = %s
                                where member_id = %s
                            """, (new_pin, member_id))
                            conn.commit()
                            ok = "月费密码已更改"

            except Exception as e:
                error = f"系统错误：{e}"

    return render_template_string(CHANGE_PIN_HTML, error=error, ok=ok)

@member_bp.route("/admin/logout")
def member_admin_logout():
    session.pop("member_admin", None)
    return redirect(url_for("member.member_admin"))

@member_bp.route("/member-management")
def member_management():

    if not session.get("member_admin"):
        return redirect(url_for("member.member_admin"))

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>会员资料管理</title>
<meta name="viewport" content="width=device-width, initial-scale=1">

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.member-menu-grid{
    display:grid;
    grid-template-columns:repeat(auto-fit,minmax(260px,1fr));
    gap:18px;
}

.member-menu-card{
    display:block;
    text-decoration:none;
    border-radius:16px;
    padding:24px;
    color:#111827;
    transition:.2s;
}

.member-menu-card:hover{
    transform:translateY(-2px);
}

.member-menu-title{
    font-size:24px;
    font-weight:bold;
    margin-bottom:10px;
}

.member-menu-desc{
    font-size:16px;
    color:#6b7280;
    line-height:1.6;
}

.card-blue{
    background:#eff6ff;
}

.card-green{
    background:#ecfdf5;
}

.card-yellow{
    background:#fffbeb;
}

.card-gray{
    background:#f9fafb;
}
</style>

</head>
<body>

<div class="page">

    <h1 class="page-title">
        👤 会员资料管理
    </h1>

    <p class="page-subtitle">
        管理会员资料、下载资料、查看迟付名单
    </p>

    <div class="card">

        <div class="member-menu-grid">

            <a class="member-menu-card card-blue"
               href="/member/member-management/search">

                <div class="member-menu-title">
                    🔍 查找 / 编辑会员
                </div>

                <div class="member-menu-desc">
                    修改电话、PIN、备注、会员资料
                </div>

            </a>

            <a class="member-menu-card card-green"
               href="/member/download-members">

                <div class="member-menu-title">
                    📥 下载会员资料 Excel
                </div>

                <div class="member-menu-desc">
                    下载最新 members 主档
                </div>

            </a>

            <a class="member-menu-card card-yellow"
               href="/member/admin/late_members">

                <div class="member-menu-title">
                    ⚠️ 月费迟付名单
                </div>

                <div class="member-menu-desc">
                    查看已过期会员
                </div>

            </a>

            <a class="member-menu-card card-gray"
               href="/member/member-management/add">

                <div class="member-menu-title">
                    ➕ 新增会员
                </div>

                <div class="member-menu-desc">
                    新增 CHE / STW 月费会员
                </div>

            </a>

        </div>

    </div>

    <div class="card">
        <a class="btn-tool btn-secondary"
           href="{{ url_for('member.member_admin') }}">
            ← 返回月费管理员
        </a>
    </div>

</div>

</body>
</html>
""")

@member_bp.route("/download-members")
def download_members():

    if not session.get("member_admin"):
        return redirect(url_for("member.member_admin"))

    df = pd.read_sql_query("""
        select
            member_id,
            branch,
            name,
            english_name,
            ic_number,
            phone,
            pin,
            status,
            member_status,
            remark
        from members
        order by branch, member_id
    """, get_conn())

    df.columns = [
        "会员编号",
        "分会",
        "姓名",
        "英文名",
        "身份证号码",
        "电话",
        "PIN",
        "状态",
        "月费状态",
        "备注"
    ]

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="会员总表",
            startrow=2
        )

        ws = writer.sheets["会员总表"]

        ws.merge_cells("A1:J1")
        ws["A1"] = "观音堂会员资料总表"
        ws["A1"].font = Font(size=18, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill("solid", fgColor="4472C4")
        ws["A1"].alignment = Alignment(horizontal="center")

        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        blue_fill = PatternFill("solid", fgColor="DDEBF7")
        white_fill = PatternFill("solid", fgColor="FFFFFF")

        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for cell in ws[3]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        for row_num in range(4, ws.max_row + 1):
            fill = blue_fill if row_num % 2 == 0 else white_fill

            for col in range(1, 11):
                cell = ws.cell(row=row_num, column=col)
                cell.fill = fill
                cell.border = border
                cell.alignment = Alignment(vertical="center")

        ws.freeze_panes = "A4"
        ws.auto_filter.ref = f"A3:J{ws.max_row}"

        widths = {
            "A": 14,
            "B": 10,
            "C": 18,
            "D": 22,
            "E": 20,
            "F": 18,
            "G": 10,
            "H": 12,
            "I": 12,
            "J": 28,
        }

        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        ws.row_dimensions[1].height = 30
        ws.row_dimensions[3].height = 24

    output.seek(0)

    filename = f"members_{date.today().strftime('%Y-%m-%d')}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@member_bp.route("/member-management/search")
def member_management_search():

    if not session.get("member_admin"):
        return redirect(url_for("member.member_admin"))

    q = request.args.get("q", "").strip()
    rows = []

    if q:
        keyword = q
        member_id = normalize_member_id(keyword)

        if keyword.isdigit():
            rows = db_query("""
                select
                    member_id,
                    name,
                    english_name,
                    phone,
                    pin,
                    branch,
                    status,
                    remark,
                    member_status
                from members
                where member_id = %s
                order by member_id
                limit 50
            """, (member_id,), fetchall=True)

        else:
            rows = db_query("""
                select
                    member_id,
                    name,
                    english_name,
                    phone,
                    pin,
                    branch,
                    status,
                    remark,
                    member_status
                from members
                where member_id = %s
                or name ilike %s
                or english_name ilike %s
                or phone ilike %s
                order by member_id
                limit 50
            """, (
                member_id,
                f"%{keyword}%",
                f"%{keyword}%",
                f"%{keyword}%"
            ), fetchall=True)

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>查找 / 编辑会员</title>
<style>
body{
    font-family:Arial,"Microsoft YaHei",sans-serif;
    background:#f3f4f6;
    padding:20px;
}
.box{
    max-width:1000px;
    margin:auto;
    background:white;
    padding:24px;
    border-radius:18px;
    box-shadow:0 3px 12px rgba(0,0,0,.08);
}
h1{
    text-align:center;
    font-size:36px;
}
input{
    width:100%;
    font-size:22px;
    padding:12px;
    border:1px solid #ccc;
    border-radius:10px;
}
button,.btn{
    display:inline-block;
    background:#2d7ff9;
    color:white;
    padding:12px 18px;
    border-radius:10px;
    border:0;
    font-size:20px;
    font-weight:bold;
    text-decoration:none;
    margin-top:12px;
}
table{
    width:100%;
    border-collapse:collapse;
    margin-top:20px;
}
th,td{
    border:1px solid #ddd;
    padding:10px;
    text-align:center;
}
th{
    background:#eaf2ff;
}
.edit{
    background:#16a34a;
    color:white;
    padding:7px 12px;
    border-radius:8px;
    text-decoration:none;
}
</style>
</head>
<body>
<div class="box">

    <a href="/member/member-management">← 返回会员资料管理</a>

    <h1>🔍 查找 / 编辑会员</h1>

    <form method="get">
        <input name="q" value="{{ q }}" placeholder="输入编号 / 姓名 / 英文名 / 电话">
        <button type="submit">查找会员</button>
    </form>

    {% if q %}
        <h3>搜索结果：{{ rows|length }} 位</h3>

        {% if rows %}
        <table>
            <tr>
                <th>编号</th>
                <th>姓名</th>
                <th>英文名</th>
                <th>电话</th>
                <th>PIN</th>
                <th>分会</th>
                <th>状态</th>
                <th>备注</th>
                <th>操作</th>
            </tr>

            {% for r in rows %}
            <tr>
                <td>{{ r.member_id }}</td>
                <td>{{ r.name }}</td>
                <td>{{ r.english_name or "-" }}</td>
                <td>{{ r.phone | member_phone_display }}</td>
                <td>{{ r.pin or "-" }}</td>
                <td>{{ r.branch or "-" }}</td>
                <td>{{ r.member_status or r.status or "-" }}</td>
                <td>{{ r.remark or "-" }}</td>
                <td>
                    <a class="edit" href="/member/member-management/edit/{{ r.member_id }}">
                        ✏ 编辑
                    </a>
                </td>
            </tr>
            {% endfor %}
        </table>
        {% else %}
            <p>找不到会员。</p>
        {% endif %}
    {% endif %}

</div>
</body>
</html>
""", q=q, rows=rows)

@member_bp.route("/member-management/edit/<member_id>", methods=["GET", "POST"])
def edit_member(member_id):

    if not session.get("member_admin"):
        return redirect(url_for("member.member_admin"))

    member = db_query("""
        select *
        from members
        where member_id = %s
    """, (member_id,), fetchone=True)

    if not member:
        return "会员不存在"

    msg = None
    error = None

    if request.method == "POST":

        old_data = dict(member)

        name = request.form.get("name", "").strip()
        english_name = request.form.get("english_name", "").strip()
        ic_number = request.form.get("ic_number", "").strip()
        phone = request.form.get("phone", "").strip()
        pin = request.form.get("pin", "").strip()
        branch = request.form.get("branch", "").strip().upper()
        member_status = request.form.get("member_status", "").strip()
        remark = request.form.get("remark", "").strip()

        if not name:
            error = "姓名不能为空"

        if not error and branch not in ["CHE", "STW"]:
            error = "分会只能是 CHE 或 STW"

        # 原身份证未改变时，允许更新状态等资料；跨分会历史编号可能共用身份证。
        # 若实际更换身份证号码，仍执行重复检查，防止误用他人证件。
        original_ic = (member.get("ic_number") or "").strip()
        if not error and ic_number and ic_number != original_ic:
            exist = db_query("""
                select member_id, name
                from members
                where ic_number = %s
                  and member_id <> %s
                limit 1
            """, (ic_number, member_id), fetchone=True)

            if exist:
                error = f"身份证号码已存在：{exist['member_id']} / {exist['name']}"

        if not error:
            db_query("""
                update members
                set
                    name = %s,
                    english_name = %s,
                    ic_number = %s,
                    phone = %s,
                    pin = %s,
                    branch = %s,
                    status = %s,
                    member_status = %s,
                    remark = %s
                where member_id = %s
            """, (
                name,
                english_name,
                ic_number,
                phone,
                pin,
                branch,
                member_status,
                member_status,
                remark or member_status,
                member_id
            ))

            new_member = db_query("""
                select *
                from members
                where member_id = %s
            """, (member_id,), fetchone=True)

            new_data = dict(new_member)

            if json.dumps(old_data, default=str, sort_keys=True) != json.dumps(new_data, default=str, sort_keys=True):
                db_query("""
                    insert into member_profile_history
                    (
                        member_id,
                        action,
                        old_data,
                        new_data,
                        changed_by
                    )
                    values (%s,%s,%s,%s,%s)
                """, (
                    member_id,
                    "edit",
                    json.dumps(old_data, default=str),
                    json.dumps(new_data, default=str),
                    "member_admin"
                ))

                msg = "保存成功，已记录修改历史"
            else:
                msg = "没有实际修改"

            member = new_member

    history_rows = db_query("""
        select *
        from member_profile_history
        where member_id = %s
        order by changed_at desc
        limit 10
    """, (member_id,), fetchall=True)

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>编辑会员</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body{
    font-family:Arial,"Microsoft YaHei",sans-serif;
    background:#f3f4f6;
    margin:0;
    padding:20px;
}
.box{
    max-width:900px;
    margin:auto;
    background:white;
    padding:24px;
    border-radius:18px;
    box-shadow:0 3px 12px rgba(0,0,0,.08);
}
h1{
    text-align:center;
    font-size:38px;
}
label{
    font-size:20px;
    font-weight:bold;
    display:block;
    margin-top:16px;
}
input,select,textarea{
    width:100%;
    padding:12px;
    font-size:20px;
    border:1px solid #ccc;
    border-radius:10px;
    box-sizing:border-box;
    margin-top:6px;
}
textarea{
    height:120px;
}
button{
    background:#16a34a;
    color:white;
    border:0;
    padding:14px 22px;
    border-radius:10px;
    font-size:22px;
    font-weight:bold;
    margin-top:20px;
    width:100%;
}
.ok{
    background:#e8f5e9;
    color:#176b2c;
    padding:12px;
    border-radius:8px;
    margin-bottom:15px;
}
.error{
    background:#ffe5e5;
    color:#a10000;
    padding:12px;
    border-radius:8px;
    margin-bottom:15px;
}
.member-id{
    background:#eef2ff;
    padding:14px;
    border-radius:10px;
    font-size:22px;
    font-weight:bold;
    margin-bottom:18px;
}
table{
    width:100%;
    border-collapse:collapse;
    margin-top:15px;
}
th,td{
    border:1px solid #ddd;
    padding:10px;
    text-align:center;
}
th{
    background:#eeeeee;
}
.note{
    background:#fff7d6;
    padding:12px;
    border-radius:10px;
    margin-top:15px;
}
</style>
</head>
<body>

<div class="box">

    <p>
        <a href="/member/member-management/search">← 返回搜索</a>
        &nbsp; | &nbsp;
        <a href="/member/member-management">返回会员资料管理</a>
    </p>

    <h1>✏ 编辑会员</h1>

    {% if error %}
    <div class="error">❌ {{ error }}</div>
    {% endif %}

    {% if msg %}
    <div class="ok">✅ {{ msg }}</div>
    {% endif %}

    <div class="member-id">
        会员编号：{{ member.member_id }}
    </div>

    <form method="post">

        <label>姓名</label>
        <input name="name" value="{{ member.name or '' }}" required>

        <label>英文名</label>
        <input name="english_name" value="{{ member.english_name or '' }}">

        <label>身份证号码</label>
        <input name="ic_number"
               value="{{ member.ic_number or '' }}"
               placeholder="例如：800101-14-1234">

        <label>电话</label>
        <input name="phone" value="{{ member.phone or '' }}">

        <label>PIN</label>
        <input name="pin" value="{{ member.pin or '' }}">

        <label>分会</label>
        <select name="branch">
            <option value="CHE" {% if member.branch=="CHE" %}selected{% endif %}>CHE</option>
            <option value="STW" {% if member.branch=="STW" %}selected{% endif %}>STW</option>
        </select>

        <label>状态</label>
        <select name="member_status">
            <option value="在供" {% if member.member_status=="在供" or member.status=="在供" %}selected{% endif %}>在供</option>
            <option value="停供" {% if member.member_status=="停供" or member.status=="停供" %}selected{% endif %}>停供</option>
            <option value="暂停" {% if member.member_status=="暂停" or member.status=="暂停" %}selected{% endif %}>暂停</option>
            <option value="已往生" {% if member.member_status=="已往生" or member.status=="已往生" %}selected{% endif %}>已往生</option>
        </select>

        <label>备注</label>
        <textarea name="remark">{{ member.remark or '' }}</textarea>

        <button type="submit">
            💾 保存修改
        </button>

    </form>

    <div class="note">
        说明：会员编号不可修改；若需要转分会，之后会另外做「转分会」功能。
    </div>

    <h2>📜 最近会员资料修改历史</h2>

    {% if history_rows %}
    <table>
        <tr>
            <th>时间</th>
            <th>动作</th>
            <th>修改者</th>
        </tr>
        {% for h in history_rows %}
        <tr>
            <td>{{ (h.changed_at + timedelta(hours=8)).strftime("%Y-%m-%d %H:%M") if h.changed_at else "-" }}</td>
            <td>{{ h.action }}</td>
            <td>{{ h.changed_by }}</td>
        </tr>
        {% endfor %}
    </table>
    {% else %}
        <div class="note">暂时没有会员资料修改历史。</div>
    {% endif %}

</div>

</body>
</html>
""",
member=member,
msg=msg,
error=error,
history_rows=history_rows,
timedelta=timedelta
)


@member_bp.route("/member-management/add", methods=["GET", "POST"])
def add_member():

    if not session.get("member_admin"):
        return redirect(url_for("member.member_admin"))

    msg = None
    error = None

    if request.method == "POST":
        branch = request.form.get("branch", "CHE").strip().upper()
        name = request.form.get("name", "").strip()
        english_name = request.form.get("english_name", "").strip()
        phone = request.form.get("phone", "").strip()
        remark = request.form.get("remark", "").strip()
        ic_number = request.form.get("ic_number", "").strip()

        if branch not in ["CHE", "STW"]:
            error = "分会不正确"
        elif not name:
            error = "姓名不能为空"
        else:
            last_no = db_query("""
                select max(
                    cast(regexp_replace(member_id, '^(CHE-|STW-)', '') as integer)
                ) as last_no
                from members
                where member_id like %s
            """, (branch + "-%",), fetchone=True)

            next_no = int(last_no["last_no"] or 0) + 1
            member_id = f"{branch}-{next_no}"

            phone_digits = "".join(ch for ch in phone if ch.isdigit())
            pin = phone_digits[-4:] if len(phone_digits) >= 4 else "0000"

            exist = db_query("""
                select member_id,name
                from members
                where ic_number=%s
            """, (ic_number,), fetchone=True)

            if exist:
                error = f"""
                身份证号码已存在：
                {exist['member_id']}
                {exist['name']}
                """

            db_query("""
                insert into members
                (
                    member_id,
                    name,
                    english_name,
                    ic_number,
                    phone,
                    pin,
                    branch,
                    status,
                    remark,
                    member_status
                )
                values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                member_id,
                name,
                english_name,
                ic_number,
                phone,
                pin,
                branch,
                "在供",
                remark or "在供",
                "在供"
            ))

            msg = f"新增成功：{member_id} / {name} / PIN：{pin}"

    return render_template_string("""
<div style="max-width:900px;margin:auto;padding:20px;">

    <p>
        <a href="/member/member-management">← 返回会员资料管理</a>
    </p>

    <h1>➕ 新增会员</h1>

    {% if error %}
    <div style="background:#ffe5e5;color:#a10000;padding:12px;border-radius:8px;">
        ❌ {{ error }}
    </div>
    {% endif %}

    {% if msg %}
    <div style="background:#e8f5e9;color:#176b2c;padding:12px;border-radius:8px;">
        ✅ {{ msg }}
    </div>
    {% endif %}

    <form method="post">

        <label>分会</label><br>
        <select name="branch" style="width:100%;padding:12px;font-size:20px;">
            <option value="CHE">CHE</option>
            <option value="STW">STW</option>
        </select><br><br>

        <label>姓名</label><br>
        <input name="name" required style="width:100%;padding:12px;font-size:20px;"><br><br>

        <label>英文名</label><br>
        <input name="english_name" style="width:100%;padding:12px;font-size:20px;"><br><br>
                                  
        <label>身份证号码</label>
        <input
            name="ic_number"
            placeholder="例如：800101-14-1234"
            style="width:100%;padding:12px;font-size:20px;">

        <br><br>

        <label>电话</label><br>
        <input name="phone" style="width:100%;padding:12px;font-size:20px;"><br><br>

        <label>备注</label><br>
        <textarea name="remark" style="width:100%;height:100px;padding:12px;font-size:20px;"></textarea><br><br>

        <button type="submit"
            style="background:#2d7ff9;color:white;border:0;padding:14px 22px;border-radius:10px;font-size:22px;font-weight:bold;">
            新增会员
        </button>

    </form>
</div>
""", msg=msg, error=error)


# 2025 / 2026 财政原始 Excel 自动识别与标准化
import re
import calendar

def _fee_month_range(value, payment_date=None):
    if value is None or pd.isna(value):
        raise ValueError("缺少供养月份")
    if isinstance(value, (datetime, date, pd.Timestamp)):
        # Excel 把 Aug26 误识别为 2026-08-26 时，以月份为准
        d = pd.Timestamp(value).date()
        return date(d.year, d.month, 1), date(d.year, d.month, 1)
    t = str(value).upper().strip().replace("’", "'").replace("_", "-")
    t = re.sub(r"^MEMBER\s+FEE\s+FOR\s+", "", t)
    t = re.sub(r"\s+", "", t).replace("SEPT", "SEP").replace("JULY", "JUL")
    t = t.replace("'", "")
    if not t or t == '-':
        raise ValueError("供养月份为空")
    # 日期文字，例如 2026-08-26（原始 Excel 自动转换）
    if re.fullmatch(r"20\d{2}-\d{1,2}-\d{1,2}.*", t):
        d = pd.to_datetime(t).date()
        return date(d.year,d.month,1),date(d.year,d.month,1)
    # 例如 Sept-26 / Oct-26：连字符后面只有年份，不是第二个月份
    # 统一转换成 SEP26 / OCT26，避免把 26 当成月份解析。
    if re.fullmatch(r"[A-Z]{3}-\d{2,4}", t):
        t = t.replace("-", "")
    parts = t.split('-')
    if len(parts) > 2:
        raise ValueError(f"无法识别供养月份：{value}")
    def token(x):
        m = re.fullmatch(r"([A-Z]{3})(\d{2,4})?", x)
        if not m:
            raise ValueError(f"无法识别月份：{x}")
        mon = list(calendar.month_abbr).index(m.group(1).title())
        if not mon:
            raise ValueError(f"月份无效：{x}")
        yr = m.group(2)
        return mon, (2000 + int(yr) if len(yr)==2 else int(yr)) if yr else None
    a,ya = token(parts[0]); b,yb = token(parts[-1])
    if ya is None and yb is None:
        raise ValueError(f"没有年份：{value}")
    if ya is None:
        ya = yb - (1 if a > b else 0)
    if yb is None:
        yb = ya + (1 if b < a else 0)
    start,end = date(ya,a,1),date(yb,b,1)
    if start > end:
        raise ValueError(f"月份起止倒置：{value}")
    return start,end

def _fee_amount_number(value):
    """Excel 金额栏：空白、横线代表 0；支持 RM 和千位逗号。"""
    if value is None or pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text in ("", "-", "–", "—"):
        return 0.0
    text = re.sub(r"(?i)RM\s*", "", text).replace(",", "").strip()
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    return float(text)

def _upload_payment_frames(upload):
    """读取财政收据表：保留 2025 旧格式，动态识别 2026+ CHE/STW sort E。

    仅返回待核对资料，不执行数据库写入。Excel 物理行号完整保留。
    """
    import io
    import unicodedata

    # 显式指定 openpyxl，支持 .xlsx/.xlsm；使用同一份字节避免流位置问题。
    upload.seek(0)
    book = pd.ExcelFile(io.BytesIO(upload.read()), engine="openpyxl")
    frames = []
    for sheet in book.sheet_names:
        key = unicodedata.normalize("NFKC", sheet).strip().upper()
        if key == "DATA FOR YR2025 CHERAS":
            raw = pd.read_excel(book, sheet_name=sheet, header=0)
            for i, r in raw.iterrows():
                no = r.iloc[2]
                if pd.isna(no):
                    continue
                frames.append(dict(
                    sheet=sheet, excel_row=i + 2, branch="CHE", number=no,
                    name=r.iloc[3], receipt=r.iloc[1], payment_date=r.iloc[0],
                    start=r.iloc[4], end=r.iloc[5], months=r.iloc[6], amount=r.iloc[7],
                ))
            continue

        # 支持 2027（CHE) sort E / 2027 (CHE) sort E / 2027 STW (sort E)
        # 只认 sort E，避免误读会员名单、审计表及其它财务工作表。
        match = re.match(r"^(20\d{2})\s*(?:\(\s*(CHE|STW)\s*\)|(CHE|STW))\s*\(?\s*SORT\s*E\s*\)?$", key)
        if not match:
            continue
        branch = match.group(2) or match.group(3)
        raw = pd.read_excel(book, sheet_name=sheet, header=None)

        # 找到实际列标题，而不是假设表头永远位于 Excel 第 4 行。
        header_index = None
        for ri in range(min(len(raw), 20)):
            values = [str(v).strip().upper() if not pd.isna(v) else "" for v in raw.iloc[ri].tolist()]
            if "收据日期" in values and "CASH" in values and "BANK" in values:
                header_index = ri
                break
        if header_index is None:
            raise ValueError(f"{sheet}：找不到收据日期、CASH、BANK 表头，请检查格式")

        def col_value(values, index):
            return values[index] if index < len(values) else None

        for ri in range(header_index + 1, len(raw)):
            v = raw.iloc[ri].tolist()
            excel_row = ri + 1
            # CHE: C收据 D日期 E编号 F姓名 G月份 H现金 I银行
            # STW: C收据 D日期 E编号 G姓名 I月份 J现金 K银行
            if branch == "CHE":
                receipt, pdate, no, name, period, cash, bank = (
                    col_value(v, idx) for idx in (2, 3, 4, 5, 6, 7, 8)
                )
            else:
                receipt, pdate, no, name, period, cash, bank = (
                    col_value(v, idx) for idx in (2, 3, 4, 6, 8, 9, 10)
                )

            # 月份分段标题、合计行及空白行都不算付款记录。
            if pd.isna(no) or pd.isna(period):
                continue
            if pd.isna(cash) and pd.isna(bank):
                continue
            try:
                amount = _fee_amount_number(cash) + _fee_amount_number(bank)
            except (ValueError, TypeError) as exc:
                raise ValueError(
                    f"{sheet} 第{excel_row}行金额无法识别：CASH={cash!r}，BANK={bank!r}"
                ) from exc
            if amount <= 0:
                continue
            frames.append(dict(
                sheet=sheet, excel_row=excel_row, branch=branch, number=no,
                name=name, receipt=receipt, payment_date=pdate,
                period=period, amount=amount,
            ))
    if not frames:
        raise ValueError("没有找到可导入的付款记录工作表（2025 CHE 或 2026 年起 CHE/STW sort E）")
    return frames

def _finance_report_excel(details, summary, count):
    """对账明细，保留每一行来源、财政数据及数据库旧值。"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = '对账总览'
    ws.append(['财政 Excel 对账报告（预览，不代表已更新）'])
    ws.append(['Excel 读取行数', count])
    for key, value in summary.items():
        ws.append([key, value])
    ws.append(['说明', '收据相同但内容不同、月份重叠等冲突均不自动覆盖；重复收据仅完全相同才折叠。'])
    cols = ['结果','Excel 工作表及行号','会员编号','Excel 姓名','主档姓名','收据号码',
            'Excel 日期','Excel 起始月份','Excel 结束月份','Excel 月数','Excel 金额',
            '数据库 ID','数据库日期','数据库起始月份','数据库结束月份','数据库月数','数据库金额','说明']
    grouped = {}
    for d in details:
        status = d.get('结果','异常')
        grouped.setdefault(status, []).append(d)
    for status, items in grouped.items():
        name = str(status).replace('/','_').replace('\\','_')[:31]
        sheet = wb.create_sheet(name)
        sheet.append(cols)
        for d in items:
            sheet.append([str(d.get(c, '') or '') for c in cols])
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = f'A1:R{sheet.max_row}'
    for sheet in wb:
        sheet.sheet_view.showGridLines = False
        sheet.row_dimensions[1].height = 28
        sheet.freeze_panes = sheet.freeze_panes or 'A2'
        for cell in sheet[1]:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='195A9D')
            cell.alignment = Alignment(vertical='center', wrap_text=True)
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.fill = PatternFill('solid', fgColor=('EAF3FD' if cell.row % 2 == 0 else 'FFFFFF'))
                cell.alignment = Alignment(vertical='center')
        for col in sheet.columns:
            letter = get_column_letter(col[0].column)
            maxlen = max((len(str(c.value or '')) for c in list(col)[:150]), default=12)
            sheet.column_dimensions[letter].width = min(max(maxlen+3, 15), 46)
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# 财政专用入口：独立 PIN，不授予月费管理员权限。
@member_bp.route("/finance-upload/login", methods=["GET", "POST"])
def finance_upload_login():
    error = ""
    if request.method == "POST":
        expected = os.environ.get("FINANCE_UPLOAD_PIN", "")
        supplied = request.form.get("pin", "")
        if expected and hmac.compare_digest(supplied, expected):
            session["finance_uploader"] = True
            return redirect(url_for("member.finance_upload"))
        error = "财政上传 PIN 不正确，或尚未配置 FINANCE_UPLOAD_PIN"
    return render_template_string("""<!doctype html><html lang="zh"><meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>财政 Excel 上传登录</title><body style="font-family:Arial,sans-serif;max-width:440px;margin:65px auto;padding:20px">
    <h2>📒 财政 Excel 上传</h2><p>使用财政专用 PIN 登录（不能进入管理员管理功能）。</p>
    {% if error %}<p style="color:#b91c1c">{{error}}</p>{% endif %}
    <form method="post"><input name="pin" type="password" autocomplete="off" required
      style="padding:12px;width:95%;font-size:18px" placeholder="财政上传 PIN">
      <button style="padding:12px;margin-top:12px;width:100%;font-size:18px">登录</button></form>
    </body></html>""", error=error)

@member_bp.route("/finance-upload/logout", methods=["POST"])
def finance_upload_logout():
    session.pop("finance_uploader", None)
    return redirect(url_for("member.finance_upload_login"))

@member_bp.route("/finance-upload", methods=["GET", "POST"])
def finance_upload():

    is_admin = bool(session.get("member_admin"))
    is_finance = bool(session.get("finance_uploader"))
    if not (is_admin or is_finance):
        return redirect(url_for("member.finance_upload_login"))

    error = None
    msg = None
    rows = []
    history_rows = []
    safety_issues = []
    issue_report = ""

    q = request.args.get("q", "").strip()
    branch = request.args.get("branch", "CHE").strip().upper()
    year = request.args.get("year", str(date.today().year)).strip()
    months = request.args.getlist("months")
    
    # 下载 Excel
    if request.args.get("download") == "1":
        if not is_admin:
            return redirect(url_for("member.member_admin"))
        try:
            params = []
            where = "where 1=1"

            if year:
                where += " and to_char(payment_date, 'YYYY') = %s"
                params.append(year)

            if months:
                where += " and to_char(payment_date, 'MM') = any(%s)"
                params.append(months)

            df = pd.read_sql_query(f"""
                select
                    payment_date as "日期\nDate",
                    receipt_no as "收据编号 \nOfficial Receipt No",
                    regexp_replace(member_id, '^(CHE-|STW-)', '') as "编号 No/",
                    name as "捐款人\n姓名\nName",
                    start_month as "START MONTH",
                    end_month as "END MONTH",
                    month_count as "No/ of Mth",
                    amount as "Total Amt"
                from member_payments
                {where}
                order by payment_date asc, id asc
            """, get_conn(), params=params)

            output = BytesIO()

            with pd.ExcelWriter(output, engine="openpyxl") as writer:

                df.to_excel(
                    writer,
                    index=False,
                    sheet_name="月费记录",
                    startrow=5
                )

                ws = writer.book["月费记录"]

                # ======================
                # 标题
                # ======================

                title = f"{year}年{'全部月份' if months == 'all' else '、'.join(months) + '月'}月费记录"

                ws.merge_cells("A1:H1")
                ws["A1"] = title

                ws["A1"].font = Font(
                    size=18,
                    bold=True,
                    color="FFFFFF"
                )

                ws["A1"].fill = PatternFill(
                    "solid",
                    fgColor="4472C4"
                )

                ws["A1"].alignment = Alignment(horizontal="center")

                # ======================
                # 摘要
                # ======================

                total_people = len(df)
                total_months = int(df["No/ of Mth"].sum())
                total_amount = float(df["Total Amt"].sum())

                # ======================
                # 摘要卡片
                # ======================

                summary_fill = PatternFill("solid", fgColor="EAF4FF")
                summary_border = Border(
                    left=Side(style="thin", color="9EADCC"),
                    right=Side(style="thin", color="9EADCC"),
                    top=Side(style="thin", color="9EADCC"),
                    bottom=Side(style="thin", color="9EADCC")
                )

                summary_items = [
                    ("A3:B4", "总人数", f"{total_people} 人"),
                    ("D3:E4", "总月数", f"{total_months}"),
                    ("G3:H4", "总金额", f"RM {total_amount:,.2f}"),
                ]

                for cell_range, label, value in summary_items:
                    ws.merge_cells(cell_range)
                    cell = ws[cell_range.split(":")[0]]
                    cell.value = f"{label}\n{value}"
                    cell.fill = summary_fill
                    cell.border = summary_border
                    cell.font = Font(bold=True, size=12)
                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=True
                    )

                ws.row_dimensions[3].height = 28
                ws.row_dimensions[4].height = 28

                # ======================
                # 表头美化
                # ======================

                header_fill = PatternFill("solid", fgColor="0F5F76")
                header_font = Font(color="FFFFFF", bold=True)

                row_fill = PatternFill("solid", fgColor="BFE8F5")

                for cell in ws[6]:

                    cell.font = Font(
                        bold=True
                    )

                    cell.fill = header_fill

                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=True
                    )

                # ======================
                # 边框样式
                # ======================

                thin = Side(
                    border_style="thin",
                    color="CCCCCC"
                )

                border = Border(
                    left=thin,
                    right=thin,
                    top=thin,
                    bottom=thin
                )

                # ======================
                # 自动栏宽
                # ======================

                for col in ws.columns:

                    max_len = 0

                    letter = get_column_letter(col[0].column)

                    for cell in col:

                        try:
                            max_len = max(
                                max_len,
                                len(str(cell.value))
                            )
                        except:
                            pass

                    ws.column_dimensions[letter].width = min(max_len + 4, 30)
                
                ws.freeze_panes = "A7"
                
                blue_fill = PatternFill(
                    "solid",
                    fgColor="D8EEF8"
                )

                white_fill = PatternFill(
                    "solid",
                    fgColor="FFFFFF"
                )

                for row_num in range(7, ws.max_row + 1):

                    fill = blue_fill if row_num % 2 == 0 else white_fill

                    for col in range(1, 9):

                        cell = ws.cell(row=row_num, column=col)

                        cell.fill = fill
                        cell.border = border

                ws.auto_filter.ref = f"A6:H{ws.max_row}"

                # ======================
                # 金额格式
                # ======================

                for row in range(7, ws.max_row + 1):

                    ws[f"A{row}"].number_format = "dd/mm/yyyy"

                    ws[f"H{row}"].number_format = '#,##0.00'

                    ws[f"H{row}"].alignment = Alignment(
                        horizontal="right"
                    )

            output.seek(0)

            month_text = "全部" if not months else "_".join(months)
            filename = f"{year}_{month_text}_月费记录.xlsx"

            return send_file(
                output,
                as_attachment=True,
                download_name=filename,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            error = f"下载失败：{e}"
   
    # 财政 Excel 对账：一次读取数据库；先预览，勾选确认才写入。
    if request.method == "POST":
        upload = request.files.get("file")
        confirmed = request.form.get("confirm_finance") == "yes"
        confirm_dates = request.form.get("confirm_dates_only") == "yes"
        if confirmed and confirm_dates:
            confirmed = False  # 日期专用按钮优先，不允许混合执行
        if is_finance and not is_admin and confirm_dates:
            confirm_dates = False
            error = "财政账号无权修正付款日期，请通知管理员"
        # 财政账号只能新增安全付款；管理员保留 V8 的原始处理方式。
        if error:
            pass
        elif not upload or not upload.filename:
            error = "请选择 Excel 文件"
        else:
            try:
                from decimal import Decimal
                records = _upload_payment_frames(upload)
                prepared, errors = [], []
                details = []
                receipt_groups = {}
                def report_line(item, status, note='', old=None, master_name=''):
                    old = old or {}
                    def fmt(v):
                        return '' if v is None else str(v)
                    details.append({
                        '结果':status, 'Excel 工作表及行号':item.get('source',''),
                        '会员编号':item.get('member_id',''), 'Excel 姓名':item.get('excel_name',''),
                        '主档姓名':master_name, '收据号码':item.get('receipt_no',''),
                        'Excel 日期':fmt(item.get('payment_date')),
                        'Excel 起始月份':fmt(item.get('start_month')),
                        'Excel 结束月份':fmt(item.get('end_month')),
                        'Excel 月数':fmt(item.get('month_count')),
                        'Excel 金额':fmt(item.get('amount')),
                        '数据库 ID':fmt(old.get('id')),
                        '数据库日期':fmt(old.get('payment_date')),
                        '数据库起始月份':fmt(old.get('start_month')),
                        '数据库结束月份':fmt(old.get('end_month')),
                        '数据库月数':fmt(old.get('month_count')),
                        '数据库金额':fmt(old.get('amount')), '说明':note
                    })
                for raw_item in records:
                    try:
                        number = str(raw_item['number']).strip().upper()
                        if number.endswith('.0'): number = number[:-2]
                        if number.startswith(('CHE-', 'STW-')):
                            member_id = number
                        elif number.isdigit():
                            member_id = f"{raw_item['branch']}-{int(number)}"
                        else:
                            raise ValueError(f"无效会员编号 {number}")
                        receipt = raw_item.get('receipt')
                        receipt = None if pd.isna(receipt) else str(receipt).strip().replace(' ', '').upper()
                        if receipt in ('', 'NAN', 'NONE', 'CHE', 'STW'): receipt = None  # 分会占位文字不是正式收据号码
                        if pd.isna(raw_item['payment_date']): raise ValueError('付款日期为空')
                        payment_date = pd.to_datetime(raw_item['payment_date']).date()
                        if 'period' in raw_item:
                            start, end = _fee_month_range(raw_item['period'], payment_date)
                        else:
                            start, end = parse_month(raw_item['start']), parse_month(raw_item['end'])
                        if not start or not end or start > end: raise ValueError('月份无效')
                        count = (end.year-start.year)*12+end.month-start.month+1
                        if 'months' in raw_item and int(raw_item['months']) != count:
                            raise ValueError('供养月数与起止月份不符')
                        amount = Decimal(str(raw_item['amount'])).quantize(Decimal('0.01'))
                        if amount <= 0: raise ValueError('金额必须大于零')
                        item = dict(member_id=member_id, receipt_no=receipt,
                            payment_date=payment_date, start_month=start, end_month=end,
                            month_count=count, amount=amount,
                            excel_name=(str(raw_item['name']).strip() if not pd.isna(raw_item['name']) else ''),
                            source=f"{raw_item['sheet']} 第{raw_item['excel_row']}行")
                        prepared.append(item)
                        if receipt:
                            receipt_groups.setdefault(receipt, []).append(item)
                    except Exception as exc:
                        msg_err = f"{raw_item['sheet']} 第{raw_item['excel_row']}行：{exc}"
                        errors.append(msg_err)
                        report_line({'source':f"{raw_item['sheet']} 第{raw_item['excel_row']}行"}, '格式错误', str(exc))
                # 同收据若出现多次，先对整组比较；内容不同则整组隔离，不能让第一笔先更新。
                blocked_receipts = set()
                duplicate_excel = set()
                for receipt, group in receipt_groups.items():
                    if len(group) < 2:
                        continue
                    first = group[0]
                    keys = ('member_id','payment_date','start_month','end_month','month_count','amount')
                    if all(all(it[k] == first[k] for k in keys) for it in group[1:]):
                        for it in group[1:]:
                            duplicate_excel.add(id(it))
                            report_line(it, 'Excel完全重复', f"与 {first['source']} 完全相同，已折叠")
                    else:
                        blocked_receipts.add(receipt)
                        for it in group:
                            errors.append(f"{it['source']}：收据 {receipt} 在 Excel 内对应不同付款资料，整组待核对")
                            report_line(it, 'Excel收据冲突', f"同收据共 {len(group)} 行，内容不一致，整组禁止自动写入")
                # 无收据/跨表重复：以编号+日期+月份+月数+金额折叠，避免重复匹配。
                signature_seen = {}
                for it in prepared:
                    if id(it) in duplicate_excel or it['receipt_no'] in blocked_receipts:
                        continue
                    signature = (it['member_id'], it['payment_date'], it['start_month'],
                                 it['end_month'], it['month_count'], it['amount'])
                    previous = signature_seen.get(signature)
                    if previous is not None:
                        # 不同收据号码不能自动合并；无收据与有收据可视为跨表重复。
                        if not it['receipt_no'] or not previous['receipt_no'] or it['receipt_no'] == previous['receipt_no']:
                            duplicate_excel.add(id(it))
                            report_line(it, 'Excel完全重复', f"与 {previous['source']} 付款内容相同，跨表折叠")
                            continue
                    else:
                        signature_seen[signature] = it
                summary = dict(一致=0, 新增=0, 更新=0, 日期待更新=0, 冲突=0, 会员不存在=0)
                changes = []
                with get_conn() as conn:
                    with conn.cursor(cursor_factory=RealDictCursor) as cur:
                        # 固定读取两张表，避免 1000 行 × 多次远程查询。
                        cur.execute('SELECT member_id,name FROM members')
                        members = {r['member_id']: r for r in cur.fetchall()}
                        cur.execute('SELECT id,member_id,receipt_no,payment_date,start_month,end_month,month_count,amount,name FROM member_payments')
                        existing = cur.fetchall()
                        receipts = {}
                        by_member = {}
                        for row in existing:
                            if row['receipt_no']:
                                receipts.setdefault(str(row['receipt_no']).strip().replace(' ', ''), []).append(row)
                            by_member.setdefault(row['member_id'], []).append(row)
                        used_ids = set()
                        planned_periods = []
                        def same_core(old, item):
                            return (old['start_month'] == item['start_month']
                                and old['end_month'] == item['end_month']
                                and int(old['month_count'] or 0) == item['month_count']
                                and Decimal(str(old['amount'])).quantize(Decimal('0.01')) == item['amount'])
                        def same_full(old, item):
                            return same_core(old, item) and old['payment_date'] == item['payment_date']
                        def planned_conflict(item):
                            return any(m == item['member_id'] and a <= item['end_month']
                                and b >= item['start_month'] for m,a,b in planned_periods)
                        for item in prepared:
                            if id(item) in duplicate_excel or item['receipt_no'] in blocked_receipts:
                                continue
                            mid, receipt = item['member_id'], item['receipt_no']
                            member = members.get(mid)
                            if not member:
                                summary['会员不存在'] += 1
                                errors.append(f"{item['source']}：会员 {mid} 不在 members 主档")
                                report_line(item, '会员不存在', '会员编号未在主档找到')
                                continue
                            name = (member['name'] or '').strip()
                            # 财政月费对账以会员编号为准；姓名差异不提醒、不拦截，
                            # 同时不以 Excel 姓名覆盖 members 主档。
                            item['name'] = name
                            matches = receipts.get(receipt, []) if receipt else []
                            if len(matches) > 1 or (matches and matches[0]['member_id'] != mid):
                                summary['冲突'] += 1
                                errors.append(f"{item['source']}：收据 {receipt} 已属于其他会员或有重复，需人工检查")
                                report_line(item, '数据库收据冲突', '数据库同收据对应多个记录或其他会员', matches[0] if matches else None, name)
                                continue
                            old = matches[0] if matches else None
                            if not old:
                                overlaps = [r for r in by_member.get(mid, [])
                                    if r['start_month'] <= item['end_month'] and r['end_month'] >= item['start_month']]
                                # 优先匹配完全相同的无收据记录（含 Excel 有收据、数据库无收据的情形）。
                                # 若同一会员/月份对应多个候选记录，禁止猜测匹配。
                                exact = [r for r in overlaps if same_full(r, item)]
                                if len(exact) == 1 and exact[0]['id'] not in used_ids:
                                    old = exact[0]
                                    used_ids.add(old['id'])
                                    if receipt and not old['receipt_no']:
                                        # 只补唯一确定的原付款收据，不新增第二笔付款。
                                        summary['更新'] += 1
                                        changes.append(('receipt_update', item, old))
                                        report_line(item, '更新', '付款内容完全相同，补上正式收据号码', old, name)
                                    else:
                                        summary['一致'] += 1
                                        report_line(item, '一致', '按编号、日期、月份、月数、金额匹配；不改动收据', old, name)
                                    continue
                                if len(exact) > 1:
                                    summary['冲突'] += 1
                                    errors.append(f"{item['source']}：数据库存在多笔相同付款，不能确定对应记录")
                                    report_line(item, '数据库重复付款', '多笔记录完全相同，禁止自动匹配', exact[0], name)
                                    continue
                                core = [r for r in overlaps if same_core(r, item)]
                                if (len(core) == 1 and core[0]['id'] not in used_ids
                                    and (not receipt and not core[0]['receipt_no'])):
                                    used_ids.add(core[0]['id'])
                                    summary['日期待更新'] += 1
                                    changes.append(('date_update', item, core[0]))
                                    report_line(item, '日期待更新', '编号、月份、月数、金额相同，仅付款日期不同；确认后以财政 Excel 日期更新，保留原收据', core[0], name)
                                    continue
                                if overlaps or planned_conflict(item):
                                    summary['冲突'] += 1
                                    errors.append(f"{item['source']}：{mid} 供养月份与现有或本次待写记录重叠，未自动覆盖")
                                    report_line(item, '月份冲突', '同会员供养月份重叠，人工核对', overlaps[0] if overlaps else None, name)
                                    continue
                                summary['新增'] += 1
                                changes.append(('insert', item, None))
                                planned_periods.append((mid,item['start_month'],item['end_month']))
                                report_line(item, '新增', '数据库无对应记录且月份不重叠', None, name)
                            else:
                                if old['id'] in used_ids:
                                    summary['冲突'] += 1
                                    errors.append(f"{item['source']}：同一数据库记录被 Excel 多次匹配")
                                    report_line(item, '数据库重复匹配', '数据库同一记录被多次匹配', old, name)
                                    continue
                                used_ids.add(old['id'])
                                if same_full(old, item):
                                    summary['一致'] += 1
                                    report_line(item, '一致', '同收据、日期、月份、月数及金额一致', old, name)
                                else:
                                    other_overlaps = [r for r in by_member.get(mid, []) if r['id'] != old['id']
                                        and r['start_month'] <= item['end_month'] and r['end_month'] >= item['start_month']]
                                    if other_overlaps or planned_conflict(item):
                                        summary['冲突'] += 1
                                        errors.append(f"{item['source']}：修改后月份会与其他付款重叠")
                                        report_line(item, '修改月份冲突', '修改后月份与其他数据库或待写付款重叠', other_overlaps[0] if other_overlaps else None, name)
                                    elif same_core(old, item):
                                        if mid.startswith('STW-'):
                                            # STW Excel 日期为开收据日期，不覆盖数据库原有收款日期。
                                            summary['一致'] += 1
                                            report_line(item, '一致', 'STW 收据日期与数据库收款日期不同；保留数据库原日期', old, name)
                                        else:
                                            summary['日期待更新'] += 1
                                            changes.append(('date_update', item, old))
                                            report_line(item, '日期待更新', '同收据、月份、月数、金额相同，仅日期不同；确认后以财政 Excel 日期更新', old, name)
                                    else:
                                        summary['更新'] += 1
                                        changes.append(('update', item, old))
                                        planned_periods.append((mid,item['start_month'],item['end_month']))
                                        report_line(item, '更新', '同收据同会员，金额或供养月份不同', old, name)
                        if confirmed and is_admin and (summary['冲突'] or blocked_receipts or summary['会员不存在']):
                            raise ValueError('仍有收据冲突、月份冲突或会员不存在；为避免部分写入，已禁止确认。请下载对账明细核对后再操作')
                        if confirm_dates:
                            # 独立的安全日期修正：冲突保留不动，只修改唯一确定的日期差异。
                            # 不触及金额、月份、收据；若有重复目标或不唯一的收据则不写。
                            date_changes = [c for c in changes if c[0] == 'date_update']
                            ids = [c[2]['id'] for c in date_changes]
                            if len(ids) != len(set(ids)):
                                raise ValueError('日期更新目标重复，本次未修改数据库')
                            for _, item, old in date_changes:
                                receipt = item.get('receipt_no')
                                old_receipt = str(old.get('receipt_no') or '').strip().replace(' ', '').upper() or None
                                if receipt != old_receipt:
                                    raise ValueError(f"{item['source']}：日期更新的收据不一致，已停止本次更新")
                        if confirmed or confirm_dates:
                            # 修改前备份整行旧值至独立审计表，避免覆盖后无法追溯。
                            cur.execute('''CREATE TABLE IF NOT EXISTS member_finance_upload_audit (
                                id BIGSERIAL PRIMARY KEY, changed_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                                operator TEXT, source TEXT, action TEXT NOT NULL,
                                payment_id BIGINT, old_data JSONB, new_data JSONB NOT NULL)''')
                            import json as _json
                            selected_changes = ([c for c in changes if c[0] == 'date_update']
                                                if confirm_dates else changes)
                            if is_finance and not is_admin:
                                # 财政不能改动已有付款；即使 Excel 有冲突，也仅导入独立的安全新增。
                                selected_changes = [c for c in selected_changes if c[0] == 'insert']
                            for action, item, old in selected_changes:
                                values = (item['receipt_no'], item['member_id'], item['name'], item['payment_date'],
                                    item['start_month'], item['end_month'], item['month_count'], item['amount'])
                                if action == 'receipt_update':
                                    payment_id = old['id']
                                    cur.execute("UPDATE member_payments SET receipt_no=%s WHERE id=%s AND (receipt_no IS NULL OR trim(receipt_no)='')",
                                        (item['receipt_no'], payment_id))
                                    if cur.rowcount != 1:
                                        raise ValueError(f"{item['source']}：原付款收据已改变，停止写入")
                                elif action == 'date_update':
                                    payment_id = old['id']
                                    cur.execute('UPDATE member_payments SET payment_date=%s WHERE id=%s',
                                        (item['payment_date'], payment_id))
                                elif action == 'insert':
                                    cur.execute('''INSERT INTO member_payments
                                        (receipt_no,member_id,name,payment_date,start_month,end_month,month_count,amount)
                                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id''', values)
                                    payment_id = cur.fetchone()['id']
                                else:
                                    payment_id = old['id']
                                    cur.execute('''UPDATE member_payments SET receipt_no=%s,member_id=%s,name=%s,
                                        payment_date=%s,start_month=%s,end_month=%s,month_count=%s,amount=%s
                                        WHERE id=%s''', values+(payment_id,))
                                old_json = _json.dumps(dict(old), default=str, ensure_ascii=False) if old else None
                                new_json = _json.dumps({k:str(v) for k,v in item.items()}, ensure_ascii=False)
                                cur.execute('''INSERT INTO member_finance_upload_audit
                                    (operator,source,action,payment_id,old_data,new_data)
                                    VALUES (%s,%s,%s,%s,%s::jsonb,%s::jsonb)''',
                                    ("管理员" if is_admin else "财政上传账号", item['source'], action, payment_id, old_json, new_json))
                            conn.commit()
                if request.form.get('download_report') == 'yes' and not confirmed:
                    report_file = _finance_report_excel(details, summary, len(records))
                    return send_file(report_file, as_attachment=True,
                        download_name='财政月费对账明细.xlsx',
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                summary_text = '，'.join(f'{k} {v}' for k,v in summary.items())
                if confirm_dates:
                    msg = (f"仅日期更新完成：成功修正 {len([c for c in changes if c[0] == 'date_update'])} 笔；"
                           f"其余金额差异、月份冲突、收据冲突均未修改。读取 {len(records)} 行；{summary_text}。")
                elif confirmed:
                    msg = f"财政对账写入完成：读取 {len(records)} 行；{summary_text}；格式错误 {len(errors)}（含冲突）。"
                else:
                    msg = f"财政对账预览（尚未修改数据库）：读取 {len(records)} 行；{summary_text}；Excel完全重复 {len(duplicate_excel)}，Excel冲突收据组 {len(blocked_receipts)}，异常 {len(errors)}。核对后重新选择同一 Excel 并勾选确认更新。"
                if is_finance and not is_admin and confirmed:
                    msg = (f"财政安全导入完成：新增 {len([c for c in changes if c[0] == 'insert'])} 笔；"
                           f"原有付款未修改。{summary_text}；异常 {len(errors)}。"
                           "如有异常请复制报告通知管理员。")
                if errors:
                    msg += ' 异常示例：' + '；'.join(errors[:12])
                issues = [d for d in details if d['结果'] in (
                    '格式错误','Excel收据冲突','会员不存在','数据库收据冲突','月份冲突',
                    '数据库重复付款','数据库重复匹配','修改月份冲突')]
                if issues or errors or summary['日期待更新'] or summary['更新']:
                    lines = ['【财政 Excel 对账异常报告】',
                        f"读取 {len(records)} 行；新增 {summary['新增']}；日期待更新 {summary['日期待更新']}；"
                        f"待更新 {summary['更新']}；冲突 {summary['冲突']}；异常 {len(errors)}",'']
                    for d in issues:
                        lines.append(f"{d['结果']}｜{d['Excel 工作表及行号']}｜{d['会员编号']} "
                            f"{d['Excel 姓名']}｜收据 {d['收据号码']}｜{d['说明']}")
                    for d in details:
                        if d['结果'] in ('日期待更新','更新'):
                            lines.append(f"{d['结果']}｜{d['Excel 工作表及行号']}｜{d['会员编号']}｜"
                                f"收据 {d['收据号码']}｜Excel日期 {d['Excel 日期']}｜数据库日期 {d['数据库日期']}｜{d['说明']}")
                    for e in errors:
                        if not any(e in line for line in lines):
                            lines.append('异常：'+e)
                    issue_report = '\n'.join(lines)
            except Exception as exc:
                error = f"财政对账失败（本次事务未提交）：{exc}"

    # 搜索记录
    try:
        if q:
            keyword = q.strip()

            if keyword.isdigit():
                if branch == "STW":
                    member_id = normalize_member_id(f"STW-{keyword}")
                else:
                    member_id = normalize_member_id(keyword)
            else:
                member_id = normalize_member_id(keyword)

            with get_conn() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:
                    if keyword.isdigit() and len(keyword) <= 6:
                        cur.execute("""
                            select *
                            from member_payments
                            where member_id = %s
                            order by payment_date desc, id desc
                            limit 200
                        """, (member_id,))
                    else:
                        cur.execute("""
                            select *
                            from member_payments
                            where member_id ilike %s
                               or name ilike %s
                               or receipt_no ilike %s
                            order by payment_date desc, id desc
                            limit 200
                        """, (
                            f"%{keyword}%",
                            f"%{keyword}%",
                            f"%{keyword}%"
                        ))

                    rows = cur.fetchall()

    except Exception as e:
        error = f"搜索失败：{e}"

    try:
        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    select *
                    from member_payment_history
                    order by changed_at desc
                    limit 30
                """)
                history_rows = cur.fetchall()

                # =========================
                # 财政安全检查
                # =========================

                # 1. 重复收据
                cur.execute("""
                    select
                        receipt_no,
                        count(*) as cnt
                    from member_payments
                    where receipt_no is not null
                    and receipt_no <> ''
                    group by receipt_no
                    having count(*) > 1
                """)

                for r in cur.fetchall():
                    safety_issues.append({
                        "type": "重复收据",
                        "detail": f"收据 {r['receipt_no']} 出现 {r['cnt']} 次"
                    })

                # 2. 金额和月数不符
                cur.execute("""
                    select
                        id,
                        member_id,
                        receipt_no,
                        amount,
                        month_count
                    from member_payments
                """)

                for r in cur.fetchall():

                    amount = float(r["amount"] or 0)
                    months = int(r["month_count"] or 0)

                    if months > 0:

                        expected = months * 50

                        if amount != expected:

                            safety_issues.append({
                                "type": "金额不符",
                                "detail":
                                f"{r['member_id']} / {r['receipt_no']}：RM {amount:.2f}，应为 RM {expected:.2f}",
                                "payment_id": r["id"]
                            })

                # 3. 未来日期
                cur.execute("""
                    select
                        id,
                        member_id,
                        receipt_no,
                        payment_date
                    from member_payments
                    where payment_date > current_date
                """)

                for r in cur.fetchall():

                    safety_issues.append({
                        "type": "未来日期",
                        "detail":
                        f"{r['member_id']} / {r['receipt_no']}：{r['payment_date']}",
                        "payment_id": r["id"]
                    })

                # 4. 会员编号不存在
                cur.execute("""
                    select
                        p.id,
                        p.member_id,
                        p.receipt_no
                    from member_payments p
                    left join members m
                        on p.member_id = m.member_id
                    where m.member_id is null
                """)

                for r in cur.fetchall():

                    safety_issues.append({
                        "type": "会员不存在",
                        "detail":
                        f"{r['member_id']} / {r['receipt_no']}",
                        "payment_id": r["id"]
                    })

    except Exception as e:
        print("读取修改历史失败:", e)

    return render_template_string("""
    <!doctype html>
    <html lang="zh">
    <head>
    <meta charset="utf-8">
    <title>月费管理中心</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <link rel="stylesheet"
        href="{{ url_for('static', filename='css/toolbox.css') }}">

    <style>
    /* 财政工作台：桌面宽屏，避免两列挤在中央 */
    .page.member-center{width:min(1180px,calc(100% - 32px));max-width:1180px;margin:24px auto;box-sizing:border-box;}
    .member-center .member-grid{align-items:start;}
    .member-center .member-grid > .card{min-width:0;}
    .member-center .full-row{grid-column:1 / -1;}
    .member-center .history-toggle{cursor:pointer;background:#f1f5f9;border:1px solid #cbd5e1;border-radius:10px;padding:14px 18px;font-weight:700;font-size:26px;list-style:none;}
    .member-center .history-toggle::-webkit-details-marker{display:none;}
    .member-center .history-toggle::after{content:'展开 ▼';float:right;color:#2563eb;font-size:16px;}
    .member-center details[open] > .history-toggle::after{content:'收起 ▲';}
    .member-center .history-content{padding-top:16px;}
    .member-center .section-title{
        font-size:26px;
        margin-bottom:16px;
    }
    .member-center .card{
        padding:22px;
    }
    .member-center .btn-tool{
        min-height:54px;
        font-size:21px;
    }
    #search-records{scroll-margin-top:18px;}
    /* 工作台采用上下排列：① 上传对账在上，② 搜索记录在下。 */
    .member-grid{
        display:grid;
        grid-template-columns:minmax(0, 1fr);
        gap:16px;
    }
    .month-grid{
        display:grid;
        grid-template-columns:repeat(4,1fr);
        gap:10px;
    }
    .month-option{
        display:flex;
        align-items:center;
        gap:8px;
        font-size:20px;
        font-weight:bold;
        cursor:pointer;
        padding:10px;
        border:1px solid #e5e7eb;
        border-radius:12px;
        background:white;
    }
    .month-option input{
        width:22px;
        height:22px;
    }
    .search-row{
        display:flex;
        gap:10px;
        align-items:center;
    }
    .branch-small-btn{
        width:95px;
        height:54px;
        font-size:20px;
        font-weight:bold;
        background:#16a34a;
        color:white;
        border:none;
        border-radius:12px;
        cursor:pointer;
        flex-shrink:0;
    }
    @media(max-width:700px){
        .member-center .full-row{grid-column:auto;}
        .member-grid{
            grid-template-columns:1fr;
        }
        .month-grid{
            grid-template-columns:repeat(2,1fr);
        }
        .search-row{
            align-items:stretch;
        }
    }
    </style>
    </head>

    <body>

    <div class="page member-center">

        <h1 class="page-title">📂 月费管理中心</h1>
        <p class="page-subtitle">上传财政 Excel、对账、搜索与检查月费资料</p>

        <div class="card">
            {% if is_admin %}<a class="btn-tool btn-secondary" href="/member/admin">← 返回月费管理员</a>
            {% else %}<form method="post" action="{{ url_for('member.finance_upload_logout') }}">
                <button class="btn-tool btn-secondary" type="submit">退出财政上传</button></form>{% endif %}
        </div>

        {% if error %}
        <div class="alert alert-danger">❌ {{ error }}</div>
        {% endif %}

        {% if msg %}
        <div class="alert alert-success">✅ {{ msg }}</div>
        {% endif %}

        {% if issue_report %}
        <div class="card" style="border:2px solid #eab308;margin-bottom:16px">
          <h3>⚠️ 请复制异常报告发给管理员（WhatsApp）</h3>
          <textarea id="financeIssueReport" readonly style="width:100%;min-height:180px;padding:10px">{{issue_report}}</textarea>
          <button class="btn-tool btn-primary" type="button" onclick="navigator.clipboard.writeText(document.getElementById('financeIssueReport').value).then(()=>alert('已复制，可粘贴到 WhatsApp')).catch(()=>{document.getElementById('financeIssueReport').select();document.execCommand('copy');alert('已复制')})">📋 复制异常报告</button>
        </div>
        {% endif %}
        <div class="member-grid">

            <div class="card">
                <div class="section-title">① 财政 Excel 上传与对账</div>

                <form method="post" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>选择 Excel 文件</label>
                        <input class="form-input" name="file" type="file" accept=".xlsx,.xls" required>
                    </div>

                    <label style="display:block;margin:12px 0;">
                        <input type="checkbox" name="confirm_finance" value="yes">
                        {% if is_admin %}我已核对预览结果，确认以财政 Excel 新增／修正同收据同会员的记录
                        {% else %}我已核对预览结果，只导入安全新增付款（冲突、日期及旧记录不修改）{% endif %}
                    </label>
                    <button class="btn-tool btn-primary" type="submit" style="width:100%;">
                        🔍 预览对账／勾选后确认更新
                    </button>
                    {% if is_admin %}<button class="btn-tool btn-secondary" type="submit" name="confirm_dates_only" value="yes"
                            style="width:100%;margin-top:10px;background:#116466;color:white;"
                            onclick="return confirm('仅更新唯一匹配且收据一致的付款日期，其他冲突与金额不修改。确定执行吗？');">
                        📅 只修正安全日期差异（其他冲突保留）
                    </button>{% endif %}
                    <button class="btn-tool btn-secondary" type="submit" name="download_report" value="yes"
                            style="width:100%;margin-top:10px;" onclick="this.form.querySelector('[name=confirm_finance]').checked=false;">
                        📥 下载完整对账明细 Excel（只预览，不更新）
                    </button>
                </form>

                <div class="alert alert-warning">
                    默认只预览，不修改数据库。日期专用按钮仅修正唯一匹配且收据一致的日期，其他异常保留；完整确认按钮仍要求无冲突。确认后批量写入；同收据同会员的差异按财政 Excel 修正并记录审计。不同收据但月份重叠、收据归属冲突及主档缺失，均保留人工核对，不自动覆盖。
                </div>
            </div>

        <div class="card" id="search-records">
            <div class="section-title">② 搜索月费记录</div>

            <form method="get" action="{{ url_for('member.finance_upload') }}#search-records">
                <div class="form-group">
                    <label>寻找编号 / 姓名 / 收据编号</label>

                    <div class="search-row">
                        <button
                            type="button"
                            id="search_branch_btn"
                            onclick="toggleSearchBranch()"
                            class="branch-small-btn">
                            {{ branch }}
                        </button>

                        <input
                            type="hidden"
                            id="search_branch"
                            name="branch"
                            value="{{ branch }}">

                        <input
                            class="form-input"
                            name="q"
                            value="{{ q }}"
                            placeholder="例如：108 / 张三 / CHE0001493"
                            style="flex:1;">
                    </div>
                </div>

                <button class="btn-tool btn-primary" type="submit" style="width:100%;">
                    🔍 搜索记录
                </button>
            </form>

            {% if q %}
            <div class="section-title" style="margin-top:22px;">
                搜索结果：{{ rows|length }} 笔
            </div>

            {% if rows %}
            <div class="table-responsive">
                <table class="record-table">
                    <tr>
                        <th>付款日期</th>
                        <th>月费编号</th>
                        <th>姓名</th>
                        <th>收据编号</th>
                        <th>开始月份</th>
                        <th>结束月份</th>
                        <th>月数</th>
                        <th>金额</th>
                        <th>操作</th>
                    </tr>

                    {% for r in rows %}
                    <tr>
                        <td>{{ r.payment_date.strftime("%Y-%m-%d") if r.payment_date else "-" }}</td>
                        <td>{{ r.member_id }}</td>
                        <td>{{ r.name }}</td>
                        <td>{{ r.receipt_no }}</td>
                        <td>{{ r.start_month.strftime("%Y-%m") if r.start_month else "-" }}</td>
                        <td>{{ r.end_month.strftime("%Y-%m") if r.end_month else "-" }}</td>
                        <td>{{ r.month_count }}</td>
                        <td>RM {{ "%.2f"|format(r.amount or 0) }}</td>
                        <td>
                            <a class="btn-tool btn-primary mini-btn"
                            href="/member/payment/edit/{{ r.id }}">
                                ✏ 编辑
                            </a>
                        </td>
                    </tr>
                    {% endfor %}
                </table>
            </div>
            {% else %}
            <div class="empty-state">
                <div class="empty-title">找不到相关记录</div>
            </div>
            {% endif %}
            {% endif %}
        </div>

        <div class="card full-row">
            <details>
              <summary class="history-toggle">④ 📜 修改历史（{{ history_rows|length }} 笔）</summary>
              <div class="history-content">

            {% if history_rows %}
            <div class="table-responsive">
                <table class="record-table">
                    <tr>
                        <th>时间</th>
                        <th>会员编号</th>
                        <th>收据编号</th>
                        <th>动作</th>
                        <th>修改者</th>
                        <th>详情</th>
                    </tr>

                    {% for h in history_rows %}
                    <tr>
                        <td>{{ (h.changed_at + timedelta(hours=8)).strftime("%Y-%m-%d %H:%M") if h.changed_at else "-" }}</td>
                        <td>{{ h.member_id }}</td>
                        <td>{{ h.receipt_no }}</td>
                        <td>{{ h.action }}</td>
                        <td>{{ h.changed_by }}</td>
                        <td>
                            <a class="btn-tool btn-primary mini-btn"
                            href="/member/payment/history/{{ h.id }}">
                                查看详情
                            </a>
                        </td>
                    </tr>
                    {% endfor %}
                </table>
            </div>
            {% else %}
            <div class="empty-state">
                <div class="empty-title">暂时没有修改记录</div>
            </div>
            {% endif %}
              </div>
            </details>
        </div>

        <div class="card full-row">
            <div class="section-title">⑤ 月费安全检查</div>

            {% if safety_issues %}
            <div class="alert alert-danger">
                ⚠️ 发现 {{ safety_issues|length }} 个问题
            </div>

            <div class="table-responsive">
                <table class="record-table">
                    <tr>
                        <th>类型</th>
                        <th>问题</th>
                        <th>操作</th>
                    </tr>

                    {% for s in safety_issues %}
                    <tr>
                        <td>{{ s.type }}</td>
                        <td>{{ s.detail }}</td>
                        <td>
                            {% if s.payment_id %}
                            <a class="btn-tool btn-primary mini-btn"
                            href="/member/payment/edit/{{ s.payment_id }}">
                                ✏ 编辑
                            </a>
                            {% else %}
                            -
                            {% endif %}
                        </td>
                    </tr>
                    {% endfor %}
                </table>
            </div>
            {% else %}
            <div class="alert alert-success">
                ✅ 没有发现问题
            </div>
            {% endif %}
        </div>

        <div class="card full-row">
            <div class="section-title">⑥ 其他管理</div>

            <div class="member-grid">
                <a class="btn-tool btn-primary" href="/member/member-management">
                    👤 会员资料管理
                </a>

                <a class="btn-tool btn-purple" href="/member/query-logs">
                    🔍 月费查询记录
                </a>
            </div>
        </div>

    </div>

    <script>
    function toggleSearchBranch(){

        const btn = document.getElementById("search_branch_btn");
        const branch = document.getElementById("search_branch");

        if(branch.value==="CHE"){
            branch.value="STW";
            btn.innerText="STW";
            btn.style.background="#dc2626";
        }else{
            branch.value="CHE";
            btn.innerText="CHE";
            btn.style.background="#16a34a";
        }
    }
    </script>

    </body>
    </html>
    """,
    error=error,
    msg=msg,
    rows=rows,
    q=q,
    year=year,
    branch=branch,
    months=months,
    timedelta=timedelta,
    safety_issues=safety_issues,
    history_rows=history_rows,
    is_admin=is_admin,
    issue_report=issue_report
    )

MEMBER_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>月费查询</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="manifest" href="/member-manifest.json?v=3">
<link rel="icon" href="/static/member_icon.png?v=3">
<link rel="apple-touch-icon" href="/static/member_icon.png?v=3">

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
/* =========================
   月费查询页面专用样式
   ========================= */

.member-page{
    max-width:760px;
    margin:0 auto;
}

/* 顶部管理员入口 */
.member-topbar{
    display:flex;
    justify-content:flex-end;
    margin-bottom:8px;
}

.member-topbar .btn-tool{
    width:auto;
    min-height:48px;
    padding:10px 18px;
    font-size:17px;
}

/* 查询输入区域 */
.member-search-row{
    display:grid;
    grid-template-columns:100px minmax(0, 1fr);
    gap:10px;
    align-items:stretch;
}

.branch-btn{
    width:100%;
    min-height:64px;
    padding:0 10px;
    font-size:23px;
    font-weight:800;
    border:0;
    border-radius:16px;
    color:#fff;
    background:#16a34a;
    cursor:pointer;
}

.member-search-row .form-input{
    width:100%;
    min-width:0;
    min-height:64px;
    box-sizing:border-box;
}

/* 会员资料 */
.member-title-line{
    font-size:30px;
    line-height:1.3;
    font-weight:800;
    color:#111827;
}

.member-sub{
    color:#666;
    font-size:17px;
    line-height:1.8;
    margin-top:8px;
    overflow-wrap:anywhere;
}

/* 多位会员选择 */
.candidate-card{
    background:#fff;
    border:1px solid #e5e7eb;
    border-radius:16px;
    padding:16px;
    margin-top:12px;
}

/* 汇总资料 */
.member-summary-grid{
    display:grid;
    grid-template-columns:repeat(3, minmax(0, 1fr));
    gap:12px;
    margin-top:18px;
}

.member-summary-box{
    min-width:0;
    text-align:left;
    box-shadow:none;
    border:1px solid #e5e7eb;
    padding:16px;
}

.member-summary-value{
    margin-top:6px;
    font-size:25px;
    line-height:1.35;
    color:#111827;
    overflow-wrap:anywhere;
}

.member-last-payment{
    margin-top:14px;
    padding:16px;
}

/* 缴费记录表 */
.member-table-wrap{
    width:100%;
    overflow-x:auto;
    -webkit-overflow-scrolling:touch;
    border:1px solid #e5e7eb;
    border-radius:14px;
    margin-top:14px;
}

.record-table{
    width:100%;
    min-width:680px;
    border-collapse:collapse;
    background:#fff;
    margin:0;
}

.record-table th,
.record-table td{
    border-right:1px solid #e5e7eb;
    border-bottom:1px solid #e5e7eb;
    padding:11px 9px;
    text-align:center;
    font-size:15px;
    line-height:1.4;
    white-space:nowrap;
}

.record-table th:last-child,
.record-table td:last-child{
    border-right:0;
}

.record-table tr:last-child td{
    border-bottom:0;
}

.record-table th{
    background:#f3f4f6;
    font-weight:700;
    color:#374151;
}

/* 手机版 */
@media(max-width:700px){

    body{
        overflow-x:hidden;
    }

    .member-page{
        width:100%;
        max-width:none;
    }

    .member-topbar{
        justify-content:flex-end;
        margin-bottom:4px;
    }

    .member-topbar .btn-tool{
        min-height:42px;
        padding:8px 13px;
        font-size:15px;
        border-radius:12px;
    }

    .member-search-row{
        grid-template-columns:78px minmax(0, 1fr);
        gap:8px;
    }

    .branch-btn{
        min-height:56px;
        font-size:19px;
        border-radius:13px;
    }

    .member-search-row .form-input{
        min-height:56px;
        height:56px;
        padding:10px 12px;
        font-size:16px;
        border-radius:13px;
    }

    .member-title-line{
        font-size:24px;
    }

    .member-sub{
        font-size:16px;
        line-height:1.7;
    }

    /*
      手机汇总排成：
      第一行两格
      第二行“月费已缴至”占满
    */
    .member-summary-grid{
        grid-template-columns:repeat(2, minmax(0, 1fr));
        gap:10px;
        margin-top:14px;
    }

    .member-summary-box{
        padding:13px;
        border-radius:13px;
    }

    .member-summary-box:nth-child(3){
        grid-column:1 / -1;
    }

    .member-summary-box .summary-title{
        font-size:14px;
        line-height:1.4;
    }

    .member-summary-value{
        font-size:20px;
        line-height:1.4;
    }

    .member-last-payment{
        padding:13px;
    }

    .candidate-card{
        padding:14px;
        border-radius:14px;
    }

    .record-table{
        min-width:650px;
    }

    .record-table th,
    .record-table td{
        font-size:13px;
        padding:9px 8px;
    }
}

/* 很小的手机 */
@media(max-width:390px){

    .member-search-row{
        grid-template-columns:70px minmax(0, 1fr);
        gap:7px;
    }

    .branch-btn{
        font-size:17px;
    }

    .member-search-row .form-input{
        font-size:15px;
        padding-left:10px;
        padding-right:10px;
    }

    .member-summary-value{
        font-size:18px;
    }
}
</style>
</head>

<body>

<div class="page member-page">

    <div class="member-topbar">
        <a class="btn-tool btn-light" href="/member/admin">
            ⚙ 管理员入口
        </a>
    </div>

    <h1 class="page-title">🙏 月费查询</h1>
    <p class="page-subtitle">请输入月费编号、姓名、英文名或电话查询供养记录</p>

    {% if query_volunteer %}
    <div class="alert alert-warning">
        当前查询义工：
        <b>{{ query_volunteer.id }} {{ query_volunteer.name }}</b>

        <div style="margin-top:12px;">
            <a class="btn-tool btn-danger" href="/member/query-logout">
                🔒 退出查询
            </a>
        </div>
    </div>
    {% endif %}

    {% if error %}
    <div class="alert alert-danger">
        ❌ {{ error }}
    </div>
    {% endif %}

    <div class="card">
        <div class="section-title">🔍 查询资料</div>

        <form method="post">
            <div class="form-group">
                <label>月费编号 / 姓名 / 英文名 / 电话</label>

                <div class="member-search-row">
                    <button
                        type="button"
                        id="branch_btn"
                        class="branch-btn"
                        onclick="toggleBranch()"
                    >
                        CHE
                    </button>

                    <input type="hidden" id="branch" name="branch" value="CHE">

                    <input
                        class="form-input"
                        id="member_id"
                        name="member_id"
                        placeholder="例如：108 / 张三 / 0123456789"
                        autocomplete="off"
                        required
                    >
                </div>
            </div>

            <button class="btn-tool btn-primary" type="submit" style="width:100%;">
                🔍 查询
            </button>
        </form>
    </div>

    {% if candidates %}
    <div id="member-result" class="result-card">
        <div class="section-title">👥 找到多位会员，请选择</div>

        <form method="post">
            {% for m in candidates %}
            <div class="candidate-card">
                <div class="member-title-line">{{ m.name }}</div>
                <div class="member-sub">
                    编号：<b>{{ m.member_id }}</b><br>
                    {% if m.english_name %}
                    英文名：{{ m.english_name }}<br>
                    {% endif %}
                    电话：{{ m.phone | member_phone_display }}
                </div>

                <button
                    class="btn-tool btn-primary"
                    type="submit"
                    name="selected_member_id"
                    value="{{ m.member_id }}"
                    style="width:100%;margin-top:12px;"
                >
                    查看这个会员
                </button>
            </div>
            {% endfor %}
        </form>
    </div>
    {% endif %}

    {% if member %}
    <div id="member-result" class="member-card">
        <div class="section-title">👤 会员资料</div>

        <div class="member-title-line">
            {{ member.name }}
        </div>

        <div class="member-sub">
            月费编号：<b>{{ member.member_id }}</b><br>
            {% if member.english_name %}
            英文名：{{ member.english_name }}<br>
            {% endif %}
            电话：{{ member.phone | member_phone_display }}
        </div>

        {% if summary %}
        {% if member and (member.member_status == '停供' or (not member.member_status and member.status == '停供')) %}
        <div style="padding:14px 18px;margin:12px 0;border-radius:12px;background:#fff1d6;color:#825000;font-weight:700;">
            ⏸️ 已停供｜最后供养月份：{{ summary.paid_until.strftime('%Y-%m') if summary and summary.paid_until else '暂无记录' }}<br>
            <small>停供期间不显示欠缴月份，历史付款记录保留。</small>
        </div>
        {% endif %}

        <div class="summary-grid member-summary-grid">

            <div class="summary-box member-summary-box">
                <div class="summary-title">已付月费总额</div>
                <div class="summary-value member-summary-value">
                    RM {{ "%.2f"|format(summary.total_payment or 0) }}
                </div>
            </div>

            <div class="summary-box member-summary-box">
                <div class="summary-title">累计已付月数</div>
                <div class="summary-value member-summary-value">
                    {{ summary.total_months or 0 }} 个月
                </div>
            </div>

            <div class="summary-box member-summary-box">
                <div class="summary-title">月费已缴至</div>
                <div class="summary-value member-summary-value">
                    {% if summary.paid_until %}
                        {{ summary.paid_until.strftime("%Y-%m") }}
                    {% else %}
                        暂无记录
                    {% endif %}
                </div>
            </div>

        </div>

        <div class="card-soft member-last-payment">
            <div class="summary-title">最后付款记录</div>
            <div class="summary-value member-summary-value">
                {% if summary.last_payment_date %}
                    {{ summary.last_payment_date.strftime("%Y-%m-%d") }}
                    ｜RM {{ "%.2f"|format(summary.last_payment_amount or 0) }}
                {% else %}
                    暂无记录
                {% endif %}
            </div>
        </div>
        {% endif %}
    </div>

    <div class="card">
        <div class="section-title">📋 缴费记录</div>

        {% if payments %}
        <div class="member-table-wrap">
            <table class="record-table">
                <tr>
                    <th>付款日期</th>
                    <th>收据编号</th>
                    <th>开始月份</th>
                    <th>结束月份</th>
                    <th>月数</th>
                    <th>金额</th>
                </tr>

                {% for p in payments %}
                <tr>
                    <td>{{ p.payment_date.strftime("%Y-%m-%d") if p.payment_date else "-" }}</td>
                    <td>{{ p.receipt_no }}</td>
                    <td>{{ p.start_month.strftime("%Y-%m") if p.start_month else "-" }}</td>
                    <td>{{ p.end_month.strftime("%Y-%m") if p.end_month else "-" }}</td>
                    <td>{{ p.month_count }}</td>
                    <td>RM {{ "%.2f"|format(p.amount or 0) }}</td>
                </tr>
                {% endfor %}
            </table>
        </div>
        {% else %}
        <div class="empty-state">
            这个会员目前没有缴费记录。
        </div>
        {% endif %}
    </div>
    {% endif %}

</div>

<script>
function toggleBranch() {
    const btn = document.getElementById("branch_btn");
    const branch = document.getElementById("branch");

    if (branch.value === "CHE") {
        branch.value = "STW";
        btn.innerText = "STW";
        btn.style.background = "#dc2626";
    } else {
        branch.value = "CHE";
        btn.innerText = "CHE";
        btn.style.background = "#16a34a";
    }
}

window.addEventListener("load", function () {

    const result = document.getElementById("member-result");

    if (result) {
        setTimeout(function () {
            result.scrollIntoView({
                behavior: "smooth",
                block: "start"
            });
        }, 150);
    }

});

</script>

</body>
</html>
"""

MEMBER_ADMIN_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>月费管理员</title>
<meta name="viewport" content="width=device-width, initial-scale=1">

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.admin-link-grid{
    display:grid;
    grid-template-columns:repeat(3, 1fr);
    gap:12px;
}

.member-search-row{
    display:flex;
    gap:10px;
    align-items:center;
}

.branch-btn{
    width:100px;
    min-height:64px;
    font-size:24px;
    font-weight:800;
    border:0;
    border-radius:16px;
    color:white;
    background:#16a34a;
    flex-shrink:0;
    cursor:pointer;
}

.pagination{
    margin-top:20px;
    text-align:center;
}

.pagination a,
.pagination span{
    display:inline-block;
    padding:8px 13px;
    margin:3px;
    border-radius:8px;
    text-decoration:none;
    font-size:16px;
}

.pagination a{
    background:#e5e7eb;
    color:#333;
}

.pagination span{
    background:#2563eb;
    color:white;
    font-weight:bold;
}

@media(max-width:700px){
    .admin-link-grid{
        grid-template-columns:1fr;
    }

    .member-search-row{
        align-items:stretch;
    }

    .branch-btn{
        width:82px;
        font-size:21px;
    }
}
</style>
</head>

<body>

<div class="page">

    <h1 class="page-title">🙏 月费管理员</h1>
    <p class="page-subtitle">管理员查询佛友月费记录、检查缴费资料与编辑记录</p>

    <div class="card">
        <div class="section-title">📂 功能入口</div>

        <div class="admin-link-grid">
            <a class="btn-tool btn-secondary" href="/member">
                ← 月费查询
            </a>

            <a class="btn-tool btn-primary" href="/member/finance-upload">
                📂 资料管理中心
            </a>

            <a class="btn-tool btn-warning" href="/member/admin/late_members">
                ⚠️ 迟付名单
            </a>
            {% if session.get("member_admin") %}
            <a class="btn-tool btn-primary" href="#admin-export">
                ⬇️ 下载月费 Excel
            </a>
            {% endif %}
        </div>
    </div>

    {% if session.get("member_admin") %}
    <div class="card" id="admin-export">
        <div class="section-title">⬇️ 下载月费 Excel（管理员专用）</div>
        <p>从 Supabase 导出月费付款记录；不选择月份即下载全年。</p>
                <form method="get" action="/member/finance-upload">
                    <input type="hidden" name="download" value="1">

                    <div class="form-group">
                        <label>年份</label>
                        <input class="form-input" name="year" value="{{ year }}">
                    </div>

                    <div class="form-group">
                        <label>月份</label>
                        <div class="small-text">不打勾 = 下载全年，可同时选择多个月</div>
                    </div>

                    <div class="month-grid">
                        {% for m in range(1,13) %}
                        <label class="month-option">
                            <input
                                type="checkbox"
                                name="months"
                                value="{{ '%02d'|format(m) }}"
                            >
                            {{ m }}月
                        </label>
                        {% endfor %}
                    </div>

                    <button class="btn-tool btn-primary" type="submit" style="width:100%;">
                        ⬇️ 下载 Excel
                    </button>
                </form>
    </div>
    {% endif %}

    {% if session.get("member_admin") %}
    <div class="alert alert-success">
        ✅ 管理员已登录

        <div style="margin-top:12px;">
            <a class="btn-tool btn-danger" href="/member/admin/logout">
                🚪 退出管理员
            </a>
        </div>
    </div>
    {% endif %}

    {% if error %}
    <div class="alert alert-danger">
        ❌ {{ error }}
    </div>
    {% endif %}

    <div class="card">
        {% if not session.get("member_admin") %}

        <div class="section-title">🔐 管理员登录</div>

        <form method="post" action="/member/admin">
            <div class="form-group">
                <label>管理员 PIN</label>

                <input
                    class="form-input"
                    id="member_admin_pin"
                    name="admin_pin"
                    type="password"
                    inputmode="numeric"
                    autocomplete="new-password"
                    readonly
                    onfocus="this.removeAttribute('readonly');"
                    required
                >
            </div>

            <button class="btn-tool btn-primary" type="submit" style="width:100%;">
                登录管理员
            </button>
        </form>

        {% else %}

        <div class="section-title">🔍 查询会员</div>

        <form method="post" action="/member/admin">
            <div class="form-group">
                <label>月费编号 / 姓名 / 英文名</label>

                <div class="member-search-row">
                    <button
                        type="button"
                        id="branch_btn"
                        class="branch-btn"
                        onclick="toggleBranch()"
                    >
                        CHE
                    </button>

                    <input type="hidden" id="branch" name="branch" value="CHE">

                    <input
                        class="form-input"
                        name="member_id"
                        placeholder="例如：108 / 张三 / zhangsan"
                        autocomplete="off"
                    >
                </div>
            </div>

            <button class="btn-tool btn-primary" type="submit" style="width:100%;">
                查找会员
            </button>
        </form>

        {% endif %}
    </div>

    {% if member %}
    <div id="member-result" class="member-card">
        <div class="section-title">👤 会员资料</div>

        <div class="person-name">
            {{ member.name }}
        </div>

        <div class="person-meta">
            月费编号：<b>{{ member.member_id }}</b><br>
            {% if member.english_name %}
            英文名：{{ member.english_name }}<br>
            {% endif %}
            电话：{{ member.phone | member_phone_display }}
        </div>
    </div>

    {% if summary %}
    <div class="card">
        <div class="section-title">📊 月费统计</div>

        <div class="info-summary-grid">

            <div class="info-summary-box">
                <div class="info-summary-title">已付月费总额</div>
                <div class="info-summary-value">
                    RM {{ "%.2f"|format(summary.total_payment or 0) }}
                </div>
            </div>

            <div class="info-summary-box">
                <div class="info-summary-title">累计已付月数</div>
                <div class="info-summary-value">
                    {{ summary.total_months or 0 }} 个月
                </div>
            </div>

            <div class="info-summary-box">
                <div class="info-summary-title">月费已缴至</div>
                <div class="info-summary-value">
                    {% if summary.paid_until %}
                        {{ summary.paid_until.strftime("%Y-%m") }}
                    {% else %}
                        暂无记录
                    {% endif %}
                </div>
            </div>

        </div>

        <div class="info-summary-box" style="margin-top:14px;">
            <div class="info-summary-title">最后付款记录</div>
            <div class="info-summary-value">
                {% if summary.last_payment_date %}
                    {{ summary.last_payment_date.strftime("%Y-%m-%d") }}
                    ｜RM {{ "%.2f"|format(summary.last_payment_amount or 0) }}
                {% else %}
                    暂无记录
                {% endif %}
            </div>
        </div>
    </div>
    {% endif %}

    {% if warnings %}
    <div class="alert alert-warning">
        <b>⚠️ 系统发现可能有错误：</b>
        <ul>
            {% for w in warnings %}
                <li>{{ w }}</li>
            {% endfor %}
        </ul>
    </div>
    {% endif %}

    <div class="card">
        <div class="section-title">📋 缴费记录</div>

        {% if payments %}
        <div class="table-responsive">
            <table class="record-table">
                <tr>
                    <th>日期<br>Date</th>
                    <th>收据编号<br>Official Receipt No</th>
                    <th>编号 No.</th>
                    <th>捐款人<br>姓名</th>
                    <th>START MONTH</th>
                    <th>END MONTH</th>
                    <th>No. of Mth</th>
                    <th>Total Amt</th>
                    <th>操作</th>
                </tr>

                {% for p in payments %}
                <tr>
                    <td>{{ p.payment_date.strftime("%d/%m/%Y") if p.payment_date else "-" }}</td>
                    <td>{{ p.receipt_no }}</td>
                    <td>{{ p.member_id.replace("CHE-", "").replace("STW-", "") if p.member_id else "-" }}</td>
                    <td>{{ p.name or member.name }}</td>
                    <td>{{ p.start_month.strftime("%b-%y") if p.start_month else "-" }}</td>
                    <td>{{ p.end_month.strftime("%b-%y") if p.end_month else "-" }}</td>
                    <td>{{ p.month_count }}</td>
                    <td>{{ "%.2f"|format(p.amount or 0) }}</td>
                    <td>
                        <a class="btn-tool btn-primary mini-btn"
                           href="/member/payment/edit/{{ p.id }}">
                            ✏ 编辑
                        </a>
                    </td>
                </tr>
                {% endfor %}
            </table>
        </div>

        {% if total_pages and total_pages > 1 %}
        <div class="pagination">
            {% for p in range(1, total_pages + 1) %}
                {% if p == page %}
                    <span>{{ p }}</span>
                {% else %}
                    <a href="/member/admin?member_id={{ raw_member_id }}&page={{ p }}">{{ p }}</a>
                {% endif %}
            {% endfor %}
        </div>
        {% endif %}

        {% else %}
        <div class="empty-state">
            <div class="empty-icon">📭</div>
            <div class="empty-title">暂无缴费记录</div>
            <div class="empty-text">这个会员目前没有缴费记录。</div>
        </div>
        {% endif %}
    </div>
    {% endif %}

</div>

<script>
function toggleBranch() {
    const btn = document.getElementById("branch_btn");
    const branch = document.getElementById("branch");

    if (branch.value === "CHE") {
        branch.value = "STW";
        btn.innerText = "STW";
        btn.style.background = "#dc2626";
    } else {
        branch.value = "CHE";
        btn.innerText = "CHE";
        btn.style.background = "#16a34a";
    }
}

document.addEventListener("DOMContentLoaded", function () {

    const result = document.getElementById("member-result");

    if (result) {
        setTimeout(function () {

            const topPosition =
                result.getBoundingClientRect().top
                + window.pageYOffset
                - 20;

            window.scrollTo({
                top: topPosition,
                behavior: "smooth"
            });

        }, 300);
    }

});

</script>

</body>
</html>
"""

PAYMENT_EDIT_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>修改缴费记录</title>
<meta name="viewport" content="width=device-width, initial-scale=1">

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">
</head>

<body>

<div class="page">

    <h1 class="page-title">✏ 修改缴费记录</h1>
    <p class="page-subtitle">请确认日期、收据编号、月份与金额是否正确</p>

    <div class="card">
        <a class="btn-tool btn-secondary"
           href="/member/admin?member_id={{ payment.member_id }}">
            ← 返回会员记录
        </a>
    </div>

    {% if error %}
    <div class="alert alert-danger">
        ❌ {{ error }}
    </div>
    {% endif %}

    <div class="card">
        <div class="section-title">📋 缴费资料</div>

        <form method="post">

            <div class="form-group">
                <label>日期 Date</label>
                <input
                    class="form-input"
                    type="date"
                    name="payment_date"
                    value="{{ payment.payment_date.strftime('%Y-%m-%d') if payment.payment_date else '' }}"
                    required
                >
            </div>

            <div class="form-group">
                <label>收据编号 Official Receipt No</label>
                <input
                    class="form-input"
                    name="receipt_no"
                    value="{{ payment.receipt_no or '' }}"
                    required
                >
            </div>

            <div class="form-group">
                <label>编号 No.</label>
                <input
                    class="form-input"
                    value="{{ payment.member_id.replace('CHE-', '').replace('STW-', '') }}"
                    disabled
                >
            </div>

            <div class="form-group">
                <label>捐款人 姓名 Name</label>
                <input
                    class="form-input"
                    name="name"
                    value="{{ payment.name or '' }}"
                    required
                >
            </div>

            <div class="form-grid">
                <div class="form-group">
                    <label>START MONTH</label>
                    <input
                        class="form-input"
                        type="month"
                        name="start_month"
                        value="{{ payment.start_month.strftime('%Y-%m') if payment.start_month else '' }}"
                        required
                    >
                </div>

                <div class="form-group">
                    <label>END MONTH</label>
                    <input
                        class="form-input"
                        type="month"
                        name="end_month"
                        value="{{ payment.end_month.strftime('%Y-%m') if payment.end_month else '' }}"
                        required
                    >
                </div>
            </div>

            <div class="form-grid">
                <div class="form-group">
                    <label>No. of Mth</label>
                    <input
                        class="form-input"
                        type="number"
                        name="month_count"
                        value="{{ payment.month_count or 1 }}"
                        required
                    >
                </div>

                <div class="form-group">
                    <label>Total Amt</label>
                    <input
                        class="form-input"
                        type="number"
                        step="0.01"
                        name="amount"
                        value="{{ payment.amount or 0 }}"
                        required
                    >
                </div>
            </div>

            <button class="btn-tool btn-primary" type="submit" style="width:100%;">
                💾 保存修改
            </button>

        </form>
    </div>

</div>

</body>
</html>
"""

CHANGE_PIN_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>更改月费密码</title>
<meta name="viewport" content="width=device-width, initial-scale=1">

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">
</head>

<body>

<div class="page">

    <h1 class="page-title">🔐 更改月费密码</h1>
    <p class="page-subtitle">请输入会员编号、旧密码及新密码</p>

    <div class="card">
        <a class="btn-tool btn-secondary" href="/member">
            ← 返回月费查询
        </a>
    </div>

    {% if error %}
    <div class="alert alert-danger">
        ❌ {{ error }}
    </div>
    {% endif %}

    {% if ok %}
    <div class="alert alert-success">
        ✅ {{ ok }}
    </div>
    {% endif %}

    <div class="card">

        <div class="section-title">
            🔑 修改密码
        </div>

        <form method="post">

            <div class="form-group">
                <label>月费编号</label>

                <input
                    class="form-input"
                    name="member_id"
                    placeholder="例如：108 / CHE-108 / STW-108"
                    autocomplete="off"
                    required
                >
            </div>

            <div class="form-group">
                <label>旧密码</label>

                <input
                    class="form-input"
                    name="old_pin"
                    type="password"
                    inputmode="numeric"
                    autocomplete="new-password"
                    autocorrect="off"
                    autocapitalize="off"
                    spellcheck="false"
                    readonly
                    onfocus="this.removeAttribute('readonly');"
                    required
                >
            </div>

            <div class="form-group">
                <label>新密码</label>

                <input
                    class="form-input"
                    name="new_pin"
                    type="password"
                    inputmode="numeric"
                    autocomplete="new-password"
                    autocorrect="off"
                    autocapitalize="off"
                    spellcheck="false"
                    readonly
                    onfocus="this.removeAttribute('readonly');"
                    required
                >
            </div>

            <div class="form-group">
                <label>确认新密码</label>

                <input
                    class="form-input"
                    name="confirm_pin"
                    type="password"
                    inputmode="numeric"
                    autocomplete="new-password"
                    autocorrect="off"
                    autocapitalize="off"
                    spellcheck="false"
                    readonly
                    onfocus="this.removeAttribute('readonly');"
                    required
                >
            </div>

            <button
                class="btn-tool btn-primary"
                style="width:100%;"
                type="submit">
                💾 确认更改密码
            </button>

        </form>

    </div>

</div>

</body>
</html>
"""

def format_history_changes(old_data, new_data):
    old = old_data or {}
    new = new_data or {}

    labels = {
        "payment_date": "付款日期",
        "receipt_no": "收据编号",
        "name": "姓名",
        "start_month": "开始月份",
        "end_month": "结束月份",
        "month_count": "月数",
        "amount": "金额",
    }

    changes = []

    for key, label in labels.items():
        old_value = old.get(key, "")
        new_value = new.get(key, "")

        # 新资料没有这个字段
        if key not in new:
            continue

        # 金额格式统一
        if key == "amount":
            try:
                old_value = float(old_value)
                new_value = float(new_value)
            except:
                pass

        if str(old_value) != str(new_value):

            changes.append({
                "label": label,
                "old": old_value or "-",
                "new": new_value or "-"
            })

    return changes


PAYMENT_HISTORY_DETAIL_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>修改历史详情</title>
<meta name="viewport" content="width=device-width, initial-scale=1">

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">
</head>

<body>

<div class="page">

    <h1 class="page-title">📝 修改历史详情</h1>
    <p class="page-subtitle">查看这笔缴费记录曾经修改过的内容</p>

    <div class="card">
        <a class="btn-tool btn-secondary"
           href="/member/finance-upload">
            ← 返回月费资料管理中心
        </a>
    </div>

    <div class="info-card">

        <div class="section-title">
            📄 基本资料
        </div>

        <div class="person-meta">
            <b>修改时间：</b>
            {{ (h.changed_at + timedelta(hours=8)).strftime("%Y-%m-%d %H:%M") if h.changed_at else "-" }}
            <br>

            <b>会员编号：</b>{{ h.member_id }}
            <br>

            <b>收据编号：</b>{{ h.receipt_no }}
            <br>

            <b>修改者：</b>{{ h.changed_by }}
        </div>

    </div>

    <div class="card">

        <div class="section-title">
            🔍 修改内容
        </div>

        {% if changes %}

        <div class="table-responsive">

            <table class="record-table">

                <tr>
                    <th>项目</th>
                    <th>修改前</th>
                    <th>修改后</th>
                </tr>

                {% for c in changes %}
                <tr>

                    <td>
                        {{ c.label }}
                    </td>

                    <td style="color:#dc2626;font-weight:bold;">
                        {{ c.old }}
                    </td>

                    <td style="color:#16a34a;font-weight:bold;">
                        {{ c.new }}
                    </td>

                </tr>
                {% endfor %}

            </table>

        </div>

        {% else %}

        <div class="empty-state">

            <div class="empty-icon">
                📭
            </div>

            <div class="empty-title">
                没有发现明显变化
            </div>

            <div class="empty-text">
                这次修改没有记录到任何字段差异。
            </div>

        </div>

        {% endif %}

    </div>

</div>

</body>
</html>
"""