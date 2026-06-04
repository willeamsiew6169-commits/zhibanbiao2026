# finance_web.py

import os
import re

from flask import Blueprint, request, redirect, url_for, render_template_string, send_file
from psycopg2.extras import RealDictCursor
from datetime import date
from openpyxl import Workbook
from db import db_query
from utils import normalize_member_id
from flask import send_file, request
from openpyxl import Workbook
from flask import session
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

finance_bp = Blueprint("finance", __name__, url_prefix="/finance")

FINANCE_PIN = "1234"

INCOME_CATEGORIES = [
    "月费",
    "财布施",
    "观音村",
    "膳食结缘"
]

FINANCE_STYLE = """
<style>
body{
    font-family: Microsoft JhengHei, Arial;
    font-size:22px;
    padding:25px;
}

h1{
    font-size:42px;
    margin-bottom:25px;
}

label{
    font-size:24px;
    font-weight:bold;
}

input,
select,
textarea{
    font-size:22px;
    padding:10px;
    min-height:50px;
    min-width:260px;
}

button,
input[type="submit"]{
    font-size:22px;
    padding:14px 28px;
    border-radius:10px;
    cursor:pointer;
}

a{
    font-size:22px;
}

table{
    font-size:20px;
    border-collapse:collapse;
    margin-top:20px;
}

th,td{
    padding:12px;
}

.card{
    padding:25px;
    border-radius:12px;
    box-shadow:0 2px 8px rgba(0,0,0,.15);
    margin-bottom:25px;
}

form{
    line-height:2.2;
}
</style>
"""


def money(v):
    try:
        return float(v or 0)
    except:
        return 0
    
def get_fund_account(category, record_type="income"):
    if record_type == "expense":
        return "观音堂日常户口"

    if category == "月费":
        return "观音堂日常户口"

    return "总会户口"
    
def add_months_ym(ym, months):
    y, m = map(int, ym.split("-"))
    m += months
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return f"{y:04d}-{m:02d}"

def date_to_ym(d):
    if not d:
        return ""
    return d.strftime("%Y-%m")

def next_month_ym(d):
    if not d:
        return date.today().strftime("%Y-%m")
    return add_months_ym(d.strftime("%Y-%m"), 1)

def calc_month_count(amount):
    if amount <= 0:
        return 1
    return max(1, round(amount / 50))

def get_next_receipt_no_by_category(category):
    row = db_query("""
        select receipt_no
        from finance_records
        where category = %s
          and receipt_no is not null
          and receipt_no <> ''
        order by id desc
        limit 1
    """, (category,), fetchone=True)

    if not row or not row["receipt_no"]:
        return ""

    last_no = row["receipt_no"].strip()

    match = re.match(r"^([A-Za-z]+)(\d+)$", last_no)

    if not match:
        return ""

    prefix = match.group(1)
    number = match.group(2)

    next_number = str(int(number) + 1).zfill(len(number))

    return prefix + next_number

@finance_bp.route("/login", methods=["GET", "POST"])
def finance_login():

    if request.method == "POST":

        pin = request.form.get("pin", "").strip()

        if pin == FINANCE_PIN:
            session["finance_login"] = True
            return redirect(url_for("finance.finance_home"))

    return render_template_string(FINANCE_STYLE + """
    <h1>🏦 财政系统登入</h1>

    <div class="card">
        <form method="post">

            <p>
                <label>财政 PIN：</label>
                <input
                    name="pin"
                    type="password"
                    inputmode="numeric"
                    autocomplete="new-password"
                    required
                >
            </p>

            <button type="submit">
                进入财政系统
            </button>

        </form>
    </div>

    <p>
        <a href="/admin-home">返回管理员首页</a>
    </p>
    """)

@finance_bp.route("/logout")
def finance_logout():

    session.pop("finance_login", None)

    return redirect(url_for("finance.finance_login"))

@finance_bp.route("/late_members")
def late_members():

    rows = db_query("""
        select
            m.member_id,
            m.name,
            m.phone,
            max(p.end_month) as paid_until
        from members m
        left join member_payments p
            on p.member_id = m.member_id
        group by
            m.member_id,
            m.name,
            m.phone
        having
            max(p.end_month) is not null
            and max(p.end_month) < date_trunc('month', current_date)
        order by
            paid_until asc,
            m.member_id
    """, fetchall=True)

    return render_template_string(FINANCE_STYLE + """
<h1>⚠️ 月费迟付名单</h1>

<p>
    共 {{ rows|length }} 人
</p>

<table border="1" cellpadding="8">
    <tr>
        <th>会员编号</th>
        <th>姓名</th>
        <th>电话</th>
        <th>已缴费至</th>
    </tr>

    {% for r in rows %}
    <tr>
        <td>{{ r.member_id }}</td>
        <td>{{ r.name }}</td>
        <td>{{ r.phone or "-" }}</td>
        <td>{{ r.paid_until.strftime("%Y-%m") }}</td>
    </tr>
    {% endfor %}
</table>

<p>
    <a href="{{ url_for('finance.finance_home') }}">
        返回财政首页
    </a>
</p>
""", rows=rows)

    
@finance_bp.route("/records/<int:record_id>/delete", methods=["POST"])
def delete_record(record_id):

    db_query("""
        delete from finance_records
        where id = %s
    """, (record_id,))

    return redirect(url_for("finance.records"))


@finance_bp.route("/")
def finance_home():

    if not session.get("finance_login"):
        return redirect(url_for("finance.finance_login"))

    today_ym = date.today().strftime("%Y-%m")

    return render_template_string("""
    <style>
        body {
            font-family: Arial, "Microsoft YaHei", sans-serif;
            background: #f5f6fa;
            padding: 30px;
        }

        h1 {
            margin-bottom: 25px;
        }

        .topbar {
            text-align: right;
            margin-bottom: 15px;
        }

        .logout-btn {
            background: #dc3545;
            color: white;
            padding: 10px 18px;
            border-radius: 10px;
            text-decoration: none;
            font-weight: bold;
        }

        .section {
            margin-bottom: 30px;
        }

        .title {
            font-size: 18px;
            margin-bottom: 12px;
            color: #444;
            font-weight: bold;
        }

        .grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 18px;
            max-width: 1100px;
        }

        .card {
            min-height: 150px;
            display: flex;
            flex-direction: column;
            padding: 22px;
            border-radius: 14px;
            text-decoration: none;
            color: #222;
            box-shadow: 0 3px 10px rgba(0,0,0,0.08);
            transition: 0.2s;
            box-sizing: border-box;
        }

        .card:hover {
            transform: translateY(-3px);
        }

        .card h2 {
            margin: 0 0 12px 0;
            font-size: 22px;
            line-height: 1.3;
        }

        .card p {
            margin: 4px 0;
            font-size: 14px;
            font-weight: normal;
            color: #555;
        }

        .card small {
            display: block;
            margin-top: 8px;
            color: #666;
            font-size: 14px;
            font-weight: normal;
        }

        .card-member {
            background: #eef5ff;
            border-left: 6px solid #0d6efd;
        }

        .card-income {
            background: #e8f7ec;
            border-left: 6px solid #28a745;
        }

        .card-expense {
            background: #fff1f1;
            border-left: 6px solid #dc3545;
        }

        .card-report {
            background: #fff8e6;
            border-left: 6px solid #f0ad4e;
        }

        .card-warning {
            background: #fff4e5;
            border-left: 6px solid #ff9800;
        }
    </style>

    <h1>💰 财政系统</h1>

    <div class="topbar">
        <a class="logout-btn" href="{{ url_for('finance.finance_logout') }}">
            🚪 退出财政系统
        </a>
    </div>

    <div class="section">
        <div class="title">提醒事项</div>

        <div class="grid">
            <a class="card card-warning" href="{{ url_for('finance.late_members') }}">
                <h2>⚠️ 月费迟付名单</h2>
                <p>查看已过期缴费会员</p>
            </a>
        </div>
    </div>

    <div class="section">
        <div class="title">日常录入</div>

        <div class="grid">
            <a class="card card-member" href="{{ url_for('finance.monthly_fee', branch='CHE') }}">
                <h2>CHE 月费收款</h2>
                <p>会员月费登记</p>
                <p>收条：CHE0000001 起</p>
            </a>

            <a class="card card-member" href="{{ url_for('finance.monthly_fee', branch='STW') }}">
                <h2>STW 月费收款</h2>
                <p>会员月费登记</p>
                <p>收条：STW0000001 起</p>
            </a>

            <a class="card card-member" href="{{ url_for('finance.bank_pending') }}">
                <h2>🏦 银行过账中心</h2>
                <p>导入 Receipt</p>
                <p>查看待确认记录</p>
            </a>
                                  
            <a class="card card-income" href="{{ url_for('finance.normal_income', category='财布施') }}">
                <h2>财布施</h2>
                <p>普通布施收入记录</p>
            </a>

            <a class="card card-income" href="{{ url_for('finance.normal_income', category='观音村') }}">
                <h2>观音村</h2>
                <p>观音村相关收入记录</p>
            </a>

            <a class="card card-income" href="{{ url_for('finance.normal_income', category='膳食结缘') }}">
                <h2>膳食结缘</h2>
                <p>膳食结缘收入记录</p>
            </a>
        </div>
    </div>

    <div class="section">
        <div class="title">报表与查询</div>

        <div class="grid">
            <a class="card card-report" href="{{ url_for('finance.dashboard') }}">
                <h2>财政 Dashboard</h2>
                <p>查看本月收入与支出统计</p>
            </a>

            <a class="card card-report" href="{{ url_for('finance.records') }}">
                <h2>财政记录搜索</h2>
                <p>搜索收条、会员编号、姓名</p>
            </a>

            <a class="card card-report" href="{{ url_for('finance.export_monthly_report', ym=today_ym) }}">
                <h2>📊 下载专业版月报</h2>
                <p>{{ today_ym }} 财政月报</p>
            </a>
        </div>
    </div>

    <div class="section">
        <div class="title">支出记录</div>

        <div class="grid">
            <a class="card card-expense" href="{{ url_for('finance.expense', category='供花') }}">
                <h2>供花</h2>
                <p>鲜花、供花相关支出</p>
            </a>

            <a class="card card-expense" href="{{ url_for('finance.expense', category='供果') }}">
                <h2>供果</h2>
                <p>水果、供品相关支出</p>
            </a>

            <a class="card card-expense" href="{{ url_for('finance.expense', category='供油') }}">
                <h2>供油</h2>
                <p>供佛灯油、酥油灯、油品支出</p>
            </a>

            <a class="card card-expense" href="{{ url_for('finance.expense', category='电费') }}">
                <h2>电费</h2>
                <p>TNB 电费记录</p>
            </a>

            <a class="card card-expense" href="{{ url_for('finance.expense', category='水费') }}">
                <h2>水费</h2>
                <p>水费记录</p>
            </a>

            <a class="card card-expense" href="{{ url_for('finance.expense', category='其它支出') }}">
                <h2>其它支出</h2>
                <p>其它杂项费用</p>
            </a>
        </div>
    </div>
    """, today_ym=today_ym)

@finance_bp.route("/monthly_fee/<branch>", methods=["GET", "POST"])
def monthly_fee(branch):

    branch = branch.upper()

    if branch not in ["CHE", "STW"]:
        return "Invalid branch", 400

    message = ""
    member = None
    paid_until = None
    raw_member_id = ""

    default_month_from = ""
    default_month_to = ""
    next_receipt_no = ""

    if request.method == "POST":

        action = request.form.get("action", "save")
        raw_member_id = request.form.get("member_id", "").strip()

        if raw_member_id.isdigit():
            member_id = f"{branch}-{int(raw_member_id)}"
        else:
            member_id = normalize_member_id(raw_member_id, default_branch=branch)

        member = db_query("""
            select *
            from members
            where member_id = %s
            limit 1
        """, (member_id,), fetchone=True)

        if not member:
            message = "找不到这个会员编号"
        else:
            paid = db_query("""
                select max(end_month) as paid_until
                from member_payments
                where member_id = %s
            """, (member_id,), fetchone=True)

            paid_until_date = paid["paid_until"] if paid else None
            paid_until = date_to_ym(paid_until_date)

            default_month_from = next_month_ym(paid_until_date)
            default_month_to = default_month_from

            last_receipt = db_query("""
                select receipt_no
                from finance_records
                where category = '月费'
                and receipt_no like %s
                order by receipt_no desc
                limit 1
            """, (branch + "%",), fetchone=True)

            if last_receipt and last_receipt["receipt_no"]:
                old_no = last_receipt["receipt_no"]
                prefix = old_no[:3]
                number = int(old_no[3:])
                next_receipt_no = prefix + str(number + 1).zfill(len(old_no) - 3)
            else:
                next_receipt_no = f"{branch}0000001"

            if action == "save":
                amount = money(request.form.get("amount"))
                receipt_no = request.form.get("receipt_no", "").strip().upper()
                payment_method = request.form.get("payment_method", "现金")
                month_from = request.form.get("month_from", "").strip()
                month_to = request.form.get("month_to", "").strip()
                remarks = request.form.get("remarks", "").strip()

                if not receipt_no.startswith(branch):
                    message = f"收条号码必须以 {branch} 开头"
                else:
                    months = (
                        (int(month_to[:4]) * 12 + int(month_to[5:7]))
                        -
                        (int(month_from[:4]) * 12 + int(month_from[5:7]))
                        + 1
                    )

                    month_from_db = month_from + "-01" if month_from and len(month_from) == 7 else month_from
                    month_to_db = month_to + "-01" if month_to and len(month_to) == 7 else month_to

                    existing = db_query("""
                        select id
                        from member_payments
                        where receipt_no = %s
                        limit 1
                    """, (receipt_no,), fetchone=True)

                    if existing:
                        message = "这个收条号码已经记录过了，请检查是否重复输入"
                    else:
                        db_query("""
                            insert into finance_records
                            (record_type, fund_account, record_date, category, receipt_no, member_id, name, phone,
                             amount, payment_method, month_from, month_to, remarks)
                            values
                            (%s, %s, %s, '月费', %s, %s, %s, %s,
                             %s, %s, %s, %s, %s)
                        """, (
                            "income",
                            get_fund_account("月费"),
                            date.today(),
                            receipt_no,
                            member["member_id"],
                            member.get("姓名") or member.get("name"),
                            member.get("电话号码") or member.get("phone"),
                            amount,
                            payment_method,
                            month_from,
                            month_to,
                            remarks
                        ))

                        db_query("""
                            insert into member_payments
                            (
                                payment_date,
                                member_id,
                                name,
                                receipt_no,
                                amount,
                                start_month,
                                end_month,
                                month_count
                            )
                            values
                            (
                                %s, %s, %s, %s, %s, %s, %s, %s
                            )
                        """, (
                            date.today(),
                            member["member_id"],
                            member.get("姓名") or member.get("name"),
                            receipt_no,
                            amount,
                            month_from_db,
                            month_to_db,
                            months
                        ))

                        return redirect(url_for("finance.records"))

    return render_template_string(FINANCE_STYLE + """
    <h1>{{ branch }} 月费收款</h1>

    {% if message %}
        <p style="color:red;">{{ message }}</p>
    {% endif %}

    <form method="post">
        <input type="hidden" name="action" value="check">

        <p>
            月费编号：
            <input name="member_id" value="{{ raw_member_id }}" placeholder="例如 {{ branch }}-108 / 108" required>
            <button type="submit">查询会员</button>
        </p>
    </form>

    {% if member %}
        <hr>

        <h2>会员资料</h2>

        <p>编号：{{ member.member_id }}</p>
        <p>姓名：{{ member.get("姓名") or member.get("name") }}</p>
        <p>电话：{{ member.get("电话号码") or member.get("phone") }}</p>
        <p>目前已缴费至：{{ paid_until or "没有记录" }}</p>

        <hr>

        <form method="post">
            <input type="hidden" name="action" value="save">
            <input type="hidden" name="member_id" value="{{ member.member_id }}">

            <p>
                收条号码：
                <input name="receipt_no" value="{{ next_receipt_no }}" required>
                <br>
                <small style="color:#666;">
                    系统会自动建议号码；如果换新收条簿，可以手动修改。
                </small>
            </p>

            <p>金额 RM：<input name="amount" type="number" step="50" min="0" required></p>

            <p>
                开始月份：
                <input id="month_from" name="month_from" value="{{ default_month_from }}" required>
            </p>

            <p>
                缴费至：
                <input id="month_to" name="month_to" value="{{ default_month_to }}" required>
            </p>

            <p id="month_hint" style="color:blue;"></p>

            <p>付款方式：
                <select name="payment_method">
                    <option>现金</option>
                    <option>银行过账</option>
                    <option>支票</option>
                </select>
            </p>

            <p>备注：<input name="remarks"></p>

            <button type="submit">保存月费</button>
        </form>

        <script>
        function addMonthsYM(ym, months) {
            let parts = ym.split("-");
            let y = parseInt(parts[0]);
            let m = parseInt(parts[1]);

            m += months;
            y += Math.floor((m - 1) / 12);
            m = ((m - 1) % 12) + 1;

            return y.toString().padStart(4, "0") + "-" + m.toString().padStart(2, "0");
        }

        function updateMonthTo() {
            let amountInput = document.querySelector('input[name="amount"]');
            let fromInput = document.getElementById("month_from");
            let toInput = document.getElementById("month_to");
            let hint = document.getElementById("month_hint");

            let amount = parseFloat(amountInput.value || "0");
            let months = Math.max(1, Math.round(amount / 50));

            if (fromInput.value && amount > 0) {
                toInput.value = addMonthsYM(fromInput.value, months - 1);
                hint.innerText = "系统判断：RM" + amount + " = " + months + "个月";
            }
        }

        document.querySelector('input[name="amount"]').addEventListener("input", updateMonthTo);
        document.getElementById("month_from").addEventListener("input", updateMonthTo);
        </script>
    {% endif %}

    <p><a href="{{ url_for('finance.finance_home') }}">返回财政首页</a></p>
    """,
    branch=branch,
    message=message,
    member=member,
    paid_until=paid_until,
    raw_member_id=raw_member_id,
    default_month_from=default_month_from,
    default_month_to=default_month_to,
    next_receipt_no=next_receipt_no
    )
    

@finance_bp.route("/income/<category>", methods=["GET", "POST"])
def normal_income(category):

    allowed_categories = ["财布施", "观音村", "膳食结缘"]

    if category not in allowed_categories:
        return "Invalid category", 400

    receipt_prefix = "CHE"

    last_receipt = db_query("""
        select receipt_no
        from finance_records
        where receipt_no like %s
        order by receipt_no desc
        limit 1
    """, (receipt_prefix + "%",), fetchone=True)

    if last_receipt and last_receipt["receipt_no"]:
        old_no = last_receipt["receipt_no"]
        number = int(old_no[3:])
        next_receipt_no = receipt_prefix + str(number + 1).zfill(len(old_no) - 3)
    else:
        next_receipt_no = "CHE0000001"

    message = ""

    if request.method == "POST":

        receipt_no = request.form.get("receipt_no", "").strip().upper()
        name = request.form.get("name", "").strip()
        phone = request.form.get("phone", "").strip()
        amount = money(request.form.get("amount"))
        payment_method = request.form.get("payment_method", "现金")
        remarks = request.form.get("remarks", "").strip()

        if not receipt_no.startswith("CHE"):
            message = "收条号码必须以 CHE 开头"
        else:
            existing = db_query("""
                select id
                from finance_records
                where receipt_no = %s
                limit 1
            """, (receipt_no,), fetchone=True)

            if existing:
                message = "这个收条号码已经记录过了，请检查是否重复输入"
            else:
                db_query("""
                    insert into finance_records
                    (
                        record_type,
                        fund_account,
                        record_date,
                        category,
                        receipt_no,
                        name,
                        phone,
                        amount,
                        payment_method,
                        remarks
                    )
                    values
                    (
                        %s, %s, %s, %s, %s,
                        %s, %s, %s, %s, %s
                    )
                """, (
                    "income",
                    get_fund_account(category),
                    date.today(),
                    category,
                    receipt_no,
                    name,
                    phone,
                    amount,
                    payment_method,
                    remarks
                ))

                return redirect(url_for("finance.records"))

    return render_template_string(FINANCE_STYLE + """
    <h1>{{ category }}</h1>

    {% if message %}
        <p style="color:red;">{{ message }}</p>
    {% endif %}

    <form method="post">

        <p>
            收条号码：
            <input
                name="receipt_no"
                value="{{ next_receipt_no }}"
                required
            >
            <br>
            <small style="color:#666;">
                系统会自动建议 CHE 编号；如果换新收条簿，可以手动修改。
            </small>
        </p>

        <p>
            姓名：
            <input name="name">
        </p>

        <p>
            电话：
            <input name="phone">
        </p>

        <p>
            金额 RM：
            <input
                name="amount"
                type="number"
                step="50"
                min="0"
                value="50"
                required
            >
        </p>

        <p>
            付款方式：
            <select name="payment_method">
                <option>现金</option>
                <option>银行过账</option>
                <option>支票</option>
            </select>
        </p>

        <p>
            备注：
            <input name="remarks">
        </p>

        <button type="submit">
            保存
        </button>

    </form>

    <p>
        <a href="{{ url_for('finance.finance_home') }}">
            返回财政首页
        </a>
    </p>

    """,
    category=category,
    next_receipt_no=next_receipt_no,
    message=message)

@finance_bp.route("/bank_pending/<int:pending_id>/delete", methods=["POST"])
def delete_bank_pending(pending_id):

    db_query("""
        delete from bank_pending_records
        where id = %s
        and status = 'pending'
    """, (pending_id,))

    return redirect(url_for("finance.bank_pending"))

@finance_bp.route("/bank_pending", methods=["GET", "POST"])
def bank_pending():

    message = ""

    if request.method == "POST":

        import re

        raw_text = request.form.get("raw_text", "").strip()

        member_id_raw = request.form.get("member_id", "").strip()
        name = request.form.get("name", "").strip()

        amount_raw = request.form.get("amount", "").strip()
        payment_date = request.form.get("payment_date") or date.today()

        bank_ref = request.form.get("bank_ref", "").strip()
        bank_name = request.form.get("bank_name", "").strip()

        category = request.form.get("category", "月费")
        remarks = request.form.get("remarks", "").strip()

        # 如果有粘贴 receipt 文字，先尝试自动解析
        if raw_text:

            # 抓会员编号：CHE-108 / CHE108 / STW-108 / STW108
            member_match = re.search(r"(CHE|STW)[-\s]?\d+", raw_text, re.IGNORECASE)

            if member_match and not member_id_raw:
                member_id_raw = member_match.group(0).replace(" ", "-").upper()

            # 抓金额：RM50 / RM 50.00 / MYR50.00
            amount_match = re.search(
                r"(RM|MYR)\s*([0-9]+(?:\.[0-9]{1,2})?)",
                raw_text,
                re.IGNORECASE
            )

            if amount_match and not amount_raw:
                amount_raw = amount_match.group(2)

            # 抓 Reference
            ref_match = re.search(
                r"(Reference|Ref|Transaction|DuitNow)\s*(No|ID|Number)?[:\s#-]*([A-Za-z0-9\-]+)",
                raw_text,
                re.IGNORECASE
            )

            if ref_match and not bank_ref:
                bank_ref = ref_match.group(3)

            # 抓银行名
            for b in ["Maybank", "Public Bank", "CIMB", "Hong Leong", "RHB", "AmBank", "Bank Islam", "BSN"]:
                if b.lower() in raw_text.lower() and not bank_name:
                    bank_name = b
                    break

            # 如果备注为空，把原始 receipt 放进备注，方便以后追查
            if not remarks:
                remarks = raw_text[:1000]

        member_id = normalize_member_id(member_id_raw) if member_id_raw else ""

        amount = money(amount_raw)

        member = None

        if category == "月费" and member_id:
            member = db_query("""
                select *
                from members
                where member_id = %s
                limit 1
            """, (member_id,), fetchone=True)

            if member:
                name = member.get("姓名") or member.get("name")

        existing_ref = db_query("""
            select id
            from bank_pending_records
            where bank_ref = %s
            and coalesce(bank_ref,'') <> ''
            limit 1
        """, (bank_ref,), fetchone=True)

        if existing_ref:
            message = "这个 Bank Reference 已经存在，请检查是否重复导入"

        db_query("""
            insert into bank_pending_records
            (
                member_id,
                name,
                amount,
                payment_date,
                bank_ref,
                bank_name,
                category,
                remarks
            )
            values
            (
                %s, %s, %s, %s,
                %s, %s, %s, %s
            )
        """, (
            member_id,
            name,
            amount,
            payment_date,
            bank_ref,
            bank_name,
            category,
            remarks
        ))

        return redirect(url_for("finance.bank_pending"))
    
    summary = db_query("""
        select
            count(*) as cnt,
            coalesce(sum(amount),0) as total
        from bank_pending_records
        where status = 'pending'
    """, fetchone=True)

    rows = db_query("""
        select *
        from bank_pending_records
        where status = 'pending'
        order by upload_date desc
    """, fetchall=True)

    return render_template_string(FINANCE_STYLE + """
<h1>🏦 银行过账中心</h1>

{% if message %}
    <p style="color:red;">{{ message }}</p>
{% endif %}

<h2>新增待确认记录</h2>

<form method="post">

    <p>
        项目：
        <select name="category">
            <option>月费</option>
            <option>财布施</option>
            <option>观音村</option>
            <option>膳食结缘</option>
        </select>
    </p>

    <p>
        粘贴 WhatsApp / 银行 Receipt 文字：
    </p>

    <textarea
        name="raw_text"
        rows="8"
        style="width:100%;max-width:760px;"
        placeholder="可以直接粘贴佛友 send 来的 receipt 文字。系统会尝试识别金额、Reference、会员编号。"
    ></textarea>

    <hr>

    <h3>手动补充 / 修正资料</h3>

    <p>
        会员编号：
        <input name="member_id" placeholder="例如 CHE-108 / STW-108 / 108">
        <small>月费建议填写会员编号，系统会自动找会员姓名。</small>
    </p>

    <p>
        姓名：
        <input name="name" placeholder="非月费或找不到会员时填写">
    </p>

    <p>
        金额 RM：
        <input name="amount" type="number" step="50.00" min="50.00" placeholder="例如 50.00">
    </p>

    <p>
        付款日期：
        <input name="payment_date" type="date">
    </p>

    <p>
        银行 Reference：
        <input name="bank_ref">
    </p>

    <p>
        银行：
        <select name="bank_name">
            <option value="">请选择</option>
            <option>Maybank</option>
            <option>Public Bank</option>
            <option>CIMB</option>
            <option>RHB</option>
            <option>Hong Leong</option>
            <option>AmBank</option>
            <option>Bank Islam</option>
        </select>
        
    </p>

    <p>
        备注：
        <input name="remarks" style="width:500px;">
    </p>

    <button type="submit">加入待确认</button>

</form>

<hr>

<h2>待确认列表</h2>

<p style="font-size:16px;">
    待确认笔数：
    <b>{{ summary.cnt }}</b>

    &nbsp;&nbsp;&nbsp;

    待确认总额：
    <b>RM {{ "%.2f"|format(summary.total) }}</b>
</p>

<table border="1" cellpadding="6">
    <tr>
        <th>付款日期</th>
        <th>编号</th>
        <th>姓名</th>
        <th>项目</th>
        <th>金额</th>
        <th>Reference</th>
        <th>银行</th>
        <th>备注</th>
        <th>操作</th>
    </tr>

    {% for r in rows %}
    <tr>
        <td>{{ r.payment_date }}</td>
        <td>{{ r.member_id or "-" }}</td>
        <td>{{ r.name or "-" }}</td>
        <td>{{ r.category }}</td>
        <td style="color:green;font-weight:bold;">
            RM {{ "%.2f"|format(r.amount or 0) }}
        </td>
        <td>{{ r.bank_ref or "-" }}</td>
        <td>{{ r.bank_name or "-" }}</td>
        <td>{{ r.remarks or "-" }}</td>
        <td>
            <form
                method="post"
                action="{{ url_for('finance.confirm_bank', pending_id=r.id) }}"
                style="display:inline;"
                onsubmit="return confirm('确定确认入账？');"
            >
                <button type="submit">确认入账</button>
            </form>

            <form
                method="post"
                action="{{ url_for('finance.delete_bank_pending', pending_id=r.id) }}"
                style="display:inline;"
                onsubmit="return confirm('确定删除这笔待确认记录？');"
            >
                <button type="submit" style="background:red;color:white;">
                    删除
                </button>
            </form>
        </td>
    </tr>
    {% endfor %}
</table>

<p>
    <a href="{{ url_for('finance.finance_home') }}">返回财政首页</a>
</p>
""",
rows=rows,
summary=summary,
message=message
)

@finance_bp.route("/bank_pending/<int:pending_id>/confirm", methods=["POST"])
def confirm_bank(pending_id):

    p = db_query("""
        select *
        from bank_pending_records
        where id = %s and status = 'pending'
    """, (pending_id,), fetchone=True)

    if not p:
        return redirect(url_for("finance.bank_pending"))

    member = None
    paid_until_date = None
    month_from = ""
    month_to = ""
    month_from_db = None
    month_to_db = None
    month_count = None

    if p["category"] == "月费" and p["member_id"]:

        member = db_query("""
            select *
            from members
            where member_id = %s
            limit 1
        """, (p["member_id"],), fetchone=True)

        paid = db_query("""
            select max(end_month) as paid_until
            from member_payments
            where member_id = %s
        """, (p["member_id"],), fetchone=True)

        paid_until_date = paid["paid_until"] if paid else None

        month_from = next_month_ym(paid_until_date)

        month_count = max(1, round(float(p["amount"] or 0) / 50))
        month_to = add_months_ym(month_from, month_count - 1)

        month_from_db = month_from + "-01"
        month_to_db = month_to + "-01"

    db_query("""
        insert into finance_records
        (
            record_type,
            fund_account,
            record_date,
            category,
            member_id,
            name,
            amount,
            payment_method,
            bank_name,
            bank_ref,
            month_from,
            month_to,
            remarks
        )
        values
        (
            %s, %s, %s, %s, %s, %s, %s,
            '银行过账', %s, %s,
            %s, %s, %s
        )
    """, (
        "income",
        get_fund_account(p["category"], "income"),
        p["payment_date"],
        p["category"],
        p["member_id"],
        p["name"],
        p["amount"],
        p["bank_name"],
        p["bank_ref"],
        month_from,
        month_to,
        p["remarks"]
    ))

    if p["category"] == "月费" and p["member_id"]:
        db_query("""
            insert into member_payments
            (
                payment_date,
                member_id,
                name,
                receipt_no,
                amount,
                start_month,
                end_month,
                month_count
            )
            values
            (
                %s, %s, %s, %s, %s, %s, %s, %s
            )
        """, (
            p["payment_date"],
            p["member_id"],
            p["name"],
            p["bank_ref"],
            p["amount"],
            month_from_db,
            month_to_db,
            month_count
        ))

    db_query("""
        update bank_pending_records
        set status = 'confirmed'
        where id = %s
    """, (pending_id,))

    return redirect(url_for("finance.records"))


@finance_bp.route("/records")
def records():

    q = request.args.get("q", "").strip()

    if q:
        keyword = f"%{q}%"
        rows = db_query("""
            select *
            from finance_records
            where
                receipt_no ilike %s
                or member_id ilike %s
                or name ilike %s
                or category ilike %s
                or payment_method ilike %s
                or bank_ref ilike %s
            order by record_date desc, id desc
            limit 300
        """, (
            keyword, keyword, keyword,
            keyword, keyword, keyword
        ), fetchall=True)
    else:
        rows = db_query("""
            select *
            from finance_records
            order by record_date desc, id desc
            limit 300
        """, fetchall=True)

    return render_template_string(FINANCE_STYLE + """
    <h1>财政记录</h1>

    <form method="get">
        <input name="q" value="{{ q }}" placeholder="搜索：收条 / 编号 / 姓名 / 项目 / Reference">
        <button type="submit">搜索</button>
        <a href="{{ url_for('finance.records') }}">清除</a>
    </form>

    <br>

    <table border="1" cellpadding="6">
        <tr>
            <th>日期</th>
            <th>项目</th>
            <th>收条</th>
            <th>编号</th>
            <th>姓名</th>
            <th>金额</th>
            <th>方式</th>
            <th>Reference</th>
            <th>月份</th>
            <th>备注</th>
            <th>操作</th>
        </tr>

        {% for r in rows %}
        <tr>
            <td>{{ r.record_date }}</td>
            <td>{{ r.category }}</td>
            <td>{{ r.receipt_no }}</td>
            <td>{{ r.member_id }}</td>
            <td>{{ r.name }}</td>
            <td>RM {{ "%.2f"|format(r.amount or 0) }}</td>
            <td>{{ r.payment_method }}</td>
            <td>{{ r.bank_ref }}</td>
            <td>
                {% if r.month_from or r.month_to %}
                    {{ r.month_from }} - {{ r.month_to }}
                {% else %}
                    -
                {% endif %}
            </td>
            <td>{{ r.remarks }}</td>
            <td>
                <form method="post"
                    action="{{ url_for('finance.delete_record', record_id=r.id) }}"
                    onsubmit="return confirm('确定要删除这笔财政记录吗？');">
                    <button type="submit">删除</button>
                </form>
            </td>
        </tr>
        {% endfor %}
    </table>

    {% if not rows %}
        <p style="color:red;">没有找到记录。</p>
    {% endif %}

    <p><a href="{{ url_for('finance.finance_home') }}">返回财政首页</a></p>
    """, rows=rows, q=q)

@finance_bp.route("/dashboard")
def dashboard():

    ym = request.args.get("ym", date.today().strftime("%Y-%m"))

    daily_income = db_query("""
        select
            category,
            coalesce(sum(amount), 0) as total
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
          and fund_account = '观音堂日常户口'
          and record_type = 'income'
        group by category
        order by category
    """, (ym,), fetchall=True)

    daily_expense = db_query("""
        select
            category,
            coalesce(sum(amount), 0) as total
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
          and fund_account = '观音堂日常户口'
          and record_type = 'expense'
        group by category
        order by category
    """, (ym,), fetchall=True)

    hq_income = db_query("""
        select
            category,
            coalesce(sum(amount), 0) as total
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
          and fund_account = '总会户口'
          and record_type = 'income'
        group by category
        order by category
    """, (ym,), fetchall=True)

    daily_income_total = db_query("""
        select coalesce(sum(amount), 0) as total
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
          and fund_account = '观音堂日常户口'
          and record_type = 'income'
    """, (ym,), fetchone=True)

    daily_expense_total = db_query("""
        select coalesce(sum(amount), 0) as total
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
          and fund_account = '观音堂日常户口'
          and record_type = 'expense'
    """, (ym,), fetchone=True)

    hq_income_total = db_query("""
        select coalesce(sum(amount), 0) as total
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
          and fund_account = '总会户口'
          and record_type = 'income'
    """, (ym,), fetchone=True)

    opening_balance = float(
        request.args.get("opening_balance") or 0
    )

    daily_balance = (
        opening_balance
        + float(daily_income_total["total"] or 0)
        - float(daily_expense_total["total"] or 0)
    )
    return render_template_string(FINANCE_STYLE + """
    <h1>财政统计 Dashboard</h1>

    <form method="get">
        <label>月份：</label>
        <input name="ym" value="{{ ym }}" placeholder="2026-06">
        <button type="submit">查看</button>
    </form>
                                  
    <label>上月结余 RM：</label>
    <input name="opening_balance" value="{{ opening_balance }}" placeholder="例如 15000">

    <hr>

    <h2>观音堂日常户口</h2>

    <h3>收入</h3>
    <table border="1" cellpadding="8">
        <tr>
            <th>项目</th>
            <th>金额 RM</th>
        </tr>

        {% for r in daily_income %}
        <tr>
            <td>{{ r.category }}</td>
            <td>{{ "%.2f"|format(r.total) }}</td>
        </tr>
        {% endfor %}

        <tr>
            <th>日常户口总收入</th>
            <th>{{ "%.2f"|format(daily_income_total.total) }}</th>
        </tr>
    </table>

    <h3>支出</h3>
    <table border="1" cellpadding="8">
        <tr>
            <th>项目</th>
            <th>金额 RM</th>
        </tr>

        {% for r in daily_expense %}
        <tr>
            <td>{{ r.category }}</td>
            <td>{{ "%.2f"|format(r.total) }}</td>
        </tr>
        {% endfor %}

        <tr>
            <th>日常户口总支出</th>
            <th>{{ "%.2f"|format(daily_expense_total.total) }}</th>
        </tr>

        <tr>
            <th>日常户口本月结余</th>
            <th>{{ "%.2f"|format(daily_balance) }}</th>
        </tr>
    </table>

    <hr>

    <h2>总会户口</h2>

    <table border="1" cellpadding="8">
        <tr>
            <th>项目</th>
            <th>金额 RM</th>
        </tr>

        {% for r in hq_income %}
        <tr>
            <td>{{ r.category }}</td>
            <td>{{ "%.2f"|format(r.total) }}</td>
        </tr>
        {% endfor %}

        <tr>
            <th>总会户口总收入</th>
            <th>{{ "%.2f"|format(hq_income_total.total) }}</th>
        </tr>
    </table>

    <p>
        
        <a href="/finance/export_monthly_report?ym={{ ym }}">
            下载专业版 Excel 月报
        </a>
    </p>

    <p><a href="{{ url_for('finance.finance_home') }}">返回财政首页</a></p>
    """,
    ym=ym,
    opening_balance=opening_balance,
    daily_income=daily_income,
    daily_expense=daily_expense,
    hq_income=hq_income,
    daily_income_total=daily_income_total,
    daily_expense_total=daily_expense_total,
    hq_income_total=hq_income_total,
    daily_balance=daily_balance)



@finance_bp.route("/expense/<category>", methods=["GET", "POST"])
def expense(category):
    next_receipt_no = get_next_receipt_no_by_category(category)

    if request.method == "POST":
        record_date = request.form.get("record_date") or date.today()
        amount = money(request.form.get("amount"))
        payment_method = request.form.get("payment_method", "现金")
        remarks = request.form.get("remarks", "").strip()

        db_query("""
            insert into finance_records
            (
                record_type,
                record_date,
                category,
                amount,
                payment_method,
                fund_account,
                remarks
            )
            values
            (
                'expense',
                %s, %s, %s, %s, %s, %s
            )
        """, (
            record_date,
            category,
            amount,
            payment_method,
            remarks
        ))

        return redirect(url_for("finance.records"))

    return render_template_string(FINANCE_STYLE + """
    <h1>支出记录：{{ category }}</h1>

    <form method="post">

        <p>
            日期：
            <input name="record_date" type="date">
        </p>

        <p>
            金额 RM：
            <input name="amount" type="number" step="0.01" required>
        </p>

        <p>
            付款方式：
            <select name="payment_method">
                <option>现金</option>
                <option>银行过账</option>
                <option>支票</option>
            </select>
        </p>

        <p>
            备注：
            <input name="remarks">
        </p>

        <button type="submit">保存支出</button>

    </form>

    <p>
        <a href="{{ url_for('finance.finance_home') }}">
            返回财政首页
        </a>
    </p>

    """, category=category, next_receipt_no=next_receipt_no)

@finance_bp.route("/export_monthly_report")
def export_monthly_report():
    ym = request.args.get("ym", date.today().strftime("%Y-%m"))

    rows = db_query("""
        select
            record_date,
            receipt_no,
            category,
            record_type,
            fund_account,
            name,
            phone,
            amount,
            '' as note
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
        order by record_date, id
    """, (ym,), fetchall=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "财政摘要"
    ws_income = wb.create_sheet("收入明细")
    ws_expense = wb.create_sheet("支出明细")

    # =========================
    # 样式
    # =========================
    green = "1F5E20"
    light_green = "EAF4E4"
    blue = "0B3A78"
    light_blue = "EAF2FF"
    gold = "F8E7B8"
    dark_text = "1F2937"
    red = "D00000"

    thin_green = Side(style="thin", color="8DBA7C")
    thin_blue = Side(style="thin", color="7FA6D9")
    thin_gold = Side(style="thin", color="C9962C")
    thin_gray = Side(style="thin", color="CCCCCC")

    border_green = Border(left=thin_green, right=thin_green, top=thin_green, bottom=thin_green)
    border_blue = Border(left=thin_blue, right=thin_blue, top=thin_blue, bottom=thin_blue)
    border_gold = Border(left=thin_gold, right=thin_gold, top=thin_gold, bottom=thin_gold)
    border_gray = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    money_fmt = '"RM"#,##0.00'

    def money(v):
        return float(v or 0)

    def set_range_border(ws, cell_range, border):
        for row in ws[cell_range]:
            for cell in row:
                cell.border = border

    def fill_range(ws, cell_range, fill):
        for row in ws[cell_range]:
            for cell in row:
                cell.fill = fill

    def center_range(ws, cell_range):
        for row in ws[cell_range]:
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    def format_money_cell(cell):
        cell.number_format = money_fmt
        cell.alignment = Alignment(horizontal="center", vertical="center")

    def auto_width(ws):
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_len + 4, 22)

    # =========================
    # 数据整理
    # =========================
    summary = {
        "观音堂日常户口": {"income": 0, "expense": 0},
        "总会户口": {"income": 0, "expense": 0},
    }

    income_rows = []
    expense_rows = []

    for r in rows:
        fund = r["fund_account"] or "未分类"
        record_type = r["record_type"] or "income"
        amount = money(r["amount"])

        if fund not in summary:
            summary[fund] = {"income": 0, "expense": 0}

        summary[fund][record_type] += amount

        if record_type == "income":
            income_rows.append(r)
        else:
            expense_rows.append(r)

    daily_income = summary.get("观音堂日常户口", {}).get("income", 0)
    daily_expense = summary.get("观音堂日常户口", {}).get("expense", 0)
    daily_balance = daily_income - daily_expense

    hq_income = summary.get("总会户口", {}).get("income", 0)
    hq_expense = summary.get("总会户口", {}).get("expense", 0)
    hq_balance = hq_income - hq_expense

    total_income = sum(v["income"] for v in summary.values())
    total_expense = sum(v["expense"] for v in summary.values())
    total_balance = total_income - total_expense

    # =========================
    # Sheet 1 财政摘要：图片同款
    # =========================
    ws.sheet_view.showGridLines = False

    for col in range(1, 16):
        ws.column_dimensions[get_column_letter(col)].width = 14

    ws.row_dimensions[1].height = 14
    ws.row_dimensions[2].height = 34
    ws.row_dimensions[3].height = 26
    ws.row_dimensions[4].height = 12

    ws.merge_cells("A1:O4")
    ws["A1"] = f"观音堂财政月报\n{ym}"
    ws["A1"].font = Font(bold=True, size=24, color=dark_text)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["A1"].fill = PatternFill("solid", fgColor="FFF8E8")
    set_range_border(ws, "A1:O4", border_gold)

    # 左表：观音堂日常户口
    ws.merge_cells("A5:G5")
    ws["A5"] = "观音堂日常户口"
    ws["A5"].fill = PatternFill("solid", fgColor=green)
    ws["A5"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A5"].alignment = Alignment(horizontal="center")

    daily_table = [
        ["项目", "收入（RM）", "支出（RM）", "本月结余（RM）"],
        ["收入", daily_income, "-", daily_income],
        ["支出", "-", daily_expense, -daily_expense],
        ["本月结余", "-", "-", daily_balance],
    ]

    start_row = 6
    for i, row in enumerate(daily_table, start_row):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=2)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=4)
        ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=5)
        ws.merge_cells(start_row=i, start_column=6, end_row=i, end_column=7)

        ws.cell(i, 1).value = row[0]
        ws.cell(i, 3).value = row[1]
        ws.cell(i, 5).value = row[2]
        ws.cell(i, 6).value = row[3]

        for c in [1, 3, 5, 6]:
            ws.cell(i, c).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(i, c).border = border_green

        if i == 6:
            fill_range(ws, f"A{i}:G{i}", PatternFill("solid", fgColor=light_green))
            for c in [1, 3, 5, 6]:
                ws.cell(i, c).font = Font(bold=True)
        elif i == 9:
            fill_range(ws, f"A{i}:G{i}", PatternFill("solid", fgColor="DDEED2"))
            ws.cell(i, 6).font = Font(bold=True, size=14, color=green)

    for cell_ref in ["C7", "E8", "F7", "F8", "F9"]:
        ws[cell_ref].number_format = money_fmt

    # 右表：总会户口
    ws.merge_cells("I5:O5")
    ws["I5"] = "总会户口"
    ws["I5"].fill = PatternFill("solid", fgColor=blue)
    ws["I5"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["I5"].alignment = Alignment(horizontal="center")

    hq_table = [
        ["项目", "收入（RM）", "支出（RM）", "本月结余（RM）"],
        ["收入", hq_income, "-", hq_income],
        ["支出", "-", hq_expense, -hq_expense],
        ["本月结余", "-", "-", hq_balance],
    ]

    for i, row in enumerate(hq_table, start_row):
        ws.merge_cells(start_row=i, start_column=9, end_row=i, end_column=10)
        ws.merge_cells(start_row=i, start_column=11, end_row=i, end_column=12)
        ws.merge_cells(start_row=i, start_column=13, end_row=i, end_column=13)
        ws.merge_cells(start_row=i, start_column=14, end_row=i, end_column=15)

        ws.cell(i, 9).value = row[0]
        ws.cell(i, 11).value = row[1]
        ws.cell(i, 13).value = row[2]
        ws.cell(i, 14).value = row[3]

        for c in [9, 11, 13, 14]:
            ws.cell(i, c).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(i, c).border = border_blue

        if i == 6:
            fill_range(ws, f"I{i}:O{i}", PatternFill("solid", fgColor=light_blue))
            for c in [9, 11, 13, 14]:
                ws.cell(i, c).font = Font(bold=True, color=blue)
        elif i == 9:
            fill_range(ws, f"I{i}:O{i}", PatternFill("solid", fgColor="DCEAFF"))
            ws.cell(i, 14).font = Font(bold=True, size=14, color=blue)

    for cell_ref in ["K7", "M8", "N7", "N8", "N9"]:
        ws[cell_ref].number_format = money_fmt

    # 总计横条
    ws.merge_cells("A11:C12")
    ws["A11"] = "总计"
    ws["A11"].font = Font(bold=True, size=16, color="7A5200")
    ws["A11"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A11"].fill = PatternFill("solid", fgColor="FFF2CC")

    ws.merge_cells("D11:F12")
    ws["D11"] = f"总收入（RM）\n{total_income}"
    ws["D11"].font = Font(bold=True, size=12, color=green)
    ws["D11"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["D11"].fill = PatternFill("solid", fgColor="FFF8E8")

    ws.merge_cells("G11:I12")
    ws["G11"] = f"总支出（RM）\n{total_expense}"
    ws["G11"].font = Font(bold=True, size=12, color=red)
    ws["G11"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["G11"].fill = PatternFill("solid", fgColor="FFF8E8")

    ws.merge_cells("J11:O12")
    ws["J11"] = f"总体结余（RM）\n{total_balance}"
    ws["J11"].font = Font(bold=True, size=12, color=blue)
    ws["J11"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["J11"].fill = PatternFill("solid", fgColor="FFF8E8")

    set_range_border(ws, "A11:O12", border_gold)

    # 收入明细表
    ws.merge_cells("A14:G14")
    ws["A14"] = "收入明细"
    ws["A14"].fill = PatternFill("solid", fgColor=green)
    ws["A14"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A14"].alignment = Alignment(horizontal="center")

    income_headers = ["日期", "收条号码", "分类", "户口", "姓名", "电话", "金额（RM）"]
    for col, h in enumerate(income_headers, 1):
        c = ws.cell(15, col)
        c.value = h
        c.fill = PatternFill("solid", fgColor=light_green)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
        c.border = border_green

    rno = 16
    for r in income_rows[:8]:
        vals = [
            r["record_date"],
            r["receipt_no"],
            r["category"],
            r["fund_account"],
            r["name"],
            r["phone"],
            money(r["amount"]),
        ]
        for col, v in enumerate(vals, 1):
            c = ws.cell(rno, col)
            c.value = v
            c.border = border_green
            c.alignment = Alignment(horizontal="center")
            if col == 7:
                c.number_format = money_fmt
        rno += 1

    while rno <= 23:
        for col in range(1, 8):
            ws.cell(rno, col).border = border_green
            ws.cell(rno, col).fill = PatternFill("solid", fgColor="FAFCF7")
        rno += 1

    ws.merge_cells("A24:F24")
    ws["A24"] = "收入总计"
    ws["A24"].fill = PatternFill("solid", fgColor="DDEED2")
    ws["A24"].font = Font(bold=True, color=green)
    ws["A24"].alignment = Alignment(horizontal="right")
    ws["G24"] = total_income
    ws["G24"].number_format = money_fmt
    ws["G24"].fill = PatternFill("solid", fgColor="DDEED2")
    ws["G24"].font = Font(bold=True, color=green)
    ws["G24"].alignment = Alignment(horizontal="center")
    set_range_border(ws, "A24:G24", border_green)

    # 支出明细表
    ws.merge_cells("I14:O14")
    ws["I14"] = "支出明细"
    ws["I14"].fill = PatternFill("solid", fgColor=blue)
    ws["I14"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["I14"].alignment = Alignment(horizontal="center")

    expense_headers = ["日期", "收条号码", "分类", "户口", "姓名", "电话", "金额（RM）"]
    for col, h in enumerate(expense_headers, 9):
        c = ws.cell(15, col)
        c.value = h
        c.fill = PatternFill("solid", fgColor=light_blue)
        c.font = Font(bold=True, color=blue)
        c.alignment = Alignment(horizontal="center")
        c.border = border_blue

    rno = 16
    for r in expense_rows[:8]:
        vals = [
            r["record_date"],
            r["receipt_no"],
            r["category"],
            r["fund_account"],
            r["name"],
            r["phone"],
            money(r["amount"]),
        ]
        for col, v in enumerate(vals, 9):
            c = ws.cell(rno, col)
            c.value = v
            c.border = border_blue
            c.alignment = Alignment(horizontal="center")
            if col == 15:
                c.number_format = money_fmt
        rno += 1

    while rno <= 23:
        for col in range(9, 16):
            ws.cell(rno, col).border = border_blue
            ws.cell(rno, col).fill = PatternFill("solid", fgColor="F8FBFF")
        rno += 1

    ws.merge_cells("I24:N24")
    ws["I24"] = "支出总计"
    ws["I24"].fill = PatternFill("solid", fgColor="DCEAFF")
    ws["I24"].font = Font(bold=True, color=blue)
    ws["I24"].alignment = Alignment(horizontal="right")
    ws["O24"] = total_expense
    ws["O24"].number_format = money_fmt
    ws["O24"].fill = PatternFill("solid", fgColor="DCEAFF")
    ws["O24"].font = Font(bold=True, color=red)
    ws["O24"].alignment = Alignment(horizontal="center")
    set_range_border(ws, "I24:O24", border_blue)

    ws.freeze_panes = "A14"

    # =========================
    # Sheet 2 / Sheet 3 明细
    # =========================
    def write_detail_sheet(ws2, title, data_rows, main_color, light_color):
        ws2.sheet_view.showGridLines = False

        ws2.merge_cells("A1:I2")
        ws2["A1"] = f"{title} - {ym}"
        ws2["A1"].font = Font(bold=True, size=18, color="FFFFFF")
        ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws2["A1"].fill = PatternFill("solid", fgColor=main_color)

        headers = ["日期", "收条号码", "分类", "户口", "姓名", "电话", "金额（RM）", "备注", "类型"]

        for col, h in enumerate(headers, 1):
            c = ws2.cell(4, col)
            c.value = h
            c.fill = PatternFill("solid", fgColor=light_color)
            c.font = Font(bold=True, color=dark_text)
            c.alignment = Alignment(horizontal="center")
            c.border = border_gray

        row_no = 5
        for r in data_rows:
            vals = [
                r["record_date"],
                r["receipt_no"],
                r["category"],
                r["fund_account"],
                r["name"],
                r["phone"],
                money(r["amount"]),
                r["note"],
                "收入" if r["record_type"] == "income" else "支出",
            ]

            for col, v in enumerate(vals, 1):
                c = ws2.cell(row_no, col)
                c.value = v
                c.border = border_gray
                c.alignment = Alignment(horizontal="center")
                if col == 7:
                    c.number_format = money_fmt

            row_no += 1

        total_row = row_no + 1
        ws2.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=6)
        ws2.cell(total_row, 1).value = "总计"
        ws2.cell(total_row, 1).font = Font(bold=True)
        ws2.cell(total_row, 1).alignment = Alignment(horizontal="right")
        ws2.cell(total_row, 7).value = sum(money(r["amount"]) for r in data_rows)
        ws2.cell(total_row, 7).number_format = money_fmt
        ws2.cell(total_row, 7).font = Font(bold=True, color=main_color)

        fill_range(ws2, f"A{total_row}:I{total_row}", PatternFill("solid", fgColor=light_color))
        set_range_border(ws2, f"A{total_row}:I{total_row}", border_gray)

        ws2.freeze_panes = "A5"
        auto_width(ws2)

    write_detail_sheet(ws_income, "收入明细", income_rows, green, light_green)
    write_detail_sheet(ws_expense, "支出明细", expense_rows, blue, light_blue)

    # =========================
    # 输出
    # =========================
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"观音堂财政月报_{ym}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

