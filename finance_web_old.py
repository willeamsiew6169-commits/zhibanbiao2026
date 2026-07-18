# finance_web.py

import os
import re
import tempfile

from copy import copy
from io import BytesIO
from db import db_query
from finance_common import (
    money,
    normalize_finance_date,
    get_finance_ym,
    get_month_close_record,
    is_finance_month_closed,
    get_month_lock_message,
    require_finance_month_open,
    get_current_finance_user,
)
from finance_audit import write_finance_audit
from openpyxl import Workbook, load_workbook
from datetime import date, datetime
from flask import send_file, request, flash
from utils import normalize_member_id
from psycopg2.extras import RealDictCursor
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from flask import Blueprint, request, redirect, url_for, render_template_string, send_file, session, jsonify


EXPENSE_SUB_CATEGORY_OPTIONS = {
    "供花": ["鲜花"],
    "供果": ["水果", "供果用品", "其它"],
    "供油": ["灯油"],
    "佛台用品": ["佛具", "佛台用品", "佛堂用品", "其它"],
    "电费": ["TNB 20-1", "TNB 20-2", "其它"],
    "水费": [
        "Air Selangor 20-1",
        "Air Selangor 20-2",
        "Indah Water 20-1",
        "Indah Water 20-2",
        "其它",
    ],
    "电话及网络费": [
        "Digi Reload",
        "Celcom Reload",
        "Unifi Internet",
        "Maxis",
        "其它",
    ],
    "维修保养": [
        "Air-con Service",
        "电器维修",
        "水管维修",
        "建筑维修",
        "其它",
    ],
    "装修": ["GYT装修", "观音堂装修", "活动中心装修", "其它"],
    "日常采购": ["文具", "厨房用品", "清洁用品", "日常用品", "其它"],
    "执照及行政费": ["License Fee", "政府费用", "行政费用", "其它"],
    "其它支出": ["印刷", "交通", "杂项", "其它"],
}

EXPENSE_VENDOR_BY_SUB_CATEGORY = {
    # 供花
    "鲜花": "Pudu Ria Florist Trading Sdn Bhd",
    "供花用品": "Pudu Ria Florist Trading Sdn Bhd",

    # 供油
    "灯油": "Laohongka Sdn Bhd",
    "油灯用品": "Laohongka Sdn Bhd",

    # 电费
    "TNB 20-1": "Tenaga Nasional",
    "TNB 20-2": "Tenaga Nasional",

    # 水费
    "Air Selangor 20-1": "Air Selangor",
    "Air Selangor 20-2": "Air Selangor",
    "Indah Water 20-1": "Indah Water",
    "Indah Water 20-2": "Indah Water",

    # 电话及网络
    "Digi Reload": "Digi",
    "Celcom Reload": "Celcom",
    "Unifi Internet": "TM Unifi",
    "Maxis": "Maxis",
}

AUTO_VENDOR_CATEGORIES = {
    "供花",
    "供油",
    "电费",
    "水费",
    "电话及网络费",
}
# =========================================================
# Finance Web local helpers
# These remain here because they are domain-specific to
# daily finance input rather than generic shared utilities.
# =========================================================

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


def next_month_ym(d):
    if not d:
        return date.today().strftime("%Y-%m")

    return add_months_ym(
        d.strftime("%Y-%m"),
        1
    )


def normalize_receipt_category(category: str) -> str:

    category = str(category or "").strip()

    mapping = {
        "月费": "月费",
        "财布施": "财布施",
        "观音村": "观音村",

        "膳食": "膳食结缘",
        "膳食结缘": "膳食结缘",
        "初一十五": "膳食结缘",
        "初一十五膳食结缘": "膳食结缘",
    }

    return mapping.get(category, category)


def get_next_receipt_no(
    branch: str,
    category: str,
) -> str:

    branch = str(branch or "CHE").strip().upper()
    category = normalize_receipt_category(category)

    row = db_query(
        """
        select
            prefix,
            current_number,
            number_width
        from finance_receipt_books
        where branch = %s
          and category = %s
        limit 1
        """,
        (
            branch,
            category,
        ),
        fetchone=True,
    )

    if not row:
        raise ValueError(
            f"找不到收条簿设置：{branch} / {category}"
        )

    next_number = int(
        row.get("current_number") or 0
    ) + 1

    prefix = str(
        row.get("prefix") or ""
    )

    width = int(
        row.get("number_width") or 6
    )

    return f"{prefix}{next_number:0{width}d}"


def update_receipt_book_number(
    branch: str,
    category: str,
    receipt_no: str,
):

    branch = str(branch or "CHE").strip().upper()
    category = normalize_receipt_category(category)
    receipt_no = str(receipt_no or "").strip().upper()

    digits = "".join(
        ch for ch in receipt_no
        if ch.isdigit()
    )

    if not digits:
        return

    number = int(digits)

    db_query(
        """
        update finance_receipt_books
        set
            current_number = greatest(
                current_number,
                %s
            ),
            updated_at = now()
        where branch = %s
          and category = %s
        """,
        (
            number,
            branch,
            category,
        ),
    )


def get_active_finance_vendors():
    return db_query("""
        select
            id,
            vendor_name,
            sort_order
        from finance_vendors
        where is_active = true
        order by
            sort_order,
            vendor_name
    """, fetchall=True)


finance_bp = Blueprint("finance", __name__, url_prefix="/finance")

FINANCE_PIN = "123456"

SPECIAL_DONATION_TITLE = "813交流会财布施"

INCOME_CATEGORIES = [
    "月费",
    "财布施",
    "观音村",
    "膳食结缘",
    "观音堂纯檀香布施",
    SPECIAL_DONATION_TITLE,
    "临时特别布施"
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

FINANCE_V5_STYLE = """
<style>
*{
    box-sizing:border-box;
}

body{
    margin:0;
    background:#f4f7fb;
    color:#1f2937;
    font-family:"Microsoft JhengHei","Microsoft YaHei",Arial,sans-serif;
}

.finance-v5{
    max-width:900px;
    margin:0 auto;
    padding:28px 20px 50px;
}

.v5-header{
    background:linear-gradient(135deg,#1769aa,#2387c9);
    color:#fff;
    padding:28px;
    border-radius:22px;
    margin-bottom:24px;
    box-shadow:0 8px 24px rgba(23,105,170,.20);
}

.v5-header h1{
    margin:0;
    font-size:36px;
}

.v5-header p{
    margin:8px 0 0;
    font-size:17px;
    opacity:.9;
}

.v5-topbar{
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:12px;
    margin-bottom:18px;
}

.v5-back,
.v5-logout{
    display:inline-flex;
    align-items:center;
    justify-content:center;
    min-height:44px;
    padding:10px 18px;
    border-radius:12px;
    text-decoration:none;
    font-weight:bold;
    font-size:17px;
}

.v5-back{
    color:#1769aa;
    background:#eaf4ff;
}

.v5-logout{
    color:#fff;
    background:#d94343;
}

.v5-menu-grid{
    display:grid;
    grid-template-columns:repeat(2,minmax(0,1fr));
    gap:16px;
}

.v5-menu-btn{
    display:flex;
    align-items:center;
    gap:16px;
    min-height:94px;
    padding:20px;
    background:#fff;
    border:1px solid #e5e7eb;
    border-radius:18px;
    text-decoration:none;
    color:#1f2937;
    box-shadow:0 4px 16px rgba(0,0,0,.06);
    transition:.18s ease;
}

.v5-menu-btn:hover{
    transform:translateY(-2px);
    box-shadow:0 8px 22px rgba(0,0,0,.10);
}

.v5-icon{
    width:54px;
    height:54px;
    flex:0 0 54px;
    display:flex;
    align-items:center;
    justify-content:center;
    border-radius:15px;
    font-size:30px;
    background:#eef6ff;
}

.v5-menu-text{
    min-width:0;
}

.v5-menu-title{
    font-size:22px;
    font-weight:800;
    margin-bottom:5px;
}

.v5-menu-desc{
    color:#667085;
    font-size:15px;
    line-height:1.45;
}

.v5-section-title{
    margin:26px 0 12px;
    font-size:18px;
    font-weight:800;
    color:#475467;
}

.v5-alert{
    background:#fff6e5;
    border-left:6px solid #f59e0b;
}

.v5-income{
    background:#edf9f1;
}

.v5-expense{
    background:#fff1f2;
}

.v5-member{
    background:#eef5ff;
}

.v5-report{
    background:#f7f2ff;
}

@media(max-width:680px){
    .finance-v5{
        padding:16px 12px 35px;
    }

    .v5-header{
        padding:22px 18px;
        border-radius:17px;
    }

    .v5-header h1{
        font-size:29px;
    }

    .v5-menu-grid{
        grid-template-columns:1fr;
        gap:12px;
    }

    .v5-menu-btn{
        min-height:82px;
        padding:16px;
    }

    .v5-menu-title{
        font-size:20px;
    }
}
</style>
"""

FINANCE_DATE_COMPONENT = """
<style>
.finance-date-control{
    display:grid;
    gap:9px;
    width:100%;
}

.finance-date-main{
    display:grid;
    grid-template-columns:52px minmax(0,1fr) 52px;
    gap:9px;
    align-items:stretch;
}

.finance-date-main .form-input,
.finance-date-main input[type="date"]{
    width:100%;
    min-width:0;
    margin:0;
    text-align:center;
    font-weight:800;
}

.finance-date-step,
.finance-date-quick{
    border:1px solid #cbd5e1;
    background:#f8fafc;
    color:#334155;
    border-radius:10px;
    cursor:pointer;
    font-weight:800;
    transition:.15s ease;
}

.finance-date-step{
    min-height:52px;
    font-size:25px;
}

.finance-date-step:hover,
.finance-date-quick:hover{
    background:#eaf4ff;
    border-color:#8fbce0;
    color:#1769aa;
}

.finance-date-step:active,
.finance-date-quick:active{
    transform:scale(.97);
}

.finance-date-shortcuts{
    display:grid;
    grid-template-columns:repeat(3,minmax(0,1fr));
    gap:8px;
}

.finance-date-quick{
    min-height:42px;
    padding:8px 10px;
    font-size:15px;
}

.finance-date-status{
    min-height:22px;
    font-size:14px;
    font-weight:700;
    line-height:1.45;
}

.finance-date-today{
    border-color:#86efac !important;
    background:#f0fdf4 !important;
    color:#166534 !important;
}

.finance-date-yesterday{
    border-color:#fde68a !important;
    background:#fffbeb !important;
    color:#92400e !important;
}

.finance-date-old{
    border-color:#fdba74 !important;
    background:#fff7ed !important;
    color:#9a3412 !important;
}

.finance-date-future{
    border-color:#fca5a5 !important;
    background:#fef2f2 !important;
    color:#b91c1c !important;
}

@media(max-width:560px){
    .finance-date-main{
        grid-template-columns:48px minmax(0,1fr) 48px;
    }

    .finance-date-shortcuts{
        grid-template-columns:1fr;
    }
}
</style>

<script>
(function(){
    function localISO(dateValue){
        const year = dateValue.getFullYear();
        const month = String(dateValue.getMonth() + 1).padStart(2, "0");
        const day = String(dateValue.getDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    }

    function parseISO(value){
        if(!value){
            return null;
        }

        const parts = value.split("-").map(Number);

        if(parts.length !== 3 || parts.some(Number.isNaN)){
            return null;
        }

        return new Date(parts[0], parts[1] - 1, parts[2]);
    }

    function startOfDay(value){
        return new Date(
            value.getFullYear(),
            value.getMonth(),
            value.getDate()
        );
    }

    function changeDate(input, days){
        let current = parseISO(input.value);

        if(!current){
            current = new Date();
        }

        current.setDate(current.getDate() + days);
        input.value = localISO(current);
        input.dispatchEvent(new Event("change", {bubbles:true}));
    }

    function setRelativeDate(input, days){
        const selected = new Date();
        selected.setDate(selected.getDate() + days);
        input.value = localISO(selected);
        input.dispatchEvent(new Event("change", {bubbles:true}));
    }

    function openCalendar(input){
        if(typeof input.showPicker === "function"){
            try{
                input.showPicker();
                return;
            }catch(error){}
        }

        input.focus();
        input.click();
    }

    function updateStatus(input, status){
        input.classList.remove(
            "finance-date-today",
            "finance-date-yesterday",
            "finance-date-old",
            "finance-date-future"
        );

        const selected = parseISO(input.value);

        if(!selected){
            status.textContent = "";
            return;
        }

        const today = startOfDay(new Date());
        const chosen = startOfDay(selected);
        const difference = Math.round(
            (chosen - today) / 86400000
        );

        if(difference === 0){
            input.classList.add("finance-date-today");
            status.textContent = "🟢 今天";
        }else if(difference === -1){
            input.classList.add("finance-date-yesterday");
            status.textContent = "🟡 昨天";
        }else if(difference > 0){
            input.classList.add("finance-date-future");
            status.textContent = `🔴 未来日期（${difference} 天后），请检查`;
        }else if(difference < -30){
            input.classList.add("finance-date-old");
            status.textContent = `🟠 ${Math.abs(difference)} 天前，可能是旧账补录`;
        }else{
            status.textContent = `📅 ${Math.abs(difference)} 天前`;
        }
    }

    function enhanceDateInput(input){
        if(input.dataset.financeDateReady === "1"){
            return;
        }

        input.dataset.financeDateReady = "1";

        const control = document.createElement("div");
        control.className = "finance-date-control";

        const main = document.createElement("div");
        main.className = "finance-date-main";

        const previous = document.createElement("button");
        previous.type = "button";
        previous.className = "finance-date-step";
        previous.textContent = "◀";
        previous.title = "前一天";

        const next = document.createElement("button");
        next.type = "button";
        next.className = "finance-date-step";
        next.textContent = "▶";
        next.title = "后一天";

        const shortcuts = document.createElement("div");
        shortcuts.className = "finance-date-shortcuts";

        const yesterday = document.createElement("button");
        yesterday.type = "button";
        yesterday.className = "finance-date-quick";
        yesterday.textContent = "昨天";

        const today = document.createElement("button");
        today.type = "button";
        today.className = "finance-date-quick";
        today.textContent = "今天";

        const calendar = document.createElement("button");
        calendar.type = "button";
        calendar.className = "finance-date-quick";
        calendar.textContent = "📅 选日期";

        const status = document.createElement("div");
        status.className = "finance-date-status";

        const parent = input.parentNode;
        parent.insertBefore(control, input);

        main.appendChild(previous);
        main.appendChild(input);
        main.appendChild(next);

        shortcuts.appendChild(yesterday);
        shortcuts.appendChild(today);
        shortcuts.appendChild(calendar);

        control.appendChild(main);
        control.appendChild(shortcuts);
        control.appendChild(status);

        previous.addEventListener("click", function(){
            changeDate(input, -1);
        });

        next.addEventListener("click", function(){
            changeDate(input, 1);
        });

        yesterday.addEventListener("click", function(){
            setRelativeDate(input, -1);
        });

        today.addEventListener("click", function(){
            setRelativeDate(input, 0);
        });

        calendar.addEventListener("click", function(){
            openCalendar(input);
        });

        input.addEventListener("change", function(){
            updateStatus(input, status);
        });

        input.addEventListener("keydown", function(event){
            if(event.altKey || event.ctrlKey || event.metaKey){
                return;
            }

            if(event.key === "ArrowLeft"){
                event.preventDefault();
                changeDate(input, -1);
            }else if(event.key === "ArrowRight"){
                event.preventDefault();
                changeDate(input, 1);
            }
        });

        updateStatus(input, status);
    }

    function initFinanceDates(root){
        const selector = [
            'input[type="date"][name="receipt_date"]',
            'input[type="date"][name="record_date"]',
            'input[type="date"][name="payment_date"]',
            'input[type="date"][name="voucher_date"]',
            'input[type="date"][name="payment_voucher_date"]'
        ].join(",");

        (root || document).querySelectorAll(selector).forEach(
            enhanceDateInput
        );
    }

    document.addEventListener("DOMContentLoaded", function(){
        initFinanceDates(document);
    });

    window.initFinanceDates = initFinanceDates;
})();
</script>
"""


# Finance shared helpers are imported from finance_common.py.

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

    error = ""

    if request.method == "POST":

        pin = request.form.get("pin", "").strip()

        if pin == FINANCE_PIN:
            session["finance_login"] = True
            return redirect(url_for("finance.finance_home"))

        error = "财政 PIN 不正确，请重新输入。"

    return render_template_string(FINANCE_DATE_COMPONENT + """
<!doctype html>
<html lang="zh">
<head>
    <meta charset="utf-8">

    <meta
        name="viewport"
        content="width=device-width, initial-scale=1"
    >

    <title>财政系统登入</title>

    <link
        rel="stylesheet"
        href="{{ url_for('static', filename='css/toolbox.css') }}"
    >

    <style>
        .finance-login-page{
            max-width:680px;
        }

        .finance-login-card{
            margin-top:22px;
        }

        .finance-login-icon{
            font-size:52px;
            text-align:center;
            margin-bottom:10px;
        }

        .finance-login-title{
            text-align:center;
            margin-bottom:8px;
        }

        .finance-login-subtitle{
            text-align:center;
            color:#667085;
            margin-bottom:26px;
        }

        .finance-pin-input{
            text-align:center;
            font-size:28px;
            letter-spacing:8px;
        }

        .finance-login-actions{
            display:grid;
            gap:12px;
            margin-top:22px;
        }

        .finance-login-actions .btn-tool{
            width:100%;
        }

        @media(max-width:600px){
            .finance-login-page{
                padding:16px;
            }

            .finance-login-icon{
                font-size:44px;
            }
        }
    </style>
</head>

<body>

<div class="page finance-login-page">

    <div class="card finance-login-card">

        <div class="finance-login-icon">
            🏦
        </div>

        <h1 class="page-title finance-login-title">
            财政系统
        </h1>

        <p class="page-subtitle finance-login-subtitle">
            Finance Management V5
        </p>

        {% if error %}
            <div class="alert alert-danger">
                {{ error }}
            </div>
        {% endif %}

        <form method="post">

            <div class="form-group">

                <label class="form-label">
                    财政 PIN
                </label>

                <input
                    class="form-input finance-pin-input"
                    name="pin"
                    type="password"
                    inputmode="numeric"
                    autocomplete="new-password"
                    maxlength="8"
                    autofocus
                    required
                >

            </div>

            <div class="finance-login-actions">

                <button
                    class="btn-tool btn-primary"
                    type="submit"
                >
                    进入财政系统
                </button>

                <a
                    class="btn-tool btn-secondary"
                    href="/admin-home"
                >
                    返回管理员首页
                </a>

            </div>

        </form>

    </div>

</div>

</body>
</html>
""",
    error=error
    )

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
            m.remark,
            max(p.end_month) as paid_until,
            max(p.payment_date) as last_payment_date
        from members m
        left join member_payments p
            on p.member_id = m.member_id
           and coalesce(p.status, 'active') = 'active'
        where coalesce(m.member_status, m.status, '') not in (
            '停供',
            '停止',
            '永久停止',
            '往生',
            '已往生'
        )
        group by
            m.member_id,
            m.name,
            m.phone,
            m.remark
        having
            max(p.end_month) is not null
            and max(p.end_month) < date_trunc(
                'month',
                current_date
            )
        order by
            paid_until asc,
            m.member_id
    """, fetchall=True)

    today = date.today()

    current_month_index = (
        today.year * 12
        + today.month
    )

    green_count = 0
    yellow_count = 0
    red_count = 0
    total_amount = 0

    for r in rows:

        paid_until = r["paid_until"]

        paid_index = (
            paid_until.year * 12
            + paid_until.month
        )

        late_months = max(
            current_month_index - paid_index,
            0
        )

        reference_amount = late_months * 50

        r["late_months"] = late_months
        r["reference_amount"] = reference_amount

        if late_months <= 2:
            r["level"] = "green"
            r["level_text"] = "最近停止"
            green_count += 1

        elif late_months <= 6:
            r["level"] = "yellow"
            r["level_text"] = "一段时间未缴费"
            yellow_count += 1

        else:
            r["level"] = "red"
            r["level_text"] = "较久未缴费"
            red_count += 1

        total_amount += reference_amount

        phone = (
            r["phone"]
            or ""
        ).strip()

        phone_digits = "".join(
            ch
            for ch in phone
            if ch.isdigit()
        )

        if phone_digits.startswith("0"):
            phone_digits = "6" + phone_digits

        if phone_digits:
            r["wa_link"] = (
                "https://wa.me/"
                + phone_digits
            )
        else:
            r["wa_link"] = ""

    return render_template_string(FINANCE_DATE_COMPONENT + """
    <!doctype html>
    <html lang="zh">
    <head>

        <meta charset="utf-8">

        <meta
            name="viewport"
            content="width=device-width, initial-scale=1"
        >

        <title>月费关怀名单</title>

        <link
            rel="stylesheet"
            href="{{ url_for(
                'static',
                filename='css/toolbox.css'
            ) }}"
        >

        <style>

            .care-page {
                max-width: 1450px;
            }

            .care-header {
                background:
                    linear-gradient(
                        135deg,
                        #2563eb,
                        #1d4ed8
                    );

                color: white;
                padding: 28px;
                border-radius: 22px;
                margin-bottom: 20px;

                box-shadow:
                    0 12px 30px
                    rgba(37, 99, 235, 0.18);
            }

            .care-header h1 {
                margin: 0 0 8px;
                font-size: 30px;
            }

            .care-header p {
                margin: 0;
                opacity: 0.92;
                line-height: 1.6;
            }

            .care-note {
                background: #eff6ff;
                border: 1px solid #bfdbfe;
                color: #1e40af;
                border-radius: 14px;
                padding: 14px 16px;
                margin-bottom: 20px;
                line-height: 1.65;
            }

            .care-summary {
                display: grid;
                grid-template-columns:
                    repeat(5, minmax(0, 1fr));
                gap: 16px;
                margin-bottom: 20px;
            }

            .care-summary .summary-box {
                min-height: 120px;
                text-align: center;

                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
            }

            .summary-icon {
                font-size: 29px;
                margin-bottom: 6px;
            }

            .summary-label {
                color: #64748b;
                font-size: 15px;
                margin-bottom: 6px;
            }

            .summary-value {
                color: #0f172a;
                font-size: 26px;
                font-weight: 800;
            }

            .summary-green {
                background: #f0fdf4;
                border-color: #bbf7d0;
            }

            .summary-yellow {
                background: #fffbeb;
                border-color: #fde68a;
            }

            .summary-red {
                background: #fef2f2;
                border-color: #fecaca;
            }

            .summary-green .summary-value {
                color: #15803d;
            }

            .summary-yellow .summary-value {
                color: #a16207;
            }

            .summary-red .summary-value {
                color: #b91c1c;
            }

            .table-topbar {
                display: flex;
                justify-content: space-between;
                align-items: center;
                gap: 12px;
                flex-wrap: wrap;
                margin-bottom: 14px;
            }

            .result-text {
                color: #64748b;
                font-size: 16px;
            }

            .care-table {
                min-width: 1280px;
            }

            .care-table td {
                vertical-align: middle;
            }

            .member-id {
                color: #1d4ed8;
                font-weight: 800;
                white-space: nowrap;
            }

            .member-name {
                color: #1e293b;
                font-weight: 700;
                min-width: 90px;
            }

            .phone-cell {
                white-space: nowrap;
            }

            .money-cell {
                color: #15803d;
                font-weight: 800;
                white-space: nowrap;
                text-align: right;
            }

            .date-cell {
                white-space: nowrap;
            }

            .remark-cell {
                min-width: 180px;
                max-width: 320px;
                white-space: normal;
                line-height: 1.5;
                overflow-wrap: anywhere;
            }

            .level-row-green {
                background: #f0fdf4;
            }

            .level-row-yellow {
                background: #fffbeb;
            }

            .level-row-red {
                background: #fef2f2;
            }

            .level-badge {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                padding: 7px 11px;
                border-radius: 999px;
                font-size: 14px;
                font-weight: 800;
                white-space: nowrap;
            }

            .level-green {
                background: #dcfce7;
                color: #166534;
            }

            .level-yellow {
                background: #fef3c7;
                color: #92400e;
            }

            .level-red {
                background: #fee2e2;
                color: #991b1b;
            }

            .wa-button {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                gap: 5px;
                padding: 8px 12px;
                border-radius: 10px;
                background: #dcfce7;
                color: #166534;
                font-weight: 800;
                text-decoration: none;
                white-space: nowrap;
                border: 1px solid #bbf7d0;
            }

            .wa-button:hover {
                background: #bbf7d0;
            }

            .empty-care {
                text-align: center;
                padding: 55px 20px;
                color: #64748b;
            }

            .empty-care-icon {
                font-size: 48px;
                margin-bottom: 10px;
            }

            .empty-care h3 {
                color: #334155;
                margin: 0 0 8px;
            }

            .bottom-actions {
                margin-top: 20px;
            }

            @media (max-width: 1100px) {

                .care-summary {
                    grid-template-columns:
                        repeat(2, minmax(0, 1fr));
                }

            }

            @media (max-width: 700px) {

                .care-page {
                    padding-left: 12px;
                    padding-right: 12px;
                }

                .care-header {
                    padding: 22px 18px;
                    border-radius: 18px;
                }

                .care-header h1 {
                    font-size: 26px;
                }

                .care-summary {
                    grid-template-columns: 1fr;
                }

                .care-summary .summary-box {
                    min-height: 100px;
                }

                .bottom-actions .btn-tool {
                    width: 100%;
                }

            }

        </style>

    </head>

    <body>

    <div class="page care-page">

        <div class="care-header">

            <h1>🌿 月费关怀名单</h1>

            <p>
                协助负责人了解会员最近缴费情况，
                方便适时联络与关怀。
            </p>

        </div>

        <div class="care-note">

            ℹ️ 此名单不是追缴名单。
            “参考金额”只按照每月 RM50 计算，
            方便负责人了解大概月份，
            不代表会员必须补缴该金额。

        </div>

        <div class="care-summary">

            <div class="summary-box">

                <div class="summary-icon">
                    👥
                </div>

                <div class="summary-label">
                    总人数
                </div>

                <div class="summary-value">
                    {{ rows|length }}
                </div>

            </div>

            <div class="summary-box">

                <div class="summary-icon">
                    💰
                </div>

                <div class="summary-label">
                    参考金额
                </div>

                <div class="summary-value">
                    RM {{ "{:,.2f}".format(total_amount) }}
                </div>

            </div>

            <div class="summary-box summary-green">

                <div class="summary-icon">
                    🟢
                </div>

                <div class="summary-label">
                    最近停止
                </div>

                <div class="summary-value">
                    {{ green_count }}
                </div>

            </div>

            <div class="summary-box summary-yellow">

                <div class="summary-icon">
                    🟡
                </div>

                <div class="summary-label">
                    一段时间未缴费
                </div>

                <div class="summary-value">
                    {{ yellow_count }}
                </div>

            </div>

            <div class="summary-box summary-red">

                <div class="summary-icon">
                    🔴
                </div>

                <div class="summary-label">
                    较久未缴费
                </div>

                <div class="summary-value">
                    {{ red_count }}
                </div>

            </div>

        </div>

        <div class="card">

            <div class="table-topbar">

                <div
                    class="section-title"
                    style="margin-bottom:0;"
                >
                    📋 会员关怀资料
                </div>

                <div class="result-text">
                    共
                    <strong>{{ rows|length }}</strong>
                    位会员
                </div>

            </div>

            {% if rows %}

                <div class="table-responsive">

                    <table class="record-table care-table">

                        <thead>

                            <tr>
                                <th>会员编号</th>
                                <th>姓名</th>
                                <th>电话</th>
                                <th>WhatsApp</th>
                                <th>已缴至</th>
                                <th>间隔月份</th>
                                <th>参考金额</th>
                                <th>关怀状态</th>
                                <th>最后付款日期</th>
                                <th>备注</th>
                            </tr>

                        </thead>

                        <tbody>

                            {% for r in rows %}

                                <tr class="
                                    {% if r.level == 'green' %}
                                        level-row-green
                                    {% elif r.level == 'yellow' %}
                                        level-row-yellow
                                    {% else %}
                                        level-row-red
                                    {% endif %}
                                ">

                                    <td>
                                        <span class="member-id">
                                            {{ r.member_id }}
                                        </span>
                                    </td>

                                    <td>
                                        <span class="member-name">
                                            {{ r.name }}
                                        </span>
                                    </td>

                                    <td class="phone-cell">
                                        {{ r.phone or "-" }}
                                    </td>

                                    <td>

                                        {% if r.wa_link %}

                                            <a
                                                class="wa-button"
                                                href="{{ r.wa_link }}"
                                                target="_blank"
                                                rel="noopener noreferrer"
                                            >
                                                💬 打开
                                            </a>

                                        {% else %}

                                            -

                                        {% endif %}

                                    </td>

                                    <td class="date-cell">
                                        {{ r.paid_until.strftime("%Y-%m") }}
                                    </td>

                                    <td>
                                        {{ r.late_months }} 个月
                                    </td>

                                    <td class="money-cell">
                                        RM
                                        {{ "%.2f"|format(
                                            r.reference_amount
                                        ) }}
                                    </td>

                                    <td>

                                        <span class="
                                            level-badge
                                            {% if r.level == 'green' %}
                                                level-green
                                            {% elif r.level == 'yellow' %}
                                                level-yellow
                                            {% else %}
                                                level-red
                                            {% endif %}
                                        ">

                                            {% if r.level == "green" %}
                                                🟢
                                            {% elif r.level == "yellow" %}
                                                🟡
                                            {% else %}
                                                🔴
                                            {% endif %}

                                            {{ r.level_text }}

                                        </span>

                                    </td>

                                    <td class="date-cell">

                                        {% if r.last_payment_date %}
                                            {{ r.last_payment_date }}
                                        {% else %}
                                            -
                                        {% endif %}

                                    </td>

                                    <td class="remark-cell">
                                        {{ r.remark or "-" }}
                                    </td>

                                </tr>

                            {% endfor %}

                        </tbody>

                    </table>

                </div>

            {% else %}

                <div class="empty-care">

                    <div class="empty-care-icon">
                        🌿
                    </div>

                    <h3>目前没有需要关怀的会员</h3>

                    <p>
                        所有在册会员的月费记录都在当前月份内。
                    </p>

                </div>

            {% endif %}

        </div>

        <div class="bottom-actions">

            <a
                class="btn-tool btn-secondary"
                href="{{ url_for(
                    'finance.finance_home'
                ) }}"
            >
                ← 返回财政首页
            </a>

        </div>

    </div>

    </body>
    </html>
    """,
        rows=rows,
        green_count=green_count,
        yellow_count=yellow_count,
        red_count=red_count,
        total_amount=total_amount
    )
    
    
@finance_bp.route("/member/<member_id>/edit",
                  methods=["GET","POST"])
def edit_member(member_id):

    member = db_query("""
        select *
        from members
        where member_id = %s
        limit 1
    """, (member_id,), fetchone=True)

    if not member:
        return "Member not found"

    if request.method == "POST":

        status = request.form.get("status")

        db_query("""
            update members
            set status = %s
            where member_id = %s
        """, (
            status,
            member_id
        ))

        return redirect(
            url_for("finance.member_management")
        )

    return render_template_string(FINANCE_STYLE + FINANCE_DATE_COMPONENT + """

    <h1>{{ member.member_id }}</h1>

    <p>姓名：{{ member.name }}</p>

    <p>电话：{{ member.phone }}</p>

    <form method="post">

        <p>

            状态：

            <select name="status">

                <option
                {% if member.status == '在册' %}
                selected
                {% endif %}
                >
                在册
                </option>

                <option
                {% if member.status == '停供' %}
                selected
                {% endif %}
                >
                停供
                </option>

                <option
                {% if member.status == '往生' %}
                selected
                {% endif %}
                >
                往生
                </option>

            </select>

        </p>

        <button type="submit">
            保存
        </button>

    </form>

    <p>
        <a href="{{ url_for('finance.member_management') }}">
            返回
        </a>
    </p>

    """, member=member)


@finance_bp.route("/")
def finance_home():

    if not session.get("finance_login"):
        return redirect(url_for("finance.finance_login"))

    return render_template_string(FINANCE_V5_STYLE + FINANCE_DATE_COMPONENT + """
    <div class="finance-v5">

        <div class="v5-topbar">
            <a class="v5-back" href="/admin-home">
                ← 管理员首页
            </a>

            <a class="v5-logout"
               href="{{ url_for('finance.finance_logout') }}">
                退出
            </a>
        </div>

        <div class="v5-header">
            <h1>💰 财政系统 V5</h1>
            <p>请选择要处理的工作</p>
        </div>

        <div class="v5-menu-grid">

            <a class="v5-menu-btn v5-income"
               href="{{ url_for('finance.finance_income_menu') }}">
                <div class="v5-icon">💵</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">收入录入</div>
                    <div class="v5-menu-desc">
                        月费、财布施、观音村、膳食结缘
                    </div>
                </div>
            </a>

            <a class="v5-menu-btn"
               href="{{ url_for('finance.bank_pending') }}">
                <div class="v5-icon">🏦</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">银行过账</div>
                    <div class="v5-menu-desc">
                        银行转账待确认与补开收条
                    </div>
                </div>
            </a>

            <a class="v5-menu-btn v5-expense"
                href="{{ url_for('finance.finance_expense_menu') }}">

                <div class="v5-icon">🧾</div>

                <div class="v5-menu-text">

                    <div class="v5-menu-title">
                        支出录入
                    </div>

                    <div class="v5-menu-desc">
                        供花、供果、供油、水电、电话网络及其它支出
                    </div>

                </div>

            </a>

            <a class="v5-menu-btn v5-member"
               href="{{ url_for('finance.finance_member_menu') }}">
                <div class="v5-icon">👥</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">会员与月费</div>
                    <div class="v5-menu-desc">
                        会员资料、月费情况与关怀名单
                    </div>
                </div>
            </a>

            <a class="v5-menu-btn v5-report"
               href="{{ url_for('finance.finance_report_menu') }}">
                <div class="v5-icon">📊</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">报表与查询</div>
                    <div class="v5-menu-desc">
                        Dashboard、财政记录及月报
                    </div>
                </div>
            </a>
                                  
            <a class="v5-menu-btn v5-report"
               href="/finance/reports/excel">

                <div class="v5-icon">📥</div>

                <div class="v5-menu-text">

                    <div class="v5-menu-title">
                        Excel 下载中心
                    </div>

                    <div class="v5-menu-desc">
                        下载月费、善款、观音村、膳食及 Petty Cash 报表
                    </div>

                </div>

            </a>

            <a class="v5-menu-btn v5-alert"
               href="{{ url_for('finance.late_members') }}">
                <div class="v5-icon">🌿</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">月费关怀名单</div>
                    <div class="v5-menu-desc">
                        查看供养间隔与会员月费资料
                    </div>
                </div>
            </a>
                                  
            <a class="v5-menu-btn v5-report"
                href="/finance/month_end">

                <div class="v5-icon">📒</div>

                <div class="v5-menu-text">

                    <div class="v5-menu-title">
                        财政月结
                    </div>

                    <div class="v5-menu-desc">

                        Cash Bank In、
                        对账及总会报表

                    </div>

                </div>

            </a>
                                  
            <a class="v5-menu-btn v5-report"
               href="{{ url_for('finance_import.import_home') }}">

                <div class="v5-icon">
                    🗂️
                </div>

                <div class="v5-menu-text">

                    <div class="v5-menu-title">
                        历史资料导入
                    </div>

                    <div class="v5-menu-desc">
                        导入月费、善款、Petty Cash 及历史 Excel
                    </div>

                </div>

            </a>

        </div>

    </div>
    """)

@finance_bp.route("/menu/income")
def finance_income_menu():

    return render_template_string(FINANCE_V5_STYLE + FINANCE_DATE_COMPONENT + """
    <div class="finance-v5">

        <div class="v5-topbar">
            <a class="v5-back"
               href="{{ url_for('finance.finance_home') }}">
                ← 返回财政首页
            </a>
        </div>

        <div class="v5-header">
            <h1>💵 收入录入</h1>
            <p>请选择收入项目</p>
        </div>

        <div class="v5-menu-grid">

            <a class="v5-menu-btn v5-member"
               href="{{ url_for('finance.monthly_fee_batch', branch='CHE') }}">
                <div class="v5-icon">C</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">CHE 月费</div>
                    <div class="v5-menu-desc">
                        批量登记 CHE 月费
                    </div>
                </div>
            </a>

            <a class="v5-menu-btn v5-member"
               href="{{ url_for('finance.monthly_fee_batch', branch='STW') }}">
                <div class="v5-icon">S</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">STW 月费</div>
                    <div class="v5-menu-desc">
                        批量登记 STW 月费
                    </div>
                </div>
            </a>

            <a class="v5-menu-btn v5-income"
               href="{{ url_for('finance.income_batch', category='财布施') }}">
                <div class="v5-icon">🙏</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">财布施</div>
                    <div class="v5-menu-desc">
                        批量登记财布施
                    </div>
                </div>
            </a>

            <a class="v5-menu-btn v5-income"
               href="{{ url_for('finance.income_batch', category='观音村') }}">
                <div class="v5-icon">🏡</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">观音村</div>
                    <div class="v5-menu-desc">
                        批量登记观音村收入
                    </div>
                </div>
            </a>

            <a class="v5-menu-btn v5-income"
               href="{{ url_for('finance.income_batch', category='膳食结缘') }}">
                <div class="v5-icon">🥗</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">膳食结缘</div>
                    <div class="v5-menu-desc">
                        批量登记膳食结缘收入
                    </div>
                </div>
            </a>
                                  
            <a class="v5-menu-btn v5-income"
            href="{{ url_for(
                'finance.income_batch',
                category='观音堂纯檀香布施'
            ) }}">
                <div class="v5-icon">🪵</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">
                        观音堂纯檀香布施
                    </div>
                    <div class="v5-menu-desc">
                        批量登记纯檀香相关布施
                    </div>
                </div>
            </a>

            <a class="v5-menu-btn v5-income"
            href="{{ url_for(
                'finance.income_batch',
                category=special_donation_title
            ) }}">
                <div class="v5-icon">🎊</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">
                        {{ special_donation_title }}
                    </div>
                    <div class="v5-menu-desc">
                        本年度特别活动布施
                    </div>
                </div>
            </a>

            <a class="v5-menu-btn v5-income"
            href="{{ url_for(
                'finance.income_batch',
                category='临时特别布施'
            ) }}">
                <div class="v5-icon">➕</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">
                        临时特别布施
                    </div>
                    <div class="v5-menu-desc">
                        突发活动或临时项目使用
                    </div>
                </div>
            </a>

        </div>

    </div>
    """,
    special_donation_title=SPECIAL_DONATION_TITLE
    )

@finance_bp.route("/menu/expense")
def finance_expense_menu():

    expense_items = [

        ("🌸", "供花", "鲜花及供花相关支出"),

        ("🍎", "供果", "水果及供品相关支出"),

        ("🪔", "供油", "灯油及油品相关支出"),

        ("🛕", "佛台用品", "佛具、佛台及佛堂用品"),

        ("⚡", "电费", "TNB 20-1、TNB 20-2"),

        ("💧", "水费", "Air Selangor、Indah Water"),

        ("📶", "电话及网络费", "Celcom、Digi、Unifi"),

        ("🛠️", "维修保养", "维修及保养费用"),

        ("🏗️", "装修", "GYT、观音堂及活动中心装修"),

        ("🛒", "日常采购", "文具、厨房及日常用品"),

        ("📄", "执照及行政费", "License、政府及行政费用"),

        ("🧾", "其它支出", "印刷、交通及其它杂项费用"),

    ]

    return render_template_string(
        FINANCE_V5_STYLE
        + FINANCE_DATE_COMPONENT
        + """
<div class="finance-v5">

    <div class="v5-topbar">

        <a class="v5-back"
           href="{{ url_for('finance.finance_home') }}">
            ← 返回财政首页
        </a>

    </div>

    <div class="v5-header">

        <h1>🧾 支出录入</h1>

        <p>请选择支出项目</p>

    </div>

    <div class="v5-menu-grid">

        {% for icon, category, desc in expense_items %}

        <a class="v5-menu-btn v5-expense"
           href="{{ url_for('finance.expense', category=category) }}">

            <div class="v5-icon">
                {{ icon }}
            </div>

            <div class="v5-menu-text">

                <div class="v5-menu-title">
                    {{ category }}
                </div>

                <div class="v5-menu-desc">
                    {{ desc }}
                </div>

            </div>

        </a>

        {% endfor %}

    </div>

</div>
""",
        expense_items=expense_items,
    )

@finance_bp.route("/menu/member")
def finance_member_menu():

    return render_template_string(FINANCE_V5_STYLE + FINANCE_DATE_COMPONENT + """
    <div class="finance-v5">

        <div class="v5-topbar">
            <a class="v5-back"
               href="{{ url_for('finance.finance_home') }}">
                ← 返回财政首页
            </a>
        </div>

        <div class="v5-header">
            <h1>👥 会员与月费</h1>
            <p>会员资料与月费情况</p>
        </div>

        <div class="v5-menu-grid">

            <a class="v5-menu-btn v5-member"
               href="{{ url_for('finance.member_management') }}">
                <div class="v5-icon">👤</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">会员管理</div>
                    <div class="v5-menu-desc">
                        搜索、查看和维护会员资料
                    </div>
                </div>
            </a>

            <a class="v5-menu-btn v5-alert"
               href="{{ url_for('finance.late_members') }}">
                <div class="v5-icon">🌿</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">月费关怀名单</div>
                    <div class="v5-menu-desc">
                        查看已缴至、供养间隔与参考资料
                    </div>
                </div>
            </a>

            <a class="v5-menu-btn v5-member"
               href="{{ url_for('finance.add_member') }}">
                <div class="v5-icon">➕</div>
                <div class="v5-menu-text">
                    <div class="v5-menu-title">新增会员</div>
                    <div class="v5-menu-desc">
                        建立新月费会员资料
                    </div>
                </div>
            </a>

        </div>

    </div>
    """)

@finance_bp.route("/menu/report")
def finance_report_menu():

    today_ym = date.today().strftime("%Y-%m")

    return render_template_string(
        FINANCE_V5_STYLE + FINANCE_DATE_COMPONENT + """
        <style>
            .v5-export-card{
                display:flex;
                align-items:center;
                gap:16px;
                padding:22px;
                border-radius:18px;
                background:#ffffff;
                box-shadow:0 4px 14px rgba(0,0,0,.08);
            }

            .v5-export-icon{
                font-size:42px;
                flex:0 0 auto;
            }

            .v5-export-content{
                flex:1;
                min-width:0;
            }

            .v5-export-title{
                font-size:24px;
                font-weight:700;
                margin-bottom:6px;
            }

            .v5-export-desc{
                font-size:17px;
                color:#666;
                margin-bottom:14px;
                line-height:1.5;
            }

            .v5-export-form{
                display:flex;
                align-items:center;
                gap:10px;
                flex-wrap:wrap;
            }

            .v5-month-input{
                min-height:48px;
                padding:8px 12px;
                border:2px solid #d6dbe3;
                border-radius:10px;
                font-size:19px;
                background:#fff;
            }

            .v5-download-btn{
                min-height:48px;
                padding:9px 18px;
                border:0;
                border-radius:10px;
                font-size:18px;
                font-weight:700;
                cursor:pointer;
                background:#198754;
                color:#fff;
            }

            .v5-download-btn:hover{
                filter:brightness(.95);
            }

            .v5-download-btn.report{
                background:#246bfd;
            }

            @media (max-width:600px){
                .v5-export-card{
                    align-items:flex-start;
                }

                .v5-export-form{
                    display:grid;
                    grid-template-columns:1fr;
                }

                .v5-month-input,
                .v5-download-btn{
                    width:100%;
                    box-sizing:border-box;
                }
            }
        </style>

        <div class="finance-v5">

            <div class="v5-topbar">
                <a class="v5-back"
                href="{{ url_for('finance.finance_home') }}">
                    ← 返回财政首页
                </a>
            </div>

            <div class="v5-header">
                <h1>📊 报表与查询</h1>
                <p>统计、记录与 Excel 月报</p>
            </div>

            <div class="v5-menu-grid">

                <a class="v5-menu-btn v5-report"
                href="{{ url_for('finance.dashboard') }}">
                    <div class="v5-icon">📈</div>

                    <div class="v5-menu-text">
                        <div class="v5-menu-title">
                            财政 Dashboard
                        </div>

                        <div class="v5-menu-desc">
                            查看每月收入、支出和户口统计
                        </div>
                    </div>
                </a>

                <a class="v5-menu-btn v5-report"
                href="{{ url_for('finance.records') }}">
                    <div class="v5-icon">🔎</div>

                    <div class="v5-menu-text">
                        <div class="v5-menu-title">
                            财政记录搜索
                        </div>

                        <div class="v5-menu-desc">
                            搜索收条、会员编号、姓名与银行 Reference
                        </div>
                    </div>
                </a>
                
                <!-- 专业版月报 -->
                <div class="v5-export-card">

                    <div class="v5-export-icon">
                        📊
                    </div>

                    <div class="v5-export-content">

                        <div class="v5-export-title">
                            下载专业版月报
                        </div>

                        <div class="v5-export-desc">
                            选择月份后，下载收入、支出及户口统计月报
                        </div>

                        <form
                            method="get"
                            action="{{ url_for('finance.export_monthly_report') }}"
                            class="v5-export-form"
                        >
                            <input
                                type="month"
                                name="ym"
                                value="{{ today_ym }}"
                                class="v5-month-input"
                                required
                            >

                            <button
                                type="submit"
                                class="v5-download-btn report"
                            >
                                📥 下载专业版月报
                            </button>
                        </form>

                    </div>
                </div>

            </div>
            
        </div>
        """,
        today_ym=today_ym,
    )

@finance_bp.route("/member_management")
def member_management():

    q = request.args.get("q", "").strip()

    if q:
        keyword = f"%{q}%"

        rows = db_query("""
            select
                member_id,
                name,
                english_name,
                phone,
                coalesce(status, '在册') as status,
                remark
            from members
            where
                member_id ilike %s
                or name ilike %s
                or english_name ilike %s
                or phone ilike %s
            order by member_id
            limit 300
        """, (
            keyword,
            keyword,
            keyword,
            keyword
        ), fetchall=True)

    else:
        rows = db_query("""
            select
                member_id,
                name,
                english_name,
                phone,
                coalesce(status, '在册') as status,
                remark
            from members
            order by member_id
            limit 300
        """, fetchall=True)

    total_members = len(rows)

    return render_template_string(FINANCE_DATE_COMPONENT + """
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">

        <meta
            name="viewport"
            content="width=device-width, initial-scale=1"
        >

        <title>月费会员管理</title>

        <link
            rel="stylesheet"
            href="{{ url_for('static', filename='css/toolbox.css') }}"
        >

        <style>
            .member-page {
                max-width: 1100px;
            }

            .member-header {
                background:
                    linear-gradient(
                        135deg,
                        #2563eb,
                        #1d4ed8
                    );

                color: white;
                border-radius: 22px;
                padding: 26px;
                margin-bottom: 20px;
                box-shadow: 0 12px 30px rgba(37, 99, 235, 0.18);
            }

            .member-header h1 {
                margin: 0 0 8px;
                font-size: 30px;
            }

            .member-header p {
                margin: 0;
                opacity: 0.92;
                line-height: 1.6;
            }

            .member-toolbar {
                display: grid;
                grid-template-columns: 1fr auto;
                gap: 14px;
                align-items: end;
            }

            .member-search-row {
                display: grid;
                grid-template-columns: 1fr auto auto;
                gap: 10px;
                align-items: center;
            }

            .member-search-row .form-input {
                margin: 0;
            }

            .member-count {
                font-size: 17px;
                color: #475569;
                margin-top: 10px;
            }

            .member-name {
                font-weight: 700;
                color: #1e293b;
            }

            .member-id {
                font-weight: 700;
                color: #1d4ed8;
                white-space: nowrap;
            }

            .member-phone {
                white-space: nowrap;
            }

            .status-badge {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                min-width: 72px;
                padding: 6px 10px;
                border-radius: 999px;
                font-size: 15px;
                font-weight: 700;
                white-space: nowrap;
            }

            .status-active {
                background: #dcfce7;
                color: #166534;
            }

            .status-paused {
                background: #fef3c7;
                color: #92400e;
            }

            .status-stopped {
                background: #fee2e2;
                color: #991b1b;
            }

            .status-default {
                background: #e2e8f0;
                color: #334155;
            }

            .table-actions {
                display: flex;
                gap: 8px;
                flex-wrap: wrap;
            }

            .table-actions .btn-tool {
                min-height: auto;
                padding: 8px 13px;
                font-size: 15px;
            }

            .member-empty {
                text-align: center;
                padding: 45px 20px;
                color: #64748b;
            }

            @media (max-width: 760px) {

                .member-page {
                    padding-left: 12px;
                    padding-right: 12px;
                }

                .member-header {
                    padding: 22px 18px;
                    border-radius: 18px;
                }

                .member-header h1 {
                    font-size: 26px;
                }

                .member-toolbar {
                    grid-template-columns: 1fr;
                }

                .member-search-row {
                    grid-template-columns: 1fr;
                }

                .member-toolbar .btn-tool,
                .member-search-row .btn-tool {
                    width: 100%;
                }

                .record-table {
                    min-width: 880px;
                }
            }
        </style>
    </head>

    <body>

    <div class="page member-page">

        <div class="member-header">
            <h1>👥 月费会员管理</h1>

            <p>
                查询、新增及编辑 CHE 与 STW 月费会员资料。
            </p>
        </div>

        <div class="card">

            <div class="member-toolbar">

                <div>
                    <div class="section-title">
                        🔍 搜索会员
                    </div>

                    <form method="get">

                        <div class="member-search-row">

                            <input
                                class="form-input"
                                name="q"
                                value="{{ q }}"
                                placeholder="输入编号、姓名、英文名或电话"
                                autocomplete="off"
                            >

                            <button
                                class="btn-tool btn-primary"
                                type="submit"
                            >
                                🔍 搜索
                            </button>

                            {% if q %}
                            <a
                                class="btn-tool btn-secondary"
                                href="{{ url_for(
                                    'finance.member_management'
                                ) }}"
                            >
                                ✕ 清除
                            </a>
                            {% endif %}

                        </div>

                    </form>
                </div>

                <a
                    class="btn-tool btn-success"
                    href="{{ url_for('finance.add_member') }}"
                >
                    ➕ 新增会员
                </a>

            </div>

            <div class="member-count">
                {% if q %}
                    搜索结果：
                    <strong>{{ total_members }}</strong>
                    位会员
                {% else %}
                    当前显示：
                    <strong>{{ total_members }}</strong>
                    位会员
                {% endif %}
            </div>

        </div>

        <div class="card">

            <div class="section-title">
                📋 会员名单
            </div>

            {% if rows %}

            <div class="table-responsive">

                <table class="record-table">

                    <thead>
                        <tr>
                            <th>会员编号</th>
                            <th>姓名</th>
                            <th>英文名</th>
                            <th>电话</th>
                            <th>会员状态</th>
                            <th>备注</th>
                            <th>操作</th>
                        </tr>
                    </thead>

                    <tbody>

                        {% for r in rows %}

                        {% set status_text = r.status or '在册' %}

                        <tr>

                            <td>
                                <span class="member-id">
                                    {{ r.member_id }}
                                </span>
                            </td>

                            <td>
                                <span class="member-name">
                                    {{ r.name }}
                                </span>
                            </td>

                            <td>
                                {{ r.english_name or "-" }}
                            </td>

                            <td class="member-phone">
                                {% if r.phone %}
                                    <a href="tel:{{ r.phone }}">
                                        {{ r.phone }}
                                    </a>
                                {% else %}
                                    -
                                {% endif %}
                            </td>

                            <td>

                                {% if status_text in [
                                    '在册',
                                    'active',
                                    '正常',
                                    '供养中'
                                ] %}

                                    <span
                                        class="
                                            status-badge
                                            status-active
                                        "
                                    >
                                        在册
                                    </span>

                                {% elif status_text in [
                                    '暂停',
                                    'paused',
                                    '停供'
                                ] %}

                                    <span
                                        class="
                                            status-badge
                                            status-paused
                                        "
                                    >
                                        暂停
                                    </span>

                                {% elif status_text in [
                                    '停止',
                                    'stopped',
                                    'inactive',
                                    '永久停止'
                                ] %}

                                    <span
                                        class="
                                            status-badge
                                            status-stopped
                                        "
                                    >
                                        停止
                                    </span>

                                {% else %}

                                    <span
                                        class="
                                            status-badge
                                            status-default
                                        "
                                    >
                                        {{ status_text }}
                                    </span>

                                {% endif %}

                            </td>

                            <td>
                                {{ r.remark or "-" }}
                            </td>

                            <td>

                                <div class="table-actions">

                                    <a
                                        class="
                                            btn-tool
                                            btn-primary
                                        "
                                        href="{{ url_for(
                                            'finance.edit_member',
                                            member_id=r.member_id
                                        ) }}"
                                    >
                                        ✏️ 编辑
                                    </a>

                                </div>

                            </td>

                        </tr>

                        {% endfor %}

                    </tbody>

                </table>

            </div>

            {% else %}

            <div class="member-empty">
                <div style="font-size:44px;">
                    🔎
                </div>

                <h3>找不到相关会员</h3>

                <p>
                    请检查会员编号、姓名、英文名或电话号码。
                </p>
            </div>

            {% endif %}

        </div>

        <div class="btn-row">

            <a
                class="btn-tool btn-secondary"
                href="{{ url_for('finance.finance_home') }}"
            >
                ← 返回财政首页
            </a>

        </div>

    </div>

    </body>
    </html>
    """,
        rows=rows,
        q=q,
        total_members=total_members
    )

@finance_bp.route("/member/<member_id>/status", methods=["POST"])
def update_member_status(member_id):

    status = request.form.get("status", "").strip()

    allowed_status = ["在册", "停供", "往生"]

    if status not in allowed_status:
        return "Invalid status", 400

    db_query("""
        update members
        set status = %s
        where member_id = %s
    """, (status, member_id))

    return redirect(url_for("finance.member_management"))

@finance_bp.route("/add_member", methods=["GET","POST"])
def add_member():
    return "Add Member Coming Soon"

@finance_bp.route("/monthly_fee_batch/<branch>", methods=["GET", "POST"])
def monthly_fee_batch(branch):

    branch = branch.upper()

    if branch not in ["CHE", "STW"]:
        return "Invalid branch", 400

    message = ""
    preview_rows = []

    raw_text = request.form.get("raw_text", "").strip()

    next_receipt_no = get_next_receipt_no(
        branch,
        "月费",
    )
    next_receipt_raw = next_receipt_no.replace(branch, "", 1)

    receipt_start_raw = request.form.get(
        "receipt_start",
        ""
    ).strip().upper()

    if not receipt_start_raw:
        receipt_start_raw = next_receipt_raw

    if receipt_start_raw.isdigit():
        receipt_start = (
            branch
            + str(int(receipt_start_raw)).zfill(7)
        )
    else:
        receipt_start = receipt_start_raw

    default_receipt_date = (
        request.form.get("receipt_date")
        or date.today().isoformat()
    )

    payment_method = request.form.get(
        "payment_method",
        "现金"
    )

    bank_payment_date = (
        request.form.get("payment_date")
        or default_receipt_date
    )

    action = request.form.get("action", "")

    default_amount = money(
        request.form.get("default_amount")
        or 350
    )

    def make_receipt_no(start_no, index):
        match = re.match(r"^([A-Z]+)(\d+)$", start_no)

        if not match:
            return ""

        prefix = match.group(1)
        number_text = match.group(2)

        return (
            prefix
            + str(int(number_text) + index).zfill(
                len(number_text)
            )
        )

    def parse_date_header(line):
        """
        支持：
        @2026-07-10
        @ 2026-07-10
        """
        match = re.match(
            r"^@\s*(\d{4}-\d{2}-\d{2})$",
            line
        )

        if not match:
            return None

        date_text = match.group(1)

        try:
            return date.fromisoformat(date_text).isoformat()
        except ValueError:
            return "INVALID"

    def build_preview():
        rows = []
        current_receipt_date = default_receipt_date
        receipt_index = 0
        today_date = date.today()

        for line_no, original_line in enumerate(
            raw_text.splitlines(),
            start=1
        ):
            line = original_line.strip()

            if not line:
                continue

            if line.startswith("@"):
                parsed_date = parse_date_header(line)

                if parsed_date == "INVALID":
                    rows.append({
                        "error": (
                            f"第 {line_no} 行日期无效：{line}"
                        ),
                        "warning": "",
                        "raw": line,
                        "receipt_date": "",
                    })

                elif not parsed_date:
                    rows.append({
                        "error": (
                            f"第 {line_no} 行日期格式错误，"
                            "请使用 @YYYY-MM-DD"
                        ),
                        "warning": "",
                        "raw": line,
                        "receipt_date": "",
                    })

                else:
                    current_receipt_date = parsed_date

                continue

            parts = line.split()

            # 支持：
            # 108
            # 108 100
            # 张三
            # 张三 100
            if (
                len(parts) >= 2
                and re.fullmatch(
                    r"\d+(?:\.\d+)?",
                    parts[-1]
                )
            ):
                amount = money(parts[-1])
                raw_member_keyword = (
                    " ".join(parts[:-1]).strip()
                )

            else:
                amount = default_amount
                raw_member_keyword = line

            receipt_no = make_receipt_no(
                receipt_start,
                receipt_index
            )
            receipt_index += 1

            error_messages = []
            warning_messages = []

            # A. 金额基本检查
            if amount <= 0:
                error_messages.append(
                    f"第 {line_no} 行金额无效：{line}"
                )

            # 月费必须是 RM50 的倍数
            elif amount % 50 != 0:
                error_messages.append(
                    f"月费 RM{amount:.2f} "
                    "不是 RM50 的倍数"
                )

            # B. 日期检查
            try:
                receipt_date_obj = date.fromisoformat(
                    current_receipt_date
                )

                # V6：检查该收条月份是否已经月结
                month_lock_error = require_finance_month_open(
                    current_receipt_date,
                    get_fund_account("月费")
                )

                if month_lock_error:
                    error_messages.append(
                        month_lock_error
                    )

                if receipt_date_obj > today_date:
                    error_messages.append(
                        "收条日期是未来日期"
                    )

                else:
                    days_old = (
                        today_date - receipt_date_obj
                    ).days

                    if days_old > 30:
                        warning_messages.append(
                            f"收条日期距今已有 "
                            f"{days_old} 天，"
                            "请确认是否为旧账补录"
                        )

            except (TypeError, ValueError):
                error_messages.append(
                    "收条日期格式无效"
                )

            if error_messages:
                rows.append({
                    "error": "；".join(error_messages),
                    "warning": "；".join(warning_messages),
                    "raw": line,
                    "receipt_no": receipt_no,
                    "receipt_date": current_receipt_date,
                })
                continue

            # C. 会员查找
            if raw_member_keyword.isdigit():
                member_id = (
                    f"{branch}-{int(raw_member_keyword)}"
                )

            else:
                member_id = normalize_member_id(
                    raw_member_keyword,
                    default_branch=branch
                )

            # 先按会员编号查找
            member = db_query("""
                select *
                from members
                where member_id = %s
                limit 1
            """, (
                member_id,
            ), fetchone=True)

            # 编号找不到时，再按中文名或英文名查找
            if not member:
                name_matches = db_query("""
                    select *
                    from members
                    where branch = %s
                      and (
                            lower(trim(name))
                                = lower(trim(%s))
                         or lower(
                                trim(
                                    coalesce(
                                        english_name,
                                        ''
                                    )
                                )
                            )
                                = lower(trim(%s))
                      )
                    order by member_id
                    limit 2
                """, (
                    branch,
                    raw_member_keyword,
                    raw_member_keyword,
                ), fetchall=True) or []

                if len(name_matches) == 1:
                    member = name_matches[0]

                elif len(name_matches) > 1:
                    rows.append({
                        "error": (
                            f"第 {line_no} 行姓名有重复，"
                            "请改用会员编号："
                            f"{raw_member_keyword}"
                        ),
                        "warning": "",
                        "raw": line,
                        "receipt_no": receipt_no,
                        "receipt_date": current_receipt_date,
                    })
                    continue

            if not member:
                rows.append({
                    "error": (
                        f"第 {line_no} 行找不到会员："
                        f"{raw_member_keyword}"
                    ),
                    "warning": "",
                    "raw": line,
                    "receipt_no": receipt_no,
                    "receipt_date": current_receipt_date,
                })
                continue

            # 姓名查找成功后，必须改用数据库里的正式编号
            member_id = member["member_id"]

            # D. 计算供养月份
            paid = db_query("""
                select max(end_month) as paid_until
                from member_payments
                where member_id = %s
                  and coalesce(status, 'active') = 'active'
            """, (
                member_id,
            ), fetchone=True)

            paid_until_date = (
                paid["paid_until"]
                if paid
                else None
            )

            month_from = next_month_ym(
                paid_until_date
            )

            month_count = int(amount / 50)

            month_to = add_months_ym(
                month_from,
                month_count - 1
            )

            # E. 收条重复检查
            existing_finance = db_query("""
                select id
                from finance_records
                where receipt_no = %s
                limit 1
            """, (
                receipt_no,
            ), fetchone=True)

            existing_payment = db_query("""
                select id
                from member_payments
                where receipt_no = %s
                  and coalesce(status, 'active') = 'active'
                limit 1
            """, (
                receipt_no,
            ), fetchone=True)

            if existing_finance and existing_payment:
                error_messages.append(
                    "收条已存在于财政记录和月费查询"
                )

            elif existing_finance:
                error_messages.append(
                    "收条已存在于财政记录"
                )

            elif existing_payment:
                error_messages.append(
                    "收条已存在于月费查询"
                )

            # F. 同一会员、同一天重复月费提醒
            same_day_payment = db_query("""
                select
                    receipt_no,
                    amount
                from member_payments
                where member_id = %s
                  and receipt_date = %s
                  and coalesce(status, 'active') = 'active'
                order by id desc
                limit 1
            """, (
                member_id,
                current_receipt_date,
            ), fetchone=True)

            if same_day_payment:
                old_receipt = (
                    same_day_payment.get("receipt_no")
                    or "无收条编号"
                )

                old_amount = float(
                    same_day_payment.get("amount")
                    or 0
                )

                warning_messages.append(
                    "此会员同一天已有月费记录："
                    f"{old_receipt}，"
                    f"RM{old_amount:.2f}"
                )

            # G. 银行付款日期检查
            if payment_method == "银行过账":
                row_payment_date = bank_payment_date

                try:
                    payment_date_obj = date.fromisoformat(
                        row_payment_date
                    )

                    if payment_date_obj > today_date:
                        error_messages.append(
                            "银行付款日期是未来日期"
                        )

                except (TypeError, ValueError):
                    error_messages.append(
                        "银行付款日期格式无效"
                    )

            else:
                row_payment_date = current_receipt_date

            rows.append({
                "error": "；".join(error_messages),
                "warning": "；".join(warning_messages),
                "receipt_no": receipt_no,
                "receipt_date": current_receipt_date,
                "payment_date": row_payment_date,
                "member_id": member_id,
                "name": (
                    member.get("姓名")
                    or member.get("name")
                ),
                "phone": (
                    member.get("电话号码")
                    or member.get("phone")
                ),
                "amount": amount,
                "month_from": month_from,
                "month_to": month_to,
                "month_count": month_count,
            })

        return rows

    if request.method == "POST":

        if not receipt_start:
            message = "请填写收条开始号码"

        elif not re.match(
            rf"^{branch}\d+$",
            receipt_start
        ):
            message = (
                "收条号码格式错误，例如："
                f"{branch}0001501，或只输入 1501"
            )

        elif not raw_text:
            message = "请先加入或贴上月费资料"

        else:
            preview_rows = build_preview()

            data_rows = [
                row
                for row in preview_rows
                if row.get("member_id")
            ]

            has_error = any(
                row.get("error")
                for row in preview_rows
            )

            if not data_rows and not has_error:
                message = "没有可预览的会员资料"

            elif action == "confirm" and not has_error:

                confirm_lock_error = None

                # 保存前重新检查所有月份
                # 防止预览后，该月份才被别人完成月结
                for row in data_rows:

                    confirm_lock_error = (
                        require_finance_month_open(
                            row["receipt_date"],
                            get_fund_account("月费")
                        )
                    )

                    if confirm_lock_error:
                        break

                if confirm_lock_error:

                    message = confirm_lock_error

                else:

                    last_receipt_no = ""

                    for row in data_rows:
                        month_from_db = (
                            row["month_from"] + "-01"
                        )

                        month_to_db = (
                            row["month_to"] + "-01"
                        )

                        db_query("""
                            insert into finance_records
                            (
                                record_type,
                                fund_account,
                                record_date,
                                receipt_date,
                                category,
                                receipt_no,
                                member_id,
                                name,
                                phone,
                                amount,
                                payment_method,
                                month_from,
                                month_to,
                                remarks
                            )
                            values
                            (
                                %s, %s, %s, %s,
                                '月费',
                                %s, %s, %s, %s,
                                %s, %s, %s, %s, %s
                            )
                        """, (
                            "income",
                            get_fund_account("月费"),
                            row["payment_date"],
                            row["receipt_date"],
                            row["receipt_no"],
                            row["member_id"],
                            row["name"],
                            row["phone"],
                            row["amount"],
                            payment_method,
                            month_from_db,
                            month_to_db,
                            "批量月费录入",
                        ))

                        db_query("""
                            insert into member_payments
                            (
                                payment_date,
                                receipt_date,
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
                                %s, %s, %s, %s, %s,
                                %s, %s, %s, %s
                            )
                        """, (
                            row["payment_date"],
                            row["receipt_date"],
                            row["member_id"],
                            row["name"],
                            row["receipt_no"],
                            row["amount"],
                            month_from_db,
                            month_to_db,
                            row["month_count"],
                        ))

                        last_receipt_no = row["receipt_no"]

                    if last_receipt_no:
                        update_receipt_book_number(
                            branch,
                            "月费",
                            last_receipt_no,
                        )

                    return redirect(
                        url_for("finance.records")
                    )

# ============================================================
# 页面预览区还要做以下两处修改
# ============================================================

# 1. 在 <style> 中加入：
#
# .preview-warning{
#     color:#b77900;
#     font-weight:700;
# }
#
# 2. 把“检查结果”单元格替换成：
#
# {% if r.error %}
#     <span class="preview-error">
#         ❌ {{ r.error }}
#     </span>
#
# {% elif r.warning %}
#     <span class="preview-warning">
#         ⚠️ {{ r.warning }}
#     </span>
#
# {% else %}
#     <span class="preview-success">
#         ✅ 可以入账
#     </span>
# {% endif %}
#
# 注意：
# - error 会阻止“确认全部入账”
# - warning 只提醒，仍然允许确认入账

    return render_template_string(FINANCE_DATE_COMPONENT + """
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">

        <title>{{ branch }} 月费录入</title>

        <link
            rel="stylesheet"
            href="{{ url_for('static', filename='css/toolbox.css') }}"
        >

        <style>
            .finance-form-page{
                max-width:920px;
            }

            .finance-header-card{
                margin-bottom:18px;
                background:linear-gradient(135deg,#1769aa,#2589c5);
                color:#fff;
            }

            .finance-header-card .page-title,
            .finance-header-card .page-subtitle{
                color:#fff;
            }

            .finance-back-row{
                margin-bottom:16px;
            }

            .step-title{
                display:flex;
                align-items:center;
                gap:10px;
                margin-bottom:18px;
            }

            .step-title .section-title{
                margin:0;
            }

            .step-badge{
                flex:0 0 auto;
                display:inline-flex;
                align-items:center;
                justify-content:center;
                width:36px;
                height:36px;
                border-radius:50%;
                background:#1769aa;
                color:#fff;
                font-size:17px;
                font-weight:800;
            }

            .settings-grid{
                display:grid;
                grid-template-columns:repeat(2,minmax(0,1fr));
                gap:16px;
            }

            .settings-grid .form-group{
                margin:0;
            }

            .finance-full{
                grid-column:1 / -1;
            }

            .receipt-input-row{
                display:flex;
                align-items:stretch;
                gap:10px;
            }

            .receipt-prefix{
                min-width:78px;
                display:flex;
                align-items:center;
                justify-content:center;
                padding:0 15px;
                border-radius:10px;
                background:#eaf4ff;
                color:#1769aa;
                font-weight:800;
                font-size:20px;
            }

            .receipt-input-row .form-input{
                flex:1;
                min-width:0;
            }

            .field-help{
                display:block;
                color:#667085;
                font-size:15px;
                line-height:1.55;
                margin-top:7px;
            }

            .quick-member-row{
                display:grid;
                grid-template-columns:minmax(0,1fr) 280px;
                gap:14px;
                align-items:end;
            }

            .amount-panel{
                display:grid;
                grid-template-columns:54px 1fr 54px;
                gap:8px;
                align-items:stretch;
            }

            .amount-step-btn{
                border:1px solid #cbd5e1;
                border-radius:10px;
                background:#f8fafc;
                color:#172033;
                font-size:25px;
                font-weight:800;
                cursor:pointer;
            }

            .amount-step-btn:hover{
                background:#eaf4ff;
                border-color:#8fbce0;
            }

            .amount-step-btn:active{
                transform:scale(.97);
            }

            #quick_amount{
                text-align:center;
                font-size:25px;
                font-weight:800;
            }

            .amount-shortcuts{
                display:flex;
                flex-wrap:wrap;
                gap:8px;
                margin-top:12px;
            }

            .amount-chip{
                min-width:72px;
                border:1px solid #cbd5e1;
                border-radius:999px;
                background:#fff;
                padding:9px 14px;
                color:#1769aa;
                font-size:16px;
                font-weight:800;
                cursor:pointer;
            }

            .amount-chip:hover{
                background:#eaf4ff;
                border-color:#8fbce0;
            }

            .quick-add-btn{
                width:100%;
                margin-top:16px;
            }

            .batch-textarea{
                width:100%;
                min-height:250px;
                resize:vertical;
                line-height:1.75;
                font-family:Consolas,"Microsoft YaHei",monospace;
                font-size:18px;
            }

            .batch-help-grid{
                display:grid;
                grid-template-columns:repeat(3,minmax(0,1fr));
                gap:10px;
                margin-top:12px;
            }

            .batch-help-item{
                padding:12px 14px;
                border:1px solid #dbe3ed;
                border-radius:12px;
                background:#f8fafc;
                color:#475467;
                font-size:14px;
                line-height:1.6;
            }

            .batch-help-item strong{
                display:block;
                margin-bottom:4px;
                color:#172033;
            }

            .batch-help-item code{
                color:#1769aa;
                font-weight:800;
            }

            .action-row{
                display:grid;
                grid-template-columns:1fr 1fr;
                gap:12px;
                margin-top:18px;
            }

            .action-row.single{
                grid-template-columns:1fr;
            }

            .action-row .btn-tool{
                width:100%;
            }

            .payment-date-box{
                display:none;
            }

            .preview-success{
                color:#16863a;
                font-weight:700;
            }

            .preview-error{
                color:#c62828;
                font-weight:700;
            }

            .preview-date{
                white-space:nowrap;
                font-weight:700;
                color:#1769aa;
            }
                                  
            .preview-warning{
                color:#b77900;
                font-weight:700;
            }

            @media(max-width:760px){
                .settings-grid,
                .quick-member-row,
                .batch-help-grid,
                .action-row{
                    grid-template-columns:1fr;
                }

                .finance-full{
                    grid-column:auto;
                }

                .receipt-input-row{
                    align-items:stretch;
                }
            }
        </style>
    </head>

    <body>
    <div class="page finance-form-page">

        <div class="finance-back-row">
            <a
                class="btn-tool btn-secondary"
                href="{{ url_for('finance.finance_income_menu') }}"
            >
                ← 返回收入录入
            </a>
        </div>

        <div class="card finance-header-card">
            <h1 class="page-title">💳 {{ branch }} 月费录入</h1>
            <p class="page-subtitle">
                先设定收条资料，再加入会员；单笔与批量共用同一份待录入清单。
            </p>
        </div>

        {% if message %}
            <div class="alert alert-danger">{{ message }}</div>
        {% endif %}

        <form method="post">

            <div class="card">
                <div class="step-title">
                    <span class="step-badge">1</span>
                    <h2 class="section-title">收条资料与单笔加入</h2>
                </div>

                <div class="settings-grid">
                    <div class="form-group">
                        <label class="form-label">收条开始号码</label>

                        <div class="receipt-input-row">
                            <div class="receipt-prefix">{{ branch }}</div>

                            <input
                                class="form-input"
                                name="receipt_start"
                                value="{{ receipt_start_raw }}"
                                inputmode="numeric"
                                placeholder="例如 1501"
                                required
                            >
                        </div>

                        <span class="field-help">
                            系统建议下一张：<strong>{{ next_receipt_no }}</strong>
                        </span>
                    </div>

                    <div class="form-group">
                        <label class="form-label">默认开收条日期</label>

                        <input
                            class="form-input"
                            name="receipt_date"
                            type="date"
                            value="{{ default_receipt_date }}"
                            required
                        >

                        <span class="field-help">
                            没有写 @日期 的会员会使用这个日期。
                        </span>
                    </div>
                </div>

                <div style="border-top:1px solid #e2e8f0;margin:22px 0 18px;"></div>

                <h3 class="entry-method-title" style="margin-bottom:12px;">
                    👤 单笔快速加入
                </h3>

                <div class="quick-member-row">
                    <div class="form-group">
                        <label class="form-label">会员编号或姓名</label>
                        <input
                            class="form-input"
                            id="quick_member_keyword"
                            type="text"
                            placeholder="例如：108、{{ branch }}-108、张三"
                            autocomplete="off"
                        >
                    </div>

                    <div class="form-group">
                        <label class="form-label">本笔金额 RM</label>
                        <div class="amount-panel">
                            <button
                                type="button"
                                class="amount-step-btn"
                                onclick="changeQuickAmount(-50)"
                            >−</button>

                            <input
                                class="form-input"
                                id="quick_amount"
                                type="number"
                                min="50"
                                step="50"
                                value="{{ default_amount }}"
                            >

                            <button
                                type="button"
                                class="amount-step-btn"
                                onclick="changeQuickAmount(50)"
                            >＋</button>
                        </div>
                    </div>
                </div>

                <div class="amount-shortcuts">
                    {% for amount in [50, 100, 150, 200, 250, 300] %}
                        <button
                            type="button"
                            class="amount-chip"
                            onclick="setQuickAmount({{ amount }})"
                        >
                            RM{{ amount }}
                        </button>
                    {% endfor %}
                </div>

                <button
                    class="btn-tool btn-success quick-add-btn"
                    type="button"
                    onclick="addQuickMember()"
                >
                    ➕ 加入待录入清单
                </button>
            </div>

            <div class="card">
                <div class="step-title">
                    <span class="step-badge">2</span>
                    <h2 class="section-title">待录入清单与付款资料</h2>
                </div>

                <p class="page-subtitle" style="margin-top:-6px;">
                    单笔加入会自动出现在这里，也可以直接贴上多位会员资料。
                </p>

                <textarea
                    class="form-input batch-textarea"
                    id="raw_text"
                    name="raw_text"
                    placeholder="例如：
    @2026-07-10
    108
    张三
    188 100

    @2026-07-11
    69
    205 300"
                >{{ raw_text }}</textarea>

                <div class="batch-help-grid">
                    <div class="batch-help-item">
                        <strong>普通月费</strong>
                        <code>108</code> 或 <code>张三</code>
                    </div>

                    <div class="batch-help-item">
                        <strong>指定金额</strong>
                        <code>188 100</code>
                    </div>

                    <div class="batch-help-item">
                        <strong>切换收条日期</strong>
                        <code>@2026-07-11</code>
                    </div>
                </div>

                <div style="border-top:1px solid #e2e8f0;margin:22px 0 18px;"></div>

                <div class="settings-grid">
                    <div class="form-group">
                        <label class="form-label">付款方式</label>

                        <select
                            class="form-input"
                            name="payment_method"
                            id="payment_method"
                            onchange="togglePaymentDate()"
                        >
                            {% for method in ['现金', '银行过账', '支票'] %}
                                <option
                                    value="{{ method }}"
                                    {% if payment_method == method %}selected{% endif %}
                                >
                                    {{ method }}
                                </option>
                            {% endfor %}
                        </select>
                    </div>

                    <div
                        class="form-group payment-date-box"
                        id="payment_date_box"
                    >
                        <label class="form-label">银行付款日期</label>

                        <input
                            class="form-input"
                            name="payment_date"
                            type="date"
                            value="{{ bank_payment_date }}"
                        >
                    </div>

                    <div class="form-group finance-full">
                        <label class="form-label">批量清单默认月费 RM</label>

                        <input
                            class="form-input"
                            name="default_amount"
                            type="number"
                            step="50"
                            min="50"
                            value="{{ default_amount }}"
                            required
                        >

                        <span class="field-help">
                            清单只写编号或姓名时使用；行内有金额时，以行内金额为准。
                        </span>
                    </div>
                </div>

                <div class="action-row {% if not (preview_rows and not has_preview_error) %}single{% endif %}">
                    <button
                        class="btn-tool btn-primary"
                        type="submit"
                        name="action"
                        value="preview"
                    >
                        👁️ 预览资料
                    </button>

                    {% if preview_rows and not has_preview_error %}
                        <button
                            class="btn-tool btn-success"
                            type="submit"
                            name="action"
                            value="confirm"
                            onclick="return confirm('确定全部入账？');"
                        >
                            ✅ 确认全部入账
                        </button>
                    {% endif %}
                </div>
            </div>

            {% if preview_rows %}
                <div class="card">
                    <div class="step-title">
                        <span class="step-badge">3</span>
                        <h2 class="section-title">月费录入预览</h2>
                    </div>

                    <div class="table-responsive">
                        <table class="record-table">
                            <thead>
                                <tr>
                                    <th>收条</th>
                                    <th>收条日期</th>
                                    <th>会员编号</th>
                                    <th>姓名</th>
                                    <th>金额</th>
                                    <th>开始月份</th>
                                    <th>缴费至</th>
                                    <th>月数</th>
                                    <th>检查结果</th>
                                </tr>
                            </thead>

                            <tbody>
                                {% for r in preview_rows %}
                                    <tr>
                                        <td>{{ r.receipt_no or '-' }}</td>
                                        <td class="preview-date">{{ r.receipt_date or '-' }}</td>
                                        <td><strong>{{ r.member_id or '-' }}</strong></td>
                                        <td>{{ r.name or '-' }}</td>
                                        <td>
                                            {% if r.amount %}
                                                RM {{ '%.2f'|format(r.amount) }}
                                            {% else %}
                                                -
                                            {% endif %}
                                        </td>
                                        <td>{{ r.month_from or '-' }}</td>
                                        <td>{{ r.month_to or '-' }}</td>
                                        <td>
                                            {% if r.month_count %}
                                                {{ r.month_count }} 个月
                                            {% else %}
                                                -
                                            {% endif %}
                                        </td>
                                        <td>
                                            {% if r.error %}
                                                <span class="preview-error">
                                                    ❌ {{ r.error }}
                                                </span>

                                            {% elif r.warning %}
                                                <span class="preview-warning">
                                                    ⚠️ {{ r.warning }}
                                                </span>

                                            {% else %}
                                                <span class="preview-success">
                                                    ✅ 可以入账
                                                </span>
                                            {% endif %}
                                        </td>
                                    </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
            {% endif %}

        </form>
    </div>

    <script>
    function changeQuickAmount(change){
        const input = document.getElementById("quick_amount");

        if(!input){
            return;
        }

        let current = Number(input.value);

        if(!Number.isFinite(current)){
            current = 50;
        }

        current += change;

        if(current < 50){
            current = 50;
        }

        input.value = current;
    }

    function setQuickAmount(amount){
        const input = document.getElementById("quick_amount");

        if(input){
            input.value = amount;
        }
    }

    function addQuickMember(){
        const keywordInput = document.getElementById("quick_member_keyword");
        const amountInput = document.getElementById("quick_amount");
        const textarea = document.getElementById("raw_text");

        if(!keywordInput || !amountInput || !textarea){
            return;
        }

        const keyword = keywordInput.value.trim();
        const amount = amountInput.value.trim();

        if(!keyword){
            alert("请输入会员编号或姓名");
            keywordInput.focus();
            return;
        }

        const newLine = amount
            ? keyword + " " + amount
            : keyword;

        const current = textarea.value.trim();

        textarea.value = current
            ? current + "\\n" + newLine
            : newLine;

        keywordInput.value = "";
        keywordInput.focus();
        textarea.scrollTop = textarea.scrollHeight;
    }

    function togglePaymentDate(){
        const method = document.getElementById("payment_method");
        const box = document.getElementById("payment_date_box");

        if(!method || !box){
            return;
        }

        box.style.display = method.value === "银行过账"
            ? "block"
            : "none";
    }

    document.addEventListener("DOMContentLoaded", function(){
        const keywordInput = document.getElementById("quick_member_keyword");

        if(keywordInput){
            keywordInput.addEventListener("keydown", function(event){
                if(event.key === "Enter"){
                    event.preventDefault();
                    addQuickMember();
                }
            });
        }

        togglePaymentDate();
    });
    </script>

    </body>
    </html>
    """,
        branch=branch,
        message=message,
        raw_text=raw_text,
        receipt_start_raw=receipt_start_raw,
        next_receipt_no=next_receipt_no,
        default_receipt_date=default_receipt_date,
        bank_payment_date=bank_payment_date,
        payment_method=payment_method,
        default_amount=default_amount,
        preview_rows=preview_rows,
        has_preview_error=any(
            row.get("error")
            for row in preview_rows
        ),
    )

def get_recent_donors(category=None):
    if category:
        return db_query("""
            select
                name,
                max(phone) as phone,
                count(*) as times,
                max(record_date) as last_date
            from finance_records
            where record_type = 'income'
              and category = %s
              and coalesce(status, 'confirmed') <> 'cancelled'
              and coalesce(name, '') <> ''
              and category <> '月费'
            group by name
            order by max(record_date) desc, count(*) desc
            limit 100
        """, (category,), fetchall=True)

    return db_query("""
        select
            name,
            max(phone) as phone,
            count(*) as times,
            max(record_date) as last_date
        from finance_records
        where record_type = 'income'
          and coalesce(status, 'confirmed') <> 'cancelled'
          and coalesce(name, '') <> ''
          and category <> '月费'
        group by name
        order by max(record_date) desc, count(*) desc
        limit 100
    """, fetchall=True)
    

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
        receipt_date = request.form.get("receipt_date") or date.today()
        record_date = request.form.get("record_date") or date.today()

        name = request.form.get("name", "").strip()
        phone = request.form.get("phone", "").strip()
        amount = money(request.form.get("amount"))
        payment_method = request.form.get("payment_method", "现金")
        remarks = request.form.get("remarks", "").strip()

        if not receipt_no.startswith("CHE"):
            message = "收条号码必须以 CHE 开头"

        elif amount <= 0:
            message = "金额必须大过 0"

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
                        receipt_date,
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
                        %s, %s, %s, %s, %s, %s
                    )
                """, (
                    "income",
                    get_fund_account(category),
                    record_date,
                    receipt_date,
                    category,
                    receipt_no,
                    name,
                    phone,
                    amount,
                    payment_method,
                    remarks
                ))

                return redirect(url_for("finance.records"))

    return render_template_string(FINANCE_STYLE + FINANCE_DATE_COMPONENT + """
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
            开收条日期：
            <input name="receipt_date" type="date" value="{{ today }}" required>
        </p>

        <p>
            付款日期：
            <input name="record_date" type="date" value="{{ today }}" required>
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
                step="1.00"
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
                                  
    <script>

    function addDonor(name){

        let textarea = document.querySelector(
            'textarea[name="raw_text"]'
        );

        if(textarea.value.trim() === ""){
            textarea.value = name;
        }else{
            textarea.value += "\\n" + name;
        }

        textarea.focus();
    }

    </script>

    """,
    category=category,
    next_receipt_no=next_receipt_no,
    message=message,
    today=date.today().isoformat()
    )

def search_finance_donors(keyword: str, branch: str = "CHE", limit_per_source: int = 20):
    """统一搜索月费会员、非会员义工及历史布施人；每个来源保留独立名额。"""

    keyword = str(keyword or "").strip()
    branch = str(branch or "CHE").strip().upper()
    limit_per_source = max(1, min(int(limit_per_source or 20), 50))

    if not keyword:
        return []

    like_keyword = f"%{keyword}%"
    compact_keyword = "".join(keyword.split())
    like_compact = f"%{compact_keyword}%"

    results = []
    seen = set()

    def add_result(item):
        name = str(item.get("name") or "").strip()
        phone = str(item.get("phone") or "").strip()
        member_id = str(item.get("member_id") or "").strip()
        volunteer_id = str(item.get("volunteer_id") or "").strip()

        if not name:
            return

        identity = (
            member_id.upper()
            or volunteer_id.upper()
            or (name.casefold(), "".join(ch for ch in phone if ch.isdigit()))
        )

        if identity in seen:
            return

        seen.add(identity)
        results.append(item)

    member_rows = db_query("""
        select
            member_id,
            name,
            coalesce(english_name, '') as english_name,
            coalesce(phone, '') as phone
        from members
        where (%s = '' or upper(coalesce(branch, 'CHE')) = %s)
          and (
                member_id ilike %s
             or name ilike %s
             or coalesce(english_name, '') ilike %s
             or regexp_replace(coalesce(phone, ''), '[^0-9]', '', 'g') ilike %s
          )
        order by
            case
                when upper(member_id) = upper(%s) then 0
                when name = %s then 1
                when coalesce(english_name, '') = %s then 2
                else 3
            end,
            name
        limit %s
    """, (
        branch,
        branch,
        like_keyword,
        like_keyword,
        like_keyword,
        like_compact,
        keyword,
        keyword,
        keyword,
        limit_per_source,
    ), fetchall=True) or []

    for row in member_rows:
        add_result({
            "source": "member",
            "source_label": "月费会员",
            "name": row.get("name") or "",
            "english_name": row.get("english_name") or "",
            "member_id": row.get("member_id") or "",
            "volunteer_id": "",
            "phone": row.get("phone") or "",
            "last_date": "",
            "times": 0,
        })

    volunteer_rows = db_query("""
        select
            id as volunteer_id,
            name,
            coalesce(phone, '') as phone
        from volunteers
        where (%s = '' or upper(coalesce(branch, 'CHE')) = %s)
          and not exists (
                select 1
                from members m
                where upper(coalesce(m.branch, 'CHE')) = upper(coalesce(volunteers.branch, 'CHE'))
                  and (
                        trim(m.name) = trim(volunteers.name)
                     or (
                            regexp_replace(coalesce(m.phone, ''), '[^0-9]', '', 'g') <> ''
                        and regexp_replace(coalesce(m.phone, ''), '[^0-9]', '', 'g')
                            = regexp_replace(coalesce(volunteers.phone, ''), '[^0-9]', '', 'g')
                     )
                  )
          )
          and (
                cast(id as text) ilike %s
             or name ilike %s
             or regexp_replace(coalesce(phone, ''), '[^0-9]', '', 'g') ilike %s
          )
        order by
            case
                when upper(cast(id as text)) = upper(%s) then 0
                when name = %s then 1
                else 2
            end,
            name
        limit %s
    """, (
        branch,
        branch,
        like_keyword,
        like_keyword,
        like_compact,
        keyword,
        keyword,
        limit_per_source,
    ), fetchall=True) or []

    for row in volunteer_rows:
        volunteer_id = str(row.get("volunteer_id") or "").strip()
        name = str(row.get("name") or "").strip()
        phone = str(row.get("phone") or "").strip()

        duplicate = any(
            (name and item.get("name") == name)
            and (
                not phone
                or not item.get("phone")
                or "".join(ch for ch in item.get("phone", "") if ch.isdigit())
                   == "".join(ch for ch in phone if ch.isdigit())
            )
            for item in results
        )
        if duplicate:
            continue

        add_result({
            "source": "volunteer",
            "source_label": "义工",
            "name": name,
            "english_name": "",
            "member_id": "",
            "volunteer_id": volunteer_id,
            "phone": phone,
            "last_date": "",
            "times": 0,
        })

    history_rows = db_query("""
        select
            name,
            (array_agg(coalesce(phone, '') order by record_date desc, id desc))[1] as phone,
            max(record_date) as last_date,
            count(*) as times
        from finance_records
        where category <> '月费'
          and coalesce(status, 'confirmed') <> 'cancelled'
          and coalesce(name, '') <> ''
          and (
                name ilike %s
             or regexp_replace(coalesce(phone, ''), '[^0-9]', '', 'g') ilike %s
          )
        group by name
        order by
            case when name = %s then 0 else 1 end,
            max(record_date) desc,
            count(*) desc
        limit %s
    """, (
        like_keyword,
        like_compact,
        keyword,
        limit_per_source,
    ), fetchall=True) or []

    for row in history_rows:
        name = str(row.get("name") or "").strip()
        phone = str(row.get("phone") or "").strip()

        duplicate = any(
            item.get("name") == name
            and (
                not phone
                or not item.get("phone")
                or "".join(ch for ch in item.get("phone", "") if ch.isdigit())
                   == "".join(ch for ch in phone if ch.isdigit())
            )
            for item in results
        )
        if duplicate:
            continue

        last_date = row.get("last_date")
        add_result({
            "source": "history",
            "source_label": "曾布施佛友",
            "name": name,
            "english_name": "",
            "member_id": "",
            "volunteer_id": "",
            "phone": phone,
            "last_date": last_date.isoformat() if last_date else "",
            "times": int(row.get("times") or 0),
        })

    # 不再让月费会员先占满总名额。
    # 三个来源各自最多 limit_per_source 笔，确保义工和历史布施人都会出现。
    return results


@finance_bp.route("/api/donor-search")
def finance_donor_search():
    keyword = request.args.get("q", "").strip()
    branch = request.args.get("branch", "CHE").strip().upper()

    if len(keyword) < 1:
        return jsonify({"ok": True, "results": []})

    try:
        results = search_finance_donors(keyword, branch, 20)
        return jsonify({"ok": True, "results": results})
    except Exception as exc:
        print("donor search error:", exc)
        return jsonify({
            "ok": False,
            "message": "搜索布施人失败",
            "results": [],
        }), 500


@finance_bp.route("/income_batch/<category>", methods=["GET", "POST"])
def income_batch(category):

    allowed_categories = [
        "财布施",
        "观音村",
        "膳食结缘",
        "观音堂纯檀香布施",
        SPECIAL_DONATION_TITLE,
        "临时特别布施"
    ]

    if category not in allowed_categories:
        return "Invalid category", 400

    message = ""
    preview_rows = []

    raw_text = request.form.get("raw_text", "").strip()

    receipt_category = normalize_receipt_category(category)

    next_receipt_no = get_next_receipt_no(
        "CHE",
        receipt_category,
    )

    next_receipt_raw = next_receipt_no.replace("CHE", "", 1)

    receipt_start_raw = (
        request.form.get("receipt_start", "")
        .strip()
        .upper()
    )

    if not receipt_start_raw:
        receipt_start_raw = next_receipt_raw

    if receipt_start_raw.isdigit():
        receipt_start = (
            "CHE"
            + str(int(receipt_start_raw)).zfill(7)
        )
    else:
        receipt_start = receipt_start_raw

    receipt_date = request.form.get("receipt_date") or date.today().isoformat()
    record_date = receipt_date
    payment_method = request.form.get("payment_method", "现金")
    default_amount = money(request.form.get("default_amount") or 50)
    remarks = request.form.get("remarks", "").strip()
    action = request.form.get("action", "")

    def make_receipt_no(start_no, index):
        m = re.match(r"^([A-Z]+)(\d+)$", start_no)
        if not m:
            return ""

        prefix = m.group(1)
        num = m.group(2)

        return prefix + str(int(num) + index).zfill(len(num))

    def find_donor_info(keyword):
        keyword = str(keyword or "").strip()
        matches = search_finance_donors(keyword, "CHE", 10)

        if matches:
            selected = matches[0]
            return {
                "name": selected.get("name") or keyword,
                "phone": selected.get("phone") or "",
                "source": selected.get("source_label") or "资料库",
            }

        return {
            "name": keyword,
            "phone": "",
            "source": "手动输入",
        }

    def build_preview():
        rows = []
        valid_index = 0
        current_receipt_date = receipt_date

        for line in raw_text.splitlines():
            line = line.strip()

            if not line:
                continue

            # 日期分组格式：
            # @2026-07-10
            # 后续记录都会使用这个收条日期，
            # 直到遇到下一个 @日期。
            if line.startswith("@") or line.startswith("＠"):
                date_text = line[1:].strip()

                try:
                    current_receipt_date = date.fromisoformat(
                        date_text
                    ).isoformat()
                except ValueError:
                    rows.append({
                        "error": (
                            f"日期格式错误：{date_text}，"
                            "请使用 @YYYY-MM-DD"
                        ),
                        "receipt_no": "",
                        "receipt_date": date_text,
                        "name": "",
                        "phone": "",
                        "amount": 0,
                        "source": "日期设置",
                    })

                continue

            parts = line.split()

            if len(parts) >= 2:
                amount = money(parts[-1])
                name = " ".join(parts[:-1]).strip()
            else:
                amount = default_amount
                name = parts[0].strip()

            receipt_no = make_receipt_no(
                receipt_start,
                valid_index
            )
            valid_index += 1

            donor = find_donor_info(name)

            name = donor["name"]
            phone = donor["phone"]
            source = donor["source"]

            existing = db_query("""
                select id
                from finance_records
                where receipt_no = %s
                limit 1
            """, (receipt_no,), fetchone=True)

            error = ""

            month_lock_error = require_finance_month_open(
                current_receipt_date,
                get_fund_account(category)
            )

            if month_lock_error:
                error = month_lock_error
            elif not name:
                error = "姓名不能为空"
            elif amount <= 0:
                error = "金额必须大过 0"
            elif existing:
                error = "收条已存在"

            rows.append({
                "error": error,
                "receipt_no": receipt_no,
                "receipt_date": current_receipt_date,
                "name": name,
                "phone": phone,
                "amount": amount,
                "source": source,
            })

        return rows

    recent_donors = db_query("""
        select
            name,
            max(phone) as phone,
            count(*) as times,
            max(record_date) as last_date
        from finance_records
        where category <> '月费'
          and coalesce(status, 'confirmed') <> 'cancelled'
          and coalesce(name, '') <> ''
        group by name
        order by max(record_date) desc, count(*) desc
        limit 50
    """, fetchall=True)

    if request.method == "POST":

        if category == "临时特别布施" and not remarks:
            message = "临时特别布施请填写活动名称或用途"

        elif not receipt_start:
            message = "请填写收条开始号码"

        elif not re.match(r"^CHE\d+$", receipt_start):
            message = "收条号码格式错误，例如：CHE0001501，或只输入 1501"

        elif not raw_text:
            message = "请粘贴批量资料"

        else:
            preview_rows = build_preview()
            has_error = any(r.get("error") for r in preview_rows)

            if action == "confirm" and not has_error:

                confirm_lock_error = None

                for r in preview_rows:
                    confirm_lock_error = require_finance_month_open(
                        r["receipt_date"],
                        get_fund_account(category)
                    )

                    if confirm_lock_error:
                        break

                if confirm_lock_error:
                    message = confirm_lock_error

                else:
                    last_receipt_no = ""

                    for r in preview_rows:

                        db_query("""
                            insert into finance_records
                            (
                                record_type,
                                fund_account,
                                record_date,
                                receipt_date,
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
                                %s, %s, %s, %s, %s, %s
                            )
                        """, (
                            "income",
                            get_fund_account(category),
                            r["receipt_date"],
                            r["receipt_date"],
                            category,
                            r["receipt_no"],
                            r["name"],
                            r["phone"],
                            r["amount"],
                            payment_method,
                            remarks or "批量布施录入"
                        ))

                        last_receipt_no = r["receipt_no"]

                    if last_receipt_no:
                        update_receipt_book_number(
                            "CHE",
                            receipt_category,
                            last_receipt_no,
                        )

                    return redirect(url_for("finance.records"))

    return render_template_string(FINANCE_DATE_COMPONENT + """
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">

        <meta
            name="viewport"
            content="width=device-width, initial-scale=1"
        >

        <title>{{ category }}录入</title>

        <link
            rel="stylesheet"
            href="{{ url_for('static', filename='css/toolbox.css') }}"
        >

        <style>
            .finance-form-page{
                max-width:900px;
            }

            .finance-header-card{
                margin-bottom:18px;
                background:linear-gradient(135deg,#1877b8,#2695cf);
                color:white;
            }

            .finance-header-card .page-title,
            .finance-header-card .page-subtitle{
                color:white;
            }

            .finance-back-row{
                margin-bottom:16px;
            }

            .finance-form-grid{
                display:grid;
                grid-template-columns:repeat(2,minmax(0,1fr));
                gap:16px;
            }

            .finance-form-grid .form-group{
                margin:0;
            }

            .finance-full{
                grid-column:1 / -1;
            }

            .receipt-input-row{
                display:flex;
                align-items:center;
                gap:10px;
            }

            .receipt-prefix{
                min-height:52px;
                display:flex;
                align-items:center;
                justify-content:center;
                padding:0 16px;
                border-radius:10px;
                background:#eaf4ff;
                color:#1769aa;
                font-weight:800;
                font-size:20px;
            }

            .receipt-input-row .form-input{
                flex:1;
                min-width:0;
            }

            .field-help{
                display:block;
                color:#667085;
                font-size:15px;
                line-height:1.5;
                margin-top:7px;
            }

            .batch-textarea{
                width:100%;
                min-height:260px;
                resize:vertical;
                line-height:1.7;
                font-family:inherit;
            }

            .batch-example{
                background:#f8fafc;
                border:1px dashed #cbd5e1;
                border-radius:12px;
                padding:14px 16px;
                margin-top:12px;
                color:#475467;
                font-size:15px;
                line-height:1.75;
            }

            .batch-example code{
                font-family:Consolas,monospace;
                font-weight:700;
                color:#1769aa;
            }

            .action-row{
                display:flex;
                flex-wrap:wrap;
                gap:12px;
                margin-top:20px;
            }

            .action-row .btn-tool{
                min-width:160px;
            }

            .preview-success{
                color:#16863a;
                font-weight:700;
            }

            .preview-error{
                color:#c62828;
                font-weight:700;
            }

            .donor-buttons{
                display:flex;
                flex-wrap:wrap;
                gap:9px;
                margin:14px 0;
            }

            .donor-chip{
                border:1px solid #cbd5e1;
                background:#fff;
                border-radius:999px;
                padding:9px 14px;
                font-size:16px;
                cursor:pointer;
            }

            .donor-chip:hover{
                background:#eef6ff;
                border-color:#69a8df;
            }

            .recent-donor-details{
                margin-top:18px;
            }

            .recent-donor-details summary{
                cursor:pointer;
                font-weight:800;
                font-size:18px;
                padding:12px 0;
            }

            .alert{
                margin-bottom:16px;
            }
                                  
            .quick-amount-row{
                display:grid;
                grid-template-columns:58px 1fr 58px;
                gap:10px;
                align-items:stretch;
            }

            .amount-step-btn{
                border:0;
                border-radius:10px;
                background:#e8eef7;
                color:#172033;
                font-size:30px;
                font-weight:800;
                cursor:pointer;
            }

            .amount-step-btn:hover{
                background:#d8e4f4;
            }

            .amount-step-btn:active{
                transform:scale(.96);
            }

            .quick-amount-row .form-input{
                text-align:center;
                font-size:30px;
                font-weight:800;
            }

            .donor-search-box{
                position:relative;
                margin-bottom:14px;
            }

            .donor-search-results{
                display:none;
                position:absolute;
                left:0;
                right:0;
                top:calc(100% + 6px);
                z-index:50;
                max-height:360px;
                overflow:auto;
                background:#fff;
                border:1px solid #cbd5e1;
                border-radius:12px;
                box-shadow:0 12px 30px rgba(15,23,42,.16);
            }

            .donor-search-results.show{
                display:block;
            }

            .donor-result-item{
                width:100%;
                border:0;
                border-bottom:1px solid #eef2f6;
                background:#fff;
                padding:13px 15px;
                text-align:left;
                cursor:pointer;
            }

            .donor-result-item:last-child{
                border-bottom:0;
            }

            .donor-result-item:hover,
            .donor-result-item.active{
                background:#eef6ff;
            }

            .donor-result-main{
                display:flex;
                gap:9px;
                align-items:center;
                flex-wrap:wrap;
                font-size:18px;
                font-weight:800;
                color:#172033;
            }

            .donor-source-badge{
                display:inline-flex;
                align-items:center;
                border-radius:999px;
                padding:3px 9px;
                font-size:13px;
                background:#eaf4ff;
                color:#1769aa;
            }

            .donor-result-meta{
                margin-top:5px;
                color:#667085;
                font-size:14px;
                line-height:1.5;
            }

            .donor-search-empty{
                padding:15px;
                color:#667085;
                text-align:center;
            }

            @media(max-width:700px){
                .finance-form-grid{
                    grid-template-columns:1fr;
                }

                .finance-full{
                    grid-column:auto;
                }

                .receipt-input-row{
                    align-items:stretch;
                }

                .action-row{
                    display:grid;
                }

                .action-row .btn-tool{
                    width:100%;
                }
            }
        </style>
    </head>

    <body>

    <div class="page finance-form-page">

        <div class="finance-back-row">
            <a
                class="btn-tool btn-secondary"
                href="{{ url_for('finance.finance_income_menu') }}"
            >
                ← 返回收入录入
            </a>
        </div>

        <div class="card finance-header-card">

            <h1 class="page-title">
                💵 {{ category }}
            </h1>

            <p class="page-subtitle">
                可批量录入多笔，并在同一批中使用不同收条日期
            </p>

        </div>

        {% if message %}
            <div class="alert alert-danger">
                {{ message }}
            </div>
        {% endif %}

        <form method="post">

            <div class="card">

                <h2 class="section-title">
                    收条与付款资料
                </h2>

                <div class="finance-form-grid">

                    <div class="form-group">

                        <label class="form-label">
                            收条开始号码
                        </label>

                        <div class="receipt-input-row">

                            <div class="receipt-prefix">
                                CHE
                            </div>

                            <input
                                class="form-input"
                                name="receipt_start"
                                value="{{ receipt_start_raw }}"
                                inputmode="numeric"
                                placeholder="例如 1501"
                                required
                            >

                        </div>

                        <span class="field-help">
                            系统建议下一张收条：
                            <strong>{{ next_receipt_no }}</strong>
                        </span>

                    </div>

                    <div class="form-group">

                        <label class="form-label">
                            默认开收条日期
                        </label>

                        <input
                            class="form-input"
                            name="receipt_date"
                            type="date"
                            value="{{ receipt_date }}"
                            required
                        >

                        <span class="field-help">
                            没有写 @日期 的记录，会使用这个日期。
                        </span>

                    </div>

                    <div class="quick-amount-row">

                        <button
                            type="button"
                            class="amount-step-btn"
                            onclick="changeQuickAmount(-50)"
                        >
                            −
                        </button>

                        <input
                            class="form-input"
                            id="quick_amount"
                            name="default_amount"
                            type="number"
                            step="50"
                            min="50"
                            value="{{ default_amount }}"
                            placeholder="金额"
                        >

                        <button
                            type="button"
                            class="amount-step-btn"
                            onclick="changeQuickAmount(50)"
                        >
                            ＋
                        </button>

                    </div>

                    <div class="form-group">

                        <label class="form-label">
                            付款方式
                        </label>

                        <select
                            class="form-input"
                            name="payment_method"
                        >
                            <option
                                {% if payment_method == '现金' %}
                                    selected
                                {% endif %}
                            >
                                现金
                            </option>

                            <option
                                {% if payment_method == '银行过账' %}
                                    selected
                                {% endif %}
                            >
                                银行过账
                            </option>

                            <option
                                {% if payment_method == '支票' %}
                                    selected
                                {% endif %}
                            >
                                支票
                            </option>
                        </select>

                    </div>

                    <div class="form-group finance-full">

                        <label class="form-label">
                            备注／活动名称
                        </label>

                        <input
                            class="form-input"
                            name="remarks"
                            value="{{ remarks }}"
                            placeholder="{% if category == '临时特别布施' %}临时特别布施必须填写活动名称{% else %}例如：观音诞、法会、特别活动{% endif %}"
                        >

                        {% if category == "临时特别布施" %}
                            <span class="field-help">
                                此项目必须填写活动名称或用途。
                            </span>
                        {% endif %}

                    </div>

                </div>

            </div>

            <div class="card">

                <h2 class="section-title">
                    批量输入
                </h2>

                <p class="page-subtitle">
                    可用 @日期 分组；只输入姓名会使用默认金额，
                    姓名后面加金额可单独修改。
                </p>

                <div class="donor-search-box">

                    <label class="form-label">
                        🔎 快速找布施人
                    </label>

                    <input
                        class="form-input"
                        id="donor_search_input"
                        type="search"
                        autocomplete="off"
                        placeholder="输入姓名、英文名、会员／义工编号或电话号码"
                    >

                    <span class="field-help">
                        输入几个字后点选，姓名会自动加入下面的批量输入框。
                    </span>

                    <div
                        class="donor-search-results"
                        id="donor_search_results"
                    ></div>

                </div>

                <textarea
                    class="form-input batch-textarea"
                    name="raw_text"
                    placeholder="例如：
@2026-07-10
王小明
李大华
陈美玲 100

@2026-07-11
郑依颖 30
林美珍 50"
                    required
                >{{ raw_text }}</textarea>

                <div class="batch-example">

                    <strong>输入规则：</strong><br>

                    <code>@2026-07-10</code>
                    → 后续记录使用 2026-07-10<br>

                    <code>王小明</code>
                    → 使用默认金额<br>

                    <code>陈美玲 100</code>
                    → 金额 RM100<br>

                    没有写 <code>@日期</code> 的记录，
                    会使用上面的默认开收条日期。

                </div>

                <div class="action-row">

                    <button
                        class="btn-tool btn-primary"
                        type="submit"
                        name="action"
                        value="preview"
                    >
                        👁️ 预览资料
                    </button>

                    {% if preview_rows %}
                        <button
                            class="btn-tool btn-success"
                            type="submit"
                            name="action"
                            value="confirm"
                            onclick="return confirm('确定全部入账？');"
                        >
                            ✅ 确认全部入账
                        </button>
                    {% endif %}

                </div>

            </div>

            {% if preview_rows %}

                <div class="card">

                    <h2 class="section-title">
                        录入预览
                    </h2>

                    <div class="table-responsive">

                        <table class="record-table">

                            <thead>
                                <tr>
                                    <th>收条</th>
                                    <th>收条日期</th>
                                    <th>姓名</th>
                                    <th>电话</th>
                                    <th>来源</th>
                                    <th>金额</th>
                                    <th>检查结果</th>
                                </tr>
                            </thead>

                            <tbody>

                                {% for r in preview_rows %}
                                <tr>

                                    <td>{{ r.receipt_no or "-" }}</td>

                                    <td>
                                        {{ r.receipt_date or "-" }}
                                    </td>

                                    <td>
                                        <strong>{{ r.name or "-" }}</strong>
                                    </td>

                                    <td>{{ r.phone or "-" }}</td>

                                    <td>{{ r.source or "-" }}</td>

                                    <td>
                                        RM {{ "%.2f"|format(r.amount or 0) }}
                                    </td>

                                    <td>
                                        {% if r.error %}
                                            <span class="preview-error">
                                                ❌ {{ r.error }}
                                            </span>
                                        {% else %}
                                            <span class="preview-success">
                                                ✅ 可以入账
                                            </span>
                                        {% endif %}
                                    </td>

                                </tr>
                                {% endfor %}

                            </tbody>

                        </table>

                    </div>

                </div>

            {% endif %}

        </form>

        {% if recent_donors %}

            <div class="card">

                <h2 class="section-title">
                    常用捐赠者
                </h2>

                <p class="page-subtitle">
                    点一下姓名，会自动加入批量输入框。
                </p>

                <div class="donor-buttons">

                    {% for d in recent_donors %}
                        <button
                            class="donor-chip"
                            type="button"
                            onclick="addDonor({{ d.name|tojson }})"
                        >
                            {{ d.name }}
                        </button>
                    {% endfor %}

                </div>

                <details class="recent-donor-details">

                    <summary>
                        查看捐赠者详细记录
                    </summary>

                    <div class="table-responsive">

                        <table class="record-table">

                            <thead>
                                <tr>
                                    <th>姓名</th>
                                    <th>电话</th>
                                    <th>次数</th>
                                    <th>最后日期</th>
                                </tr>
                            </thead>

                            <tbody>

                                {% for d in recent_donors %}
                                <tr>
                                    <td>{{ d.name }}</td>
                                    <td>{{ d.phone or "-" }}</td>
                                    <td>{{ d.times }}</td>
                                    <td>{{ d.last_date }}</td>
                                </tr>
                                {% endfor %}

                            </tbody>

                        </table>

                    </div>

                </details>

            </div>

        {% endif %}

    </div>

    <script>
const donorSearchInput = document.getElementById("donor_search_input");
const donorSearchResults = document.getElementById("donor_search_results");
let donorSearchTimer = null;
let donorSearchController = null;

function escapeHtml(value){
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function donorIcon(source){
    if(source === "member") return "👤";
    if(source === "volunteer") return "👷";
    return "💖";
}

function closeDonorResults(){
    if(donorSearchResults){
        donorSearchResults.classList.remove("show");
        donorSearchResults.innerHTML = "";
    }
}

function selectDonor(result){
    addDonor(result.name || "");

    if(donorSearchInput){
        donorSearchInput.value = "";
        donorSearchInput.focus();
    }

    closeDonorResults();
}

function renderDonorResults(results){
    if(!donorSearchResults){
        return;
    }

    if(!Array.isArray(results) || results.length === 0){
        donorSearchResults.innerHTML = `
            <div class="donor-search-empty">
                找不到记录，可直接在下面手动输入新姓名。
            </div>
        `;
        donorSearchResults.classList.add("show");
        return;
    }

    donorSearchResults.innerHTML = results.map((r, index) => {
        const code = r.member_id || r.volunteer_id || "";
        const englishName = r.english_name ? ` · ${escapeHtml(r.english_name)}` : "";
        const phone = r.phone ? `电话：${escapeHtml(r.phone)}` : "电话：-";
        const history = r.source === "history"
            ? ` · 最后布施：${escapeHtml(r.last_date || "-")} · ${Number(r.times || 0)} 次`
            : "";

        return `
            <button
                type="button"
                class="donor-result-item"
                data-index="${index}"
            >
                <div class="donor-result-main">
                    <span>${donorIcon(r.source)}</span>
                    <span>${escapeHtml(r.name)}</span>
                    ${code ? `<span>${escapeHtml(code)}</span>` : ""}
                    <span class="donor-source-badge">
                        ${escapeHtml(r.source_label || "资料库")}
                    </span>
                </div>
                <div class="donor-result-meta">
                    ${phone}${englishName}${history}
                </div>
            </button>
        `;
    }).join("");

    donorSearchResults.querySelectorAll(".donor-result-item").forEach((button) => {
        button.addEventListener("click", () => {
            const index = Number(button.dataset.index);
            selectDonor(results[index]);
        });
    });

    donorSearchResults.classList.add("show");
}

async function runDonorSearch(keyword){
    if(!keyword){
        closeDonorResults();
        return;
    }

    if(donorSearchController){
        donorSearchController.abort();
    }

    donorSearchController = new AbortController();

    try{
        const url = new URL(
            {{ url_for('finance.finance_donor_search')|tojson }},
            window.location.origin
        );
        url.searchParams.set("q", keyword);
        url.searchParams.set("branch", "CHE");

        const response = await fetch(url, {
            signal: donorSearchController.signal,
            headers: {"Accept": "application/json"}
        });

        const data = await response.json();

        if(!response.ok || !data.ok){
            throw new Error(data.message || "搜索失败");
        }

        renderDonorResults(data.results || []);
    }catch(error){
        if(error.name === "AbortError"){
            return;
        }

        if(donorSearchResults){
            donorSearchResults.innerHTML = `
                <div class="donor-search-empty">
                    搜索暂时失败，请直接输入姓名。
                </div>
            `;
            donorSearchResults.classList.add("show");
        }
    }
}

if(donorSearchInput){
    donorSearchInput.addEventListener("input", () => {
        const keyword = donorSearchInput.value.trim();
        clearTimeout(donorSearchTimer);

        if(!keyword){
            closeDonorResults();
            return;
        }

        donorSearchTimer = setTimeout(() => {
            runDonorSearch(keyword);
        }, 220);
    });
}

document.addEventListener("click", (event) => {
    if(
        donorSearchInput
        && donorSearchResults
        && !donorSearchInput.contains(event.target)
        && !donorSearchResults.contains(event.target)
    ){
        closeDonorResults();
    }
});

function changeQuickAmount(change){

    const input = document.getElementById(
        "quick_amount"
    );

    if(!input){
        return;
    }

    let current = Number(input.value);

    if(!Number.isFinite(current)){
        current = 50;
    }

    current += change;

    if(current < 50){
        current = 50;
    }

    input.value = current;
}


function addDonor(name){

    const textarea = document.querySelector(
        'textarea[name="raw_text"]'
    );

    if(!textarea){
        return;
    }

    const currentText = textarea.value.trim();

    if(currentText === ""){
        textarea.value = name;
    }else{
        textarea.value += "\\n" + name;
    }

    textarea.focus();
    textarea.scrollTop = textarea.scrollHeight;
}
</script>

    </body>
    </html>
    """,
        category=category,
        message=message,
        raw_text=raw_text,
        receipt_start_raw=receipt_start_raw,
        next_receipt_no=next_receipt_no,
        receipt_date=receipt_date,
        payment_method=payment_method,
        default_amount=default_amount,
        remarks=remarks,
        preview_rows=preview_rows,
        recent_donors=recent_donors
        )

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
    message_type = "bad"

    form_data = {
        "raw_text": "",
        "member_id": "",
        "name": "",
        "amount": "",
        "payment_date": date.today().isoformat(),
        "bank_ref": "",
        "bank_name": "",
        "category": "月费",
        "remarks": "",
    }

    if request.method == "POST":

        raw_text = request.form.get("raw_text", "").strip()

        member_id_raw = request.form.get(
            "member_id",
            ""
        ).strip()

        name = request.form.get(
            "name",
            ""
        ).strip()

        amount_raw = request.form.get(
            "amount",
            ""
        ).strip()

        payment_date = (
            request.form.get("payment_date")
            or date.today().isoformat()
        )

        bank_ref = request.form.get(
            "bank_ref",
            ""
        ).strip()

        bank_name = request.form.get(
            "bank_name",
            ""
        ).strip()

        category = request.form.get(
            "category",
            "月费"
        ).strip()

        remarks = request.form.get(
            "remarks",
            ""
        ).strip()

        if raw_text:

            member_match = re.search(
                r"(CHE|STW)[-\s]?\d+",
                raw_text,
                re.IGNORECASE
            )

            if member_match and not member_id_raw:

                member_id_raw = (
                    member_match
                    .group(0)
                    .replace(" ", "-")
                    .upper()
                )

            amount_match = re.search(
                r"(RM|MYR)\s*([0-9]+(?:\.[0-9]{1,2})?)",
                raw_text,
                re.IGNORECASE
            )

            if amount_match and not amount_raw:
                amount_raw = amount_match.group(2)

            ref_match = re.search(
                r"(Reference|Ref|Transaction|DuitNow)"
                r"\s*(No|ID|Number)?[:\s#-]*"
                r"([A-Za-z0-9\-]+)",
                raw_text,
                re.IGNORECASE
            )

            if ref_match and not bank_ref:
                bank_ref = ref_match.group(3)

            bank_options = [
                "Maybank",
                "Public Bank",
                "CIMB",
                "Hong Leong",
                "RHB",
                "AmBank",
                "Bank Islam",
                "BSN",
            ]

            for bank in bank_options:

                if (
                    bank.lower() in raw_text.lower()
                    and not bank_name
                ):
                    bank_name = bank
                    break

            if not remarks:
                remarks = raw_text[:1000]

        member_id = (
            normalize_member_id(member_id_raw)
            if member_id_raw
            else ""
        )

        amount = money(amount_raw)

        if category == "月费" and member_id:

            member = db_query("""
                select *
                from members
                where member_id = %s
                limit 1
            """, (
                member_id,
            ), fetchone=True)

            if member:
                name = (
                    member.get("name")
                    or member.get("姓名")
                    or name
                )

        form_data = {
            "raw_text": raw_text,
            "member_id": member_id_raw,
            "name": name,
            "amount": amount_raw,
            "payment_date": str(payment_date),
            "bank_ref": bank_ref,
            "bank_name": bank_name,
            "category": category,
            "remarks": remarks,
        }

        existing_ref = None

        if bank_ref:

            existing_ref = db_query("""
                select id
                from bank_pending_records
                where bank_ref = %s
                  and coalesce(bank_ref, '') <> ''
                limit 1
            """, (
                bank_ref,
            ), fetchone=True)

        if existing_ref:

            message = (
                "这个 Bank Reference 已经存在，"
                "请检查是否重复导入。"
            )

        elif amount <= 0:

            message = "请输入正确的银行过账金额。"

        else:

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

            return redirect(
                url_for("finance.bank_pending")
            )

    summary = db_query("""
        select
            count(*) as cnt,
            coalesce(sum(amount), 0) as total
        from bank_pending_records
        where status = 'pending'
    """, fetchone=True)

    rows = db_query("""
        select *
        from bank_pending_records
        where status = 'pending'
        order by upload_date desc
    """, fetchall=True)

    pending_count = int(
        summary["cnt"] or 0
    )

    pending_total = float(
        summary["total"] or 0
    )

    return render_template_string(FINANCE_DATE_COMPONENT + """
    <!doctype html>
    <html lang="zh">
    <head>

        <meta charset="utf-8">

        <meta
            name="viewport"
            content="width=device-width, initial-scale=1"
        >

        <title>银行过账中心</title>

        <link
            rel="stylesheet"
            href="{{ url_for(
                'static',
                filename='css/toolbox.css'
            ) }}"
        >

        <style>

            .bank-page {
                max-width: 1450px;
            }

            .bank-header {
                background:
                    linear-gradient(
                        135deg,
                        #2563eb,
                        #1d4ed8
                    );

                color: white;
                padding: 28px;
                border-radius: 22px;
                margin-bottom: 20px;

                box-shadow:
                    0 12px 30px
                    rgba(37, 99, 235, 0.18);
            }

            .bank-header h1 {
                margin: 0 0 8px;
                font-size: 30px;
            }

            .bank-header p {
                margin: 0;
                opacity: 0.92;
                line-height: 1.6;
            }

            .bank-summary {
                display: grid;
                grid-template-columns:
                    repeat(2, minmax(0, 1fr));
                gap: 16px;
                margin-bottom: 20px;
            }

            .bank-summary .summary-box {
                min-height: 125px;
                text-align: center;

                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
            }

            .summary-icon {
                font-size: 30px;
                margin-bottom: 6px;
            }

            .summary-label {
                color: #64748b;
                font-size: 16px;
                margin-bottom: 6px;
            }

            .summary-value {
                font-size: 28px;
                font-weight: 800;
                color: #0f172a;
            }

            .summary-value.money-value {
                color: #15803d;
            }

            .entry-grid {
                display: grid;
                grid-template-columns:
                    minmax(0, 1.2fr)
                    minmax(360px, 0.8fr);
                gap: 20px;
                align-items: start;
            }

            .form-grid {
                display: grid;
                grid-template-columns:
                    repeat(2, minmax(0, 1fr));
                gap: 16px;
            }

            .form-grid .full-width {
                grid-column: 1 / -1;
            }

            .receipt-textarea {
                width: 100%;
                min-height: 245px;
                resize: vertical;
                line-height: 1.6;
            }

            .form-help {
                margin-top: 7px;
                color: #64748b;
                font-size: 14px;
                line-height: 1.5;
            }

            .smart-note {
                background: #eff6ff;
                border: 1px solid #bfdbfe;
                color: #1e40af;
                border-radius: 14px;
                padding: 14px 16px;
                margin-bottom: 16px;
                line-height: 1.6;
            }

            .records-card {
                margin-top: 20px;
            }

            .table-topbar {
                display: flex;
                justify-content: space-between;
                align-items: center;
                gap: 12px;
                flex-wrap: wrap;
                margin-bottom: 14px;
            }

            .record-count {
                color: #64748b;
                font-size: 16px;
            }

            .pending-table {
                min-width: 1400px;
            }

            .pending-table td {
                vertical-align: middle;
            }

            .member-id {
                font-weight: 800;
                color: #1d4ed8;
                white-space: nowrap;
            }

            .member-name {
                font-weight: 700;
                color: #1e293b;
                min-width: 90px;
            }

            .money-cell {
                color: #15803d;
                font-weight: 800;
                white-space: nowrap;
                text-align: right;
            }

            .reference-cell {
                font-family: Consolas, monospace;
                font-size: 14px;
                white-space: nowrap;
            }

            .remarks-cell {
                min-width: 180px;
                max-width: 320px;
                white-space: normal;
                line-height: 1.5;
                overflow-wrap: anywhere;
            }

            .confirm-box {
                min-width: 285px;
                background: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 14px;
                padding: 12px;
            }

            .confirm-grid {
                display: grid;
                gap: 10px;
            }

            .confirm-box .form-group {
                margin: 0;
            }

            .confirm-box .form-label {
                font-size: 14px;
            }

            .confirm-box .form-input {
                min-height: 41px;
                padding: 8px 10px;
                font-size: 14px;
                margin: 0;
            }

            .confirm-box .btn-tool {
                width: 100%;
                min-height: 42px;
                font-size: 14px;
                padding: 9px 12px;
            }

            .delete-form .btn-tool {
                min-height: 42px;
                padding: 9px 13px;
                font-size: 14px;
                white-space: nowrap;
            }

            .empty-bank {
                text-align: center;
                padding: 55px 20px;
                color: #64748b;
            }

            .empty-bank-icon {
                font-size: 48px;
                margin-bottom: 10px;
            }

            .empty-bank h3 {
                margin: 0 0 8px;
                color: #334155;
            }

            .bottom-actions {
                margin-top: 20px;
            }

            @media (max-width: 950px) {

                .entry-grid {
                    grid-template-columns: 1fr;
                }

            }

            @media (max-width: 700px) {

                .bank-page {
                    padding-left: 12px;
                    padding-right: 12px;
                }

                .bank-header {
                    padding: 22px 18px;
                    border-radius: 18px;
                }

                .bank-header h1 {
                    font-size: 26px;
                }

                .bank-summary {
                    grid-template-columns: 1fr;
                }

                .bank-summary .summary-box {
                    min-height: 105px;
                }

                .form-grid {
                    grid-template-columns: 1fr;
                }

                .form-grid .full-width {
                    grid-column: auto;
                }

                .entry-grid {
                    display: block;
                }

                .entry-grid .card + .card {
                    margin-top: 20px;
                }

                .bottom-actions .btn-tool {
                    width: 100%;
                }

            }

        </style>

    </head>

    <body>

    <div class="page bank-page">

        <div class="bank-header">

            <h1>🏦 银行过账中心</h1>

            <p>
                先登记银行转账资料，财政确认后再填写收条号码并正式入账。
            </p>

        </div>

        {% if message %}

            <div class="alert alert-danger">
                ⚠️ {{ message }}
            </div>

        {% endif %}

        <div class="bank-summary">

            <div class="summary-box">

                <div class="summary-icon">
                    ⏳
                </div>

                <div class="summary-label">
                    待确认笔数
                </div>

                <div class="summary-value">
                    {{ pending_count }}
                </div>

            </div>

            <div class="summary-box">

                <div class="summary-icon">
                    💰
                </div>

                <div class="summary-label">
                    待确认总额
                </div>

                <div class="summary-value money-value">
                    RM {{ "%.2f"|format(pending_total) }}
                </div>

            </div>

        </div>

        <form method="post">

            <div class="entry-grid">

                <div class="card">

                    <div class="section-title">
                        📄 银行资料智能识别
                    </div>

                    <div class="smart-note">
                        可以直接粘贴 WhatsApp 信息或银行转账
                        Receipt 文字。系统会尝试识别会员编号、
                        金额、银行及 Reference。
                    </div>

                    <div class="form-group">

                        <label class="form-label">
                            项目
                        </label>

                        <select
                            class="form-input"
                            name="category"
                        >

                            {% set categories = [
                                '月费',
                                '财布施',
                                '观音村',
                                '膳食结缘',
                                '观音堂纯檀香布施',
                                special_donation_title,
                                '临时特别布施'
                            ] %}

                            {% for category_item in categories %}

                                <option
                                    value="{{ category_item }}"
                                    {% if
                                        form_data.category
                                        == category_item
                                    %}
                                        selected
                                    {% endif %}
                                >
                                    {{ category_item }}
                                </option>

                            {% endfor %}

                        </select>

                    </div>

                    <div class="form-group">

                        <label class="form-label">
                            粘贴 WhatsApp／银行 Receipt 文字
                        </label>

                        <textarea
                            class="form-input receipt-textarea"
                            name="raw_text"
                            placeholder="例如：

CHE-108
RM50.00
Maybank
Reference：ABC123456"
                        >{{ form_data.raw_text }}</textarea>

                        <div class="form-help">
                            系统识别后，仍可在右边检查和修改资料。
                        </div>

                    </div>

                </div>

                <div class="card">

                    <div class="section-title">
                        ✏️ 手动补充／修正资料
                    </div>

                    <div class="form-grid">

                        <div class="form-group">

                            <label class="form-label">
                                会员编号
                            </label>

                            <input
                                class="form-input"
                                name="member_id"
                                value="{{ form_data.member_id }}"
                                placeholder="CHE-108 / STW-108 / 108"
                                autocomplete="off"
                            >

                        </div>

                        <div class="form-group">

                            <label class="form-label">
                                姓名
                            </label>

                            <input
                                class="form-input"
                                name="name"
                                value="{{ form_data.name }}"
                                placeholder="非月费或找不到会员时填写"
                                autocomplete="off"
                            >

                        </div>

                        <div class="form-group">

                            <label class="form-label">
                                金额 RM
                            </label>

                            <input
                                class="form-input"
                                name="amount"
                                type="number"
                                step="0.01"
                                min="0.01"
                                value="{{ form_data.amount }}"
                                placeholder="例如 50.00"
                                required
                            >

                        </div>

                        <div class="form-group">

                            <label class="form-label">
                                付款日期
                            </label>

                            <input
                                class="form-input"
                                name="payment_date"
                                type="date"
                                value="{{ form_data.payment_date }}"
                                required
                            >

                        </div>

                        <div class="form-group">

                            <label class="form-label">
                                银行 Reference
                            </label>

                            <input
                                class="form-input"
                                name="bank_ref"
                                value="{{ form_data.bank_ref }}"
                                placeholder="Transaction Reference"
                                autocomplete="off"
                            >

                        </div>

                        <div class="form-group">

                            <label class="form-label">
                                银行
                            </label>

                            <select
                                class="form-input"
                                name="bank_name"
                            >

                                <option value="">
                                    请选择银行
                                </option>

                                {% for bank in bank_names %}

                                    <option
                                        value="{{ bank }}"
                                        {% if
                                            form_data.bank_name == bank
                                        %}
                                            selected
                                        {% endif %}
                                    >
                                        {{ bank }}
                                    </option>

                                {% endfor %}

                            </select>

                        </div>

                        <div class="form-group full-width">

                            <label class="form-label">
                                备注
                            </label>

                            <textarea
                                class="form-input"
                                name="remarks"
                                rows="4"
                                placeholder="填写补充说明或保留银行转账文字"
                            >{{ form_data.remarks }}</textarea>

                        </div>

                    </div>

                    <button
                        class="btn-tool btn-primary"
                        type="submit"
                        style="width:100%;"
                    >
                        ➕ 加入待确认
                    </button>

                </div>

            </div>

        </form>

        <div class="card records-card">

            <div class="table-topbar">

                <div
                    class="section-title"
                    style="margin-bottom:0;"
                >
                    📋 待确认列表
                </div>

                <div class="record-count">
                    共
                    <strong>{{ pending_count }}</strong>
                    笔待确认记录
                </div>

            </div>

            {% if rows %}

                <div class="table-responsive">

                    <table class="record-table pending-table">

                        <thead>

                            <tr>
                                <th>付款日期</th>
                                <th>编号</th>
                                <th>姓名</th>
                                <th>项目</th>
                                <th>金额</th>
                                <th>Reference</th>
                                <th>银行</th>
                                <th>备注</th>
                                <th>确认入账</th>
                                <th>删除</th>
                            </tr>

                        </thead>

                        <tbody>

                            {% for r in rows %}

                                <tr>

                                    <td style="white-space:nowrap;">
                                        {{ r.payment_date }}
                                    </td>

                                    <td>
                                        <span class="member-id">
                                            {{ r.member_id or "-" }}
                                        </span>
                                    </td>

                                    <td>
                                        <span class="member-name">
                                            {{ r.name or "-" }}
                                        </span>
                                    </td>

                                    <td>
                                        {{ r.category or "-" }}
                                    </td>

                                    <td class="money-cell">
                                        RM
                                        {{ "%.2f"|format(
                                            r.amount or 0
                                        ) }}
                                    </td>

                                    <td class="reference-cell">
                                        {{ r.bank_ref or "-" }}
                                    </td>

                                    <td style="white-space:nowrap;">
                                        {{ r.bank_name or "-" }}
                                    </td>

                                    <td class="remarks-cell">
                                        {{ r.remarks or "-" }}
                                    </td>

                                    <td>

                                        <div class="confirm-box">

                                            <form
                                                method="post"
                                                action="{{ url_for(
                                                    'finance.confirm_bank',
                                                    pending_id=r.id
                                                ) }}"
                                                onsubmit="
                                                    return confirm(
                                                        '确定确认入账？'
                                                    );
                                                "
                                            >

                                                <div class="confirm-grid">

                                                    <div class="form-group">

                                                        <label
                                                            class="form-label"
                                                        >
                                                            收条号码
                                                        </label>

                                                        <input
                                                            class="form-input"
                                                            name="receipt_no"
                                                            placeholder="CHE0000001"
                                                            autocomplete="off"
                                                            required
                                                        >

                                                    </div>

                                                    <div class="form-group">

                                                        <label
                                                            class="form-label"
                                                        >
                                                            开收条日期
                                                        </label>

                                                        <input
                                                            class="form-input"
                                                            name="receipt_date"
                                                            type="date"
                                                            value="{{ today }}"
                                                            required
                                                        >

                                                    </div>

                                                    <button
                                                        class="
                                                            btn-tool
                                                            btn-success
                                                        "
                                                        type="submit"
                                                    >
                                                        ✅ 确认入账
                                                    </button>

                                                </div>

                                            </form>

                                        </div>

                                    </td>

                                    <td>

                                        <form
                                            class="delete-form"
                                            method="post"
                                            action="{{ url_for(
                                                'finance.delete_bank_pending',
                                                pending_id=r.id
                                            ) }}"
                                            onsubmit="
                                                return confirm(
                                                    '确定删除这笔待确认记录？'
                                                );
                                            "
                                        >

                                            <button
                                                class="
                                                    btn-tool
                                                    btn-danger
                                                "
                                                type="submit"
                                            >
                                                🗑️ 删除
                                            </button>

                                        </form>

                                    </td>

                                </tr>

                            {% endfor %}

                        </tbody>

                    </table>

                </div>

            {% else %}

                <div class="empty-bank">

                    <div class="empty-bank-icon">
                        ✅
                    </div>

                    <h3>目前没有待确认记录</h3>

                    <p>
                        所有银行转账记录都已经处理完成。
                    </p>

                </div>

            {% endif %}

        </div>

        <div class="bottom-actions">

            <a
                class="btn-tool btn-secondary"
                href="{{ url_for(
                    'finance.finance_home'
                ) }}"
            >
                ← 返回财政首页
            </a>

        </div>

    </div>

    </body>
    </html>
    """,
        rows=rows,
        summary=summary,
        message=message,
        today=date.today().isoformat(),
        pending_count=pending_count,
        pending_total=pending_total,
        form_data=form_data,
        special_donation_title=SPECIAL_DONATION_TITLE,
        bank_names=[
            "Maybank",
            "Public Bank",
            "CIMB",
            "Hong Leong",
            "RHB",
            "AmBank",
            "Bank Islam",
            "BSN",
        ]
    )

@finance_bp.route("/bank_pending/<int:pending_id>/confirm", methods=["POST"])
def confirm_bank(pending_id):

    receipt_no = request.form.get("receipt_no", "").strip().upper()
    receipt_date = request.form.get("receipt_date") or date.today()

    if not receipt_no:
        return "必须填写收条号码", 400

    existing = db_query("""
        select id
        from finance_records
        where receipt_no = %s
        limit 1
    """, (receipt_no,), fetchone=True)

    if existing:
        return "这个收条号码已经用过了，请检查是否重复输入", 400

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
              and coalesce(status, 'active') = 'active'
        """, (p["member_id"],), fetchone=True)

        paid_until_date = paid["paid_until"] if paid else None

        month_from = next_month_ym(paid_until_date)

        month_count = max(1, round(float(p["amount"] or 0) / 50))
        month_to = add_months_ym(month_from, month_count - 1)

        month_from_db = month_from + "-01"
        month_to_db = month_to + "-01"

    pending_fund_account = get_fund_account(
        p["category"],
        "income"
    )

    month_lock_error = require_finance_month_open(
        receipt_date,
        pending_fund_account
    )

    if month_lock_error:
        flash(month_lock_error, "danger")
        return redirect(
            url_for("finance.bank_pending")
        )

    db_query("""
        insert into finance_records
        (
            record_type,
            fund_account,
            record_date,
            receipt_date,
            category,
            receipt_no,
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
            %s, %s, %s, %s, %s,
            %s, %s, %s, %s,
            '银行过账',
            %s, %s, %s, %s, %s
        )
    """, (
        "income",
        pending_fund_account,
        p["payment_date"],
        receipt_date,
        p["category"],
        receipt_no,
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
                receipt_date,
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
                %s, %s, %s, %s, %s,
                %s, %s, %s, %s
            )
        """, (
            p["payment_date"],
            receipt_date,
            p["member_id"],
            p["name"],
            receipt_no,
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

@finance_bp.route(
    "/records/<int:record_id>/cancel",
    methods=["POST"]
)
def cancel_record(record_id):

    # 先读取完整资料
    record = db_query("""
        select
            id,
            receipt_no,
            record_date,
            fund_account,
            status
        from finance_records
        where id = %s
        limit 1
    """, (
        record_id,
    ), fetchone=True)

    if not record:

        flash(
            "找不到这笔财政记录。",
            "danger"
        )

        return redirect(
            url_for("finance.records")
        )

    # 已经作废，不可重复操作
    if record.get("status") == "cancelled":

        flash(
            "这笔财政记录已经作废。",
            "warning"
        )

        return redirect(
            url_for("finance.records")
        )

    # =====================================================
    # V6：Month Close Lock
    # 已月结月份禁止作废
    # =====================================================
    month_lock_error = require_finance_month_open(
        record["record_date"],
        record.get("fund_account")
    )

    if month_lock_error:

        flash(
            month_lock_error,
            "danger"
        )

        return redirect(
            url_for("finance.records")
        )

    cancel_reason = request.form.get(
        "cancel_reason",
        ""
    ).strip()

    if not cancel_reason:

        flash(
            "请填写作废原因。",
            "danger"
        )

        return redirect(
            url_for("finance.records")
        )

    # 暂时使用 session 中的财政用户
    # 如果没有登录名字，则显示 Finance User
    cancelled_by = get_current_finance_user()

    # =====================================================
    # 作废 finance_records
    # 不再 DELETE
    # =====================================================
    db_query("""
        update finance_records
        set
            status = 'cancelled',
            cancel_reason = %s,
            cancelled_by = %s,
            cancelled_at = now()
        where id = %s
          and coalesce(status, 'confirmed') != 'cancelled'
    """, (
        cancel_reason,
        cancelled_by,
        record_id,
    ))

    # =====================================================
    # 若这笔记录有收条编号，
    # 同步作废 member_payments
    # 不再永久删除月费记录
    # =====================================================
    if record.get("receipt_no"):

        db_query("""
            update member_payments
            set
                status = 'cancelled',
                cancel_reason = %s,
                cancelled_by = %s,
                cancelled_at = now()
            where receipt_no = %s
              and coalesce(status, 'active') != 'cancelled'
        """, (
            cancel_reason,
            cancelled_by,
            record["receipt_no"],
        ))

    updated_record = db_query("""
        select *
        from finance_records
        where id = %s
        limit 1
    """, (record_id,), fetchone=True)

    write_finance_audit(
        module="finance_records",
        action="cancel",
        record_id=record_id,
        old_value=dict(record),
        new_value=dict(updated_record) if updated_record else None,
        reason=cancel_reason,
        actor=cancelled_by,
    )

    flash(
        "财政记录已成功作废，原始资料仍保留在系统中。",
        "success"
    )

    return redirect(
        url_for("finance.records")
    )


@finance_bp.route("/records")
def records():

    q = request.args.get("q", "").strip()
    date_from = request.args.get("date_from", "").strip()
    date_to = request.args.get("date_to", "").strip()
    record_type = request.args.get("record_type", "").strip()
    category = request.args.get("category", "").strip()
    number_type = request.args.get("number_type", "").strip().upper()
    status_filter = request.args.get("status", "").strip()
    category = request.args.get("category", "").strip()

    where_parts = []
    params = []

    if q:
        keyword = f"%{q}%"
        where_parts.append("""
            (
                coalesce(receipt_no, '') ilike %s
                or coalesce(payment_voucher_no, '') ilike %s
                or coalesce(member_id, '') ilike %s
                or coalesce(name, '') ilike %s
                or coalesce(phone, '') ilike %s
                or coalesce(category, '') ilike %s
                or coalesce(payment_method, '') ilike %s
                or coalesce(bank_ref, '') ilike %s
                or coalesce(remarks, '') ilike %s
                or coalesce(fund_account, '') ilike %s
            )
        """)
        params.extend([keyword] * 10)

    if date_from:
        where_parts.append("record_date >= %s")
        params.append(date_from)

    if date_to:
        where_parts.append("record_date <= %s")
        params.append(date_to)

    if record_type in ["income", "expense"]:
        where_parts.append("record_type = %s")
        params.append(record_type)

    if category:
        where_parts.append("category = %s")
        params.append(category)

    if number_type == "CHE":
        where_parts.append(
            "coalesce(receipt_no, '') ilike %s"
        )
        params.append("CHE%")

    elif number_type == "STW":
        where_parts.append(
            "coalesce(receipt_no, '') ilike %s"
        )
        params.append("STW%")

    elif number_type == "PV":
        where_parts.append("""
            coalesce(payment_voucher_no, '') <> ''
        """)

    elif number_type == "CDM":
        where_parts.append("""
            coalesce(bank_ref, '') ilike %s
        """)
        params.append("CDM-%")

    if status_filter == "active":
        where_parts.append("""
            coalesce(status, 'confirmed') <> 'cancelled'
        """)

    elif status_filter == "cancelled":
        where_parts.append("status = 'cancelled'")

    where_sql = ""

    if where_parts:
        where_sql = " where " + " and ".join(where_parts)

    rows = db_query(
        f"""
            select *
            from finance_records
            {where_sql}
            order by record_date desc, id desc
            limit 500
        """,
        tuple(params),
        fetchall=True
    )

    summary = db_query(
        f"""
            select
                count(*) as shown_count,

                coalesce(sum(
                    case
                        when record_type = 'income'
                         and coalesce(status, 'confirmed') <> 'cancelled'
                        then amount
                        else 0
                    end
                ), 0) as income_total,

                coalesce(sum(
                    case
                        when record_type = 'expense'
                         and coalesce(status, 'confirmed') <> 'cancelled'
                        then amount
                        else 0
                    end
                ), 0) as expense_total,

                coalesce(sum(
                    case
                        when coalesce(status, 'confirmed') <> 'cancelled'
                        then 1
                        else 0
                    end
                ), 0) as active_count,

                coalesce(sum(
                    case
                        when status = 'cancelled'
                        then 1
                        else 0
                    end
                ), 0) as cancelled_count
            from finance_records
            {where_sql}
        """,
        tuple(params),
        fetchone=True
    )

    categories = db_query("""
        select distinct category
        from finance_records
        where category is not null
          and trim(category) <> ''
        order by category
    """, fetchall=True)

    shown_count = int(summary["shown_count"] or 0)
    active_count = int(summary["active_count"] or 0)
    cancelled_count = int(summary["cancelled_count"] or 0)
    income_total = float(summary["income_total"] or 0)
    expense_total = float(summary["expense_total"] or 0)
    balance_total = income_total - expense_total

    return render_template_string(FINANCE_DATE_COMPONENT + """
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">

        <meta
            name="viewport"
            content="width=device-width, initial-scale=1"
        >

        <title>财政总账</title>

        <link
            rel="stylesheet"
            href="{{ url_for('static', filename='css/toolbox.css') }}"
        >

        <style>
            .records-page{
                max-width:1500px;
            }

            .records-header{
                background:linear-gradient(135deg,#2563eb,#1d4ed8);
                color:#fff;
                padding:28px;
                border-radius:22px;
                margin-bottom:20px;
                box-shadow:0 12px 30px rgba(37,99,235,.18);
            }

            .records-header h1{
                margin:0 0 8px;
                font-size:32px;
            }

            .records-header p{
                margin:0;
                opacity:.92;
                line-height:1.6;
            }

            .quick-filter-row{
                display:flex;
                flex-wrap:wrap;
                gap:10px;
                margin-top:18px;
            }

            .quick-filter-btn{
                display:inline-flex;
                align-items:center;
                justify-content:center;
                min-height:42px;
                padding:8px 14px;
                border-radius:999px;
                background:rgba(255,255,255,.16);
                color:#fff;
                text-decoration:none;
                font-weight:800;
                border:1px solid rgba(255,255,255,.28);
            }

            .quick-filter-btn:hover{
                background:rgba(255,255,255,.25);
            }

            .filter-grid{
                display:grid;
                grid-template-columns:repeat(4,minmax(0,1fr));
                gap:14px;
            }

            .filter-grid .form-group{
                margin:0;
            }

            .filter-full{
                grid-column:1 / -1;
            }

            .filter-actions{
                display:flex;
                flex-wrap:wrap;
                gap:10px;
                margin-top:18px;
            }

            .records-summary{
                display:grid;
                grid-template-columns:repeat(5,minmax(0,1fr));
                gap:14px;
                margin:20px 0;
            }

            .records-summary .summary-box{
                text-align:center;
                min-height:112px;
                display:flex;
                flex-direction:column;
                align-items:center;
                justify-content:center;
            }

            .summary-label{
                color:#64748b;
                font-size:15px;
                margin-bottom:7px;
            }

            .summary-value{
                color:#0f172a;
                font-size:26px;
                font-weight:900;
            }

            .summary-income{
                color:#15803d;
            }

            .summary-expense{
                color:#b91c1c;
            }

            .summary-balance{
                color:#1d4ed8;
            }

            .table-topbar{
                display:flex;
                align-items:center;
                justify-content:space-between;
                flex-wrap:wrap;
                gap:12px;
                margin-bottom:14px;
            }

            .result-text{
                color:#64748b;
                font-size:16px;
            }

            .record-table{
                min-width:1320px;
            }

            .record-table th{
                white-space:nowrap;
            }

            .record-table td{
                vertical-align:middle;
            }

            .record-link{
                color:#1d4ed8;
                font-weight:900;
                text-decoration:none;
                white-space:nowrap;
            }

            .record-link:hover{
                text-decoration:underline;
            }

            .type-badge,
            .number-badge,
            .status-badge{
                display:inline-flex;
                align-items:center;
                justify-content:center;
                padding:7px 11px;
                border-radius:999px;
                font-size:14px;
                font-weight:900;
                white-space:nowrap;
            }

            .type-income{
                background:#dcfce7;
                color:#166534;
            }

            .type-expense{
                background:#fee2e2;
                color:#991b1b;
            }

            .number-che{
                background:#dbeafe;
                color:#1d4ed8;
            }

            .number-stw{
                background:#fee2e2;
                color:#b91c1c;
            }

            .number-pv{
                background:#f3e8ff;
                color:#7e22ce;
            }

            .status-confirmed{
                background:#dcfce7;
                color:#166534;
            }

            .status-cancelled{
                background:#fee2e2;
                color:#991b1b;
            }

            .status-pending{
                background:#fef3c7;
                color:#92400e;
            }

            .member-name{
                font-weight:800;
                color:#1e293b;
            }

            .member-id{
                color:#64748b;
                font-size:14px;
                margin-top:3px;
            }

            .money{
                font-weight:900;
                white-space:nowrap;
                text-align:right;
            }

            .money-income{
                color:#15803d;
            }

            .money-expense{
                color:#b91c1c;
            }

            .month-purpose{
                min-width:180px;
                line-height:1.5;
            }

            .remarks-cell{
                min-width:150px;
                max-width:250px;
                white-space:normal;
                line-height:1.5;
            }

            .cancelled-row{
                background:#f8fafc;
                color:#94a3b8;
            }

            .cancelled-row td:not(.action-cell){
                text-decoration:line-through;
            }

            .cancelled-row .status-badge,
            .cancelled-row .record-link,
            .cancelled-row .type-badge,
            .cancelled-row .number-badge{
                text-decoration:none;
            }

            .row-actions{
                display:flex;
                align-items:center;
                gap:8px;
                white-space:nowrap;
            }

            .mini-action{
                min-height:38px;
                padding:8px 12px;
                font-size:14px;
            }

            .cancel-form{
                display:flex;
                gap:8px;
                align-items:center;
            }

            .cancel-form .form-input{
                min-width:150px;
                min-height:38px;
                padding:7px 9px;
                font-size:14px;
                margin:0;
            }

            .empty-state-custom{
                text-align:center;
                padding:55px 20px;
                color:#64748b;
            }

            .bottom-actions{
                margin-top:20px;
            }

            @media(max-width:900px){
                .filter-grid{
                    grid-template-columns:1fr 1fr;
                }

                .records-summary{
                    grid-template-columns:1fr 1fr;
                }
            }

            @media(max-width:620px){
                .records-page{
                    padding-left:12px;
                    padding-right:12px;
                }

                .records-header{
                    padding:22px 18px;
                    border-radius:18px;
                }

                .records-header h1{
                    font-size:27px;
                }

                .filter-grid,
                .records-summary{
                    grid-template-columns:1fr;
                }

                .filter-full{
                    grid-column:auto;
                }

                .filter-actions{
                    display:grid;
                }

                .filter-actions .btn-tool,
                .bottom-actions .btn-tool{
                    width:100%;
                }
            }
        </style>
    </head>

    <body>

    <div class="page records-page">

        <div class="records-header">
            <h1>📚 财政总账</h1>

            <p>
                收条、Payment Voucher、收入、支出、月费月份和各类布施，都可在这里查询。
            </p>

            <div class="quick-filter-row">

                <a class="quick-filter-btn"
                href="{{ url_for('finance.records') }}">
                    全部
                </a>

                <a class="quick-filter-btn"
                href="{{ url_for('finance.records', record_type='income') }}">
                    收入
                </a>

                <a class="quick-filter-btn"
                href="{{ url_for('finance.records', record_type='expense') }}">
                    支出
                </a>

                <a class="quick-filter-btn"
                href="{{ url_for('finance.records', category='月费') }}">
                    月费
                </a>

                <a class="quick-filter-btn"
                href="{{ url_for('finance.records', category='财布施') }}">
                    财布施
                </a>

                <a class="quick-filter-btn"
                href="{{ url_for('finance.records', category='观音村') }}">
                    观音村
                </a>

                <a class="quick-filter-btn"
                href="{{ url_for('finance.records', category='膳食结缘') }}">
                    初一十五
                </a>

                <a class="quick-filter-btn"
                href="{{ url_for('finance.records', number_type='PV') }}">
                    PV
                </a>

                <a class="quick-filter-btn"
                href="{{ url_for('finance.records', number_type='CDM') }}">
                    CDM
                </a>

            </div>
        </div>

        <div class="card">
            <div class="section-title">
                🔎 查询条件
            </div>

            <form method="get">
                <div class="filter-grid">

                    <div class="form-group">
                        <label class="form-label">开始日期</label>

                        <input
                            class="form-input"
                            type="date"
                            name="date_from"
                            value="{{ date_from }}"
                        >
                    </div>

                    <div class="form-group">
                        <label class="form-label">结束日期</label>

                        <input
                            class="form-input"
                            type="date"
                            name="date_to"
                            value="{{ date_to }}"
                        >
                    </div>

                    <div class="form-group">
                        <label class="form-label">记录类型</label>

                        <select class="form-input" name="record_type">
                            <option value="">全部</option>
                            <option value="income"
                                {% if record_type == 'income' %}selected{% endif %}>
                                收入
                            </option>
                            <option value="expense"
                                {% if record_type == 'expense' %}selected{% endif %}>
                                支出
                            </option>
                        </select>
                    </div>

                    <div class="form-group">
                        <label class="form-label">编号类型</label>

                        <select class="form-input" name="number_type">
                            <option value="">全部</option>
                            <option value="CHE"
                                {% if number_type == 'CHE' %}selected{% endif %}>
                                CHE 收条
                            </option>
                            <option value="STW"
                                {% if number_type == 'STW' %}selected{% endif %}>
                                STW 收条
                            </option>
                            <option value="PV"
                                {% if number_type == 'PV' %}selected{% endif %}>
                                Payment Voucher
                            </option>
                        </select>
                    </div>

                    <div class="form-group">
                        <label class="form-label">类别</label>

                        <select class="form-input" name="category">
                            <option value="">全部类别</option>

                            {% for c in categories %}
                                <option
                                    value="{{ c.category }}"
                                    {% if category == c.category %}selected{% endif %}
                                >
                                    {{ c.category }}
                                </option>
                            {% endfor %}
                        </select>
                    </div>

                    <div class="form-group">
                        <label class="form-label">状态</label>

                        <select class="form-input" name="status">
                            <option value="">全部状态</option>
                            <option value="active"
                                {% if status_filter == 'active' %}selected{% endif %}>
                                正常记录
                            </option>
                            <option value="cancelled"
                                {% if status_filter == 'cancelled' %}selected{% endif %}>
                                已作废
                            </option>
                        </select>
                    </div>

                    <div class="form-group filter-full">
                        <label class="form-label">关键字</label>

                        <input
                            class="form-input"
                            name="q"
                            value="{{ q }}"
                            placeholder="收条、PV、会员编号、姓名、电话、类别、付款方式、Reference、备注"
                            autocomplete="off"
                        >
                    </div>
                </div>

                <div class="filter-actions">
                    <button
                        class="btn-tool btn-primary"
                        type="submit"
                    >
                        🔍 查询记录
                    </button>

                    <a
                        class="btn-tool btn-secondary"
                        href="{{ url_for('finance.records') }}"
                    >
                        ✕ 清除条件
                    </a>
                </div>
            </form>
        </div>

        <div class="records-summary">
            <div class="summary-box">
                <div class="summary-label">筛选后收入</div>
                <div class="summary-value summary-income">
                    RM {{ "%.2f"|format(income_total) }}
                </div>
            </div>

            <div class="summary-box">
                <div class="summary-label">筛选后支出</div>
                <div class="summary-value summary-expense">
                    RM {{ "%.2f"|format(expense_total) }}
                </div>
            </div>

            <div class="summary-box">
                <div class="summary-label">筛选后结余</div>
                <div class="summary-value summary-balance">
                    RM {{ "%.2f"|format(balance_total) }}
                </div>
            </div>

            <div class="summary-box">
                <div class="summary-label">正常记录</div>
                <div class="summary-value">
                    {{ active_count }}
                </div>
            </div>

            <div class="summary-box">
                <div class="summary-label">已作废</div>
                <div class="summary-value">
                    {{ cancelled_count }}
                </div>
            </div>
        </div>

        <div class="card">
            <div class="table-topbar">
                <div class="section-title" style="margin-bottom:0;">
                    🧾 记录明细
                </div>

                <div class="result-text">
                    当前找到
                    <strong>{{ shown_count }}</strong>
                    笔记录
                </div>
            </div>

            {% if rows %}
                <div class="table-responsive">
                    <table class="record-table">
                        <thead>
                            <tr>
                                <th>日期</th>
                                <th>编号</th>
                                <th>类型</th>
                                <th>类别</th>
                                <th>会员／对象</th>
                                <th>月份／用途</th>
                                <th>金额</th>
                                <th>方式</th>
                                <th>状态</th>
                                <th>操作</th>
                            </tr>
                        </thead>

                        <tbody>
                            {% for r in rows %}
                                {% if r.record_type == 'expense' %}
                                    {% set display_no =
                                        r.payment_voucher_no
                                        or '-'
                                    %}
                                {% else %}
                                    {% set display_no =
                                        r.receipt_no
                                        or '-'
                                    %}
                                {% endif %}

                                <tr
                                    {% if r.status == 'cancelled' %}
                                        class="cancelled-row"
                                    {% endif %}
                                >
                                    <td style="white-space:nowrap;">
                                        {{ r.record_date }}
                                    </td>

                                    <td>
                                        <a
                                            class="record-link"
                                            href="{{ url_for(
                                                'finance.record_detail',
                                                record_id=r.id
                                            ) }}"
                                        >
                                            {{ display_no }}
                                        </a>

                                        <div style="margin-top:5px;">
                                            {% if r.payment_voucher_no %}
                                                <span class="number-badge number-pv">
                                                    PV
                                                </span>
                                            {% elif r.receipt_no and r.receipt_no.startswith('STW') %}
                                                <span class="number-badge number-stw">
                                                    STW
                                                </span>
                                            {% elif r.receipt_no and r.receipt_no.startswith('CHE') %}
                                                <span class="number-badge number-che">
                                                    CHE
                                                </span>
                                            {% endif %}
                                        </div>
                                    </td>

                                    <td>
                                        {% if r.record_type == 'expense' %}
                                            <span class="type-badge type-expense">
                                                支出
                                            </span>
                                        {% else %}
                                            <span class="type-badge type-income">
                                                收入
                                            </span>
                                        {% endif %}
                                    </td>

                                    <td>
                                        {{ r.category or '-' }}
                                    </td>

                                    <td>
                                        <div class="member-name">
                                            {{ r.name or '-' }}
                                        </div>

                                        {% if r.member_id %}
                                            <div class="member-id">
                                                {{ r.member_id }}
                                            </div>
                                        {% endif %}
                                    </td>

                                    <td class="month-purpose">
                                        {% if r.month_from or r.month_to %}
                                            {{ r.month_from or '-' }}
                                            至
                                            {{ r.month_to or '-' }}
                                        {% else %}
                                            {{ r.remarks or '-' }}
                                        {% endif %}
                                    </td>

                                    <td class="
                                        money
                                        {% if r.record_type == 'expense' %}
                                            money-expense
                                        {% else %}
                                            money-income
                                        {% endif %}
                                    ">
                                        {% if r.record_type == 'expense' %}
                                            −
                                        {% endif %}
                                        RM {{ "%.2f"|format(r.amount or 0) }}
                                    </td>

                                    <td>
                                        {{ r.payment_method or '-' }}
                                    </td>

                                    <td>
                                        {% if r.status == 'cancelled' %}
                                            <span class="status-badge status-cancelled">
                                                已作废
                                            </span>
                                        {% elif r.status == 'confirmed' or not r.status %}
                                            <span class="status-badge status-confirmed">
                                                已入账
                                            </span>
                                        {% else %}
                                            <span class="status-badge status-pending">
                                                待确认
                                            </span>
                                        {% endif %}
                                    </td>

                                    <td class="action-cell">
                                        <div class="row-actions">
                                            <a
                                                class="btn-tool btn-secondary mini-action"
                                                href="{{ url_for(
                                                    'finance.record_detail',
                                                    record_id=r.id
                                                ) }}"
                                            >
                                                查看
                                            </a>

                                            {% if r.status != 'cancelled' %}
                                                <form
                                                    class="cancel-form"
                                                    method="post"
                                                    action="{{ url_for(
                                                        'finance.cancel_record',
                                                        record_id=r.id
                                                    ) }}"
                                                    onsubmit="return confirm(
                                                        '确定要作废这笔记录吗？作废后不会计入财政统计。'
                                                    );"
                                                >
                                                    <input
                                                        class="form-input"
                                                        name="cancel_reason"
                                                        placeholder="作废原因"
                                                        required
                                                    >

                                                    <button
                                                        class="btn-tool btn-warning mini-action"
                                                        type="submit"
                                                    >
                                                        作废
                                                    </button>
                                                </form>
                                            {% endif %}
                                        </div>
                                    </td>
                                </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>

            {% else %}
                <div class="empty-state-custom">
                    <div style="font-size:48px;margin-bottom:10px;">
                        🔎
                    </div>

                    <h3>没有找到财政记录</h3>

                    <p>
                        请调整日期、类型、类别、编号或关键字。
                    </p>
                </div>
            {% endif %}
        </div>

        <div class="bottom-actions">
            <a
                class="btn-tool btn-secondary"
                href="{{ url_for('finance.finance_home') }}"
            >
                ← 返回财政首页
            </a>
        </div>

    </div>

    </body>
    </html>
    """,
        rows=rows,
        categories=categories,
        q=q,
        date_from=date_from,
        date_to=date_to,
        record_type=record_type,
        category=category,
        number_type=number_type,
        status_filter=status_filter,
        shown_count=shown_count,
        active_count=active_count,
        cancelled_count=cancelled_count,
        income_total=income_total,
        expense_total=expense_total,
        balance_total=balance_total
    )


@finance_bp.route("/records/<int:record_id>")
def record_detail(record_id):

    record = db_query("""
        select *
        from finance_records
        where id = %s
        limit 1
    """, (
        record_id,
    ), fetchone=True)

    if not record:
        return "找不到这笔财政记录", 404

    if record["record_type"] == "expense":
        display_no = (
            record.get("payment_voucher_no")
            or "-"
        )
        document_label = "Payment Voucher No."
        person_label = "付款对象"
    else:
        display_no = (
            record.get("receipt_no")
            or "-"
        )
        document_label = "Receipt No."
        person_label = "姓名"

    return render_template_string(FINANCE_DATE_COMPONENT + """
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">

        <meta
            name="viewport"
            content="width=device-width, initial-scale=1"
        >

        <title>财政记录详情</title>

        <link
            rel="stylesheet"
            href="{{ url_for(
                'static',
                filename='css/toolbox.css'
            ) }}"
        >

        <style>
            .detail-page{
                max-width:900px;
            }

            .detail-header{
                background:linear-gradient(
                    135deg,
                    #1d4ed8,
                    #2563eb
                );
                color:#fff;
                padding:28px;
                border-radius:22px;
                margin-bottom:20px;
                box-shadow:0 12px 30px rgba(37,99,235,.18);
            }

            .detail-document-label{
                font-size:14px;
                font-weight:800;
                opacity:.86;
                margin-bottom:6px;
                text-transform:uppercase;
                letter-spacing:.5px;
            }

            .detail-number{
                font-size:34px;
                font-weight:900;
                margin:0 0 10px;
                overflow-wrap:anywhere;
            }

            .detail-header-money{
                font-size:36px;
                font-weight:900;
                margin:4px 0 10px;
                line-height:1.2;
            }

            .detail-header-income{
                color:#bbf7d0;
            }

            .detail-header-expense{
                color:#fecaca;
            }

            .detail-subtitle{
                margin:0;
                opacity:.92;
                line-height:1.6;
            }

            .detail-grid{
                display:grid;
                grid-template-columns:repeat(2,minmax(0,1fr));
                gap:14px;
            }

            .detail-item{
                background:#f8fafc;
                border:1px solid #e2e8f0;
                border-radius:14px;
                padding:16px;
            }

            .detail-full{
                grid-column:1 / -1;
            }

            .detail-label{
                color:#64748b;
                font-size:14px;
                margin-bottom:6px;
            }

            .detail-value{
                color:#0f172a;
                font-size:19px;
                font-weight:800;
                line-height:1.5;
                overflow-wrap:anywhere;
            }

            .detail-money{
                color:#15803d;
                font-size:30px;
            }

            .detail-money.expense{
                color:#b91c1c;
            }

            .status-badge{
                display:inline-flex;
                align-items:center;
                justify-content:center;
                padding:8px 13px;
                border-radius:999px;
                font-size:15px;
                font-weight:900;
                white-space:nowrap;
            }

            .status-confirmed{
                background:#dcfce7;
                color:#166534;
            }

            .status-cancelled{
                background:#fee2e2;
                color:#991b1b;
            }

            .status-pending{
                background:#fef3c7;
                color:#92400e;
            }

            .cancel-box{
                background:#fef2f2;
                border-color:#fecaca;
            }

            .detail-actions{
                display:flex;
                flex-wrap:wrap;
                gap:10px;
                margin-top:20px;
            }

            @media print{

                body{
                    background:#fff;
                }

                .detail-actions{
                    display:none;
                }

                .detail-page{
                    max-width:none;
                    padding:0;
                }

                .detail-header,
                .card{
                    box-shadow:none;
                }
            }

            @media(max-width:650px){

                .detail-page{
                    padding-left:12px;
                    padding-right:12px;
                }

                .detail-header{
                    padding:22px 18px;
                    border-radius:18px;
                }

                .detail-number{
                    font-size:28px;
                }

                .detail-header-money{
                    font-size:30px;
                }

                .detail-grid{
                    grid-template-columns:1fr;
                }

                .detail-full{
                    grid-column:auto;
                }

                .detail-actions{
                    display:grid;
                }

                .detail-actions .btn-tool{
                    width:100%;
                }
            }
        </style>
    </head>

    <body>

    <div class="page detail-page">

        <div class="detail-header">

            <div class="detail-document-label">
                {{ document_label }}
            </div>

            <div class="detail-number">
                {{ display_no }}
            </div>

            <div class="
                detail-header-money
                {% if record.record_type == 'expense' %}
                    detail-header-expense
                {% else %}
                    detail-header-income
                {% endif %}
            ">
                {% if record.record_type == "expense" %}
                    −
                {% endif %}
                RM {{ "%.2f"|format(record.amount or 0) }}
            </div>

            <p class="detail-subtitle">
                {{ "支出记录" if record.record_type == "expense" else "收入记录" }}
                ·
                {{ record.category or "未分类" }}
            </p>

        </div>

        <div class="card">

            <h2 class="section-title">
                📄 完整记录
            </h2>

            <div class="detail-grid">

                <div class="detail-item">
                    <div class="detail-label">
                        记录类型
                    </div>

                    <div class="detail-value">
                        {{ "支出" if record.record_type == "expense" else "收入" }}
                    </div>
                </div>

                <div class="detail-item">
                    <div class="detail-label">
                        类别
                    </div>

                    <div class="detail-value">
                        {{ record.category or "-" }}
                    </div>
                </div>

                <div class="detail-item">
                    <div class="detail-label">
                        {{ "支出日期" if record.record_type == "expense" else "付款日期" }}
                    </div>

                    <div class="detail-value">
                        {{ record.record_date or "-" }}
                    </div>
                </div>

                {% if record.record_type != "expense" %}

                    <div class="detail-item">
                        <div class="detail-label">
                            开收条日期
                        </div>

                        <div class="detail-value">
                            {{ record.receipt_date or "-" }}
                        </div>
                    </div>

                {% endif %}

                <div class="detail-item">
                    <div class="detail-label">
                        {{ document_label }}
                    </div>

                    <div class="detail-value">
                        {{ display_no }}
                    </div>
                </div>

                {% if record.member_id %}

                    <div class="detail-item">
                        <div class="detail-label">
                            会员编号
                        </div>

                        <div class="detail-value">
                            {{ record.member_id }}
                        </div>
                    </div>

                {% endif %}

                <div class="detail-item">
                    <div class="detail-label">
                        {{ person_label }}
                    </div>

                    <div class="detail-value">
                        {{ record.name or "-" }}
                    </div>
                </div>

                {% if record.phone %}

                    <div class="detail-item">
                        <div class="detail-label">
                            电话号码
                        </div>

                        <div class="detail-value">
                            {{ record.phone }}
                        </div>
                    </div>

                {% endif %}

                <div class="detail-item">
                    <div class="detail-label">
                        付款方式
                    </div>

                    <div class="detail-value">
                        {{ record.payment_method or "-" }}
                    </div>
                </div>

                {% if record.bank_ref %}

                    <div class="detail-item">
                        <div class="detail-label">
                            银行 Reference
                        </div>

                        <div class="detail-value">
                            {{ record.bank_ref }}
                        </div>
                    </div>

                {% endif %}

                <div class="detail-item">
                    <div class="detail-label">
                        基金户口
                    </div>

                    <div class="detail-value">
                        {{ record.fund_account or "-" }}
                    </div>
                </div>

                {% if record.month_from or record.month_to %}

                    <div class="detail-item">
                        <div class="detail-label">
                            开始月份
                        </div>

                        <div class="detail-value">
                            {{ record.month_from or "-" }}
                        </div>
                    </div>

                    <div class="detail-item">
                        <div class="detail-label">
                            缴费至
                        </div>

                        <div class="detail-value">
                            {{ record.month_to or "-" }}
                        </div>
                    </div>

                {% endif %}

                <div class="detail-item detail-full">
                    <div class="detail-label">
                        金额
                    </div>

                    <div class="
                        detail-value
                        detail-money
                        {% if record.record_type == 'expense' %}
                            expense
                        {% endif %}
                    ">
                        {% if record.record_type == "expense" %}
                            −
                        {% endif %}
                        RM {{ "%.2f"|format(record.amount or 0) }}
                    </div>
                </div>

                <div class="detail-item detail-full">
                    <div class="detail-label">
                        备注／用途
                    </div>

                    <div class="detail-value">
                        {{ record.remarks or "-" }}
                    </div>
                </div>

                <div class="detail-item detail-full">
                    <div class="detail-label">
                        状态
                    </div>

                    <div class="detail-value">

                        {% if record.status == "cancelled" %}

                            <span class="status-badge status-cancelled">
                                ❌ 已作废
                            </span>

                        {% elif record.status == "confirmed"
                            or not record.status %}

                            <span class="status-badge status-confirmed">
                                ✅ 已入账
                            </span>

                        {% else %}

                            <span class="status-badge status-pending">
                                ⏳ {{ record.status }}
                            </span>

                        {% endif %}

                    </div>
                </div>

                {% if record.cancel_reason %}

                    <div class="
                        detail-item
                        detail-full
                        cancel-box
                    ">
                        <div class="detail-label">
                            作废原因
                        </div>

                        <div class="detail-value">
                            {{ record.cancel_reason }}
                        </div>
                    </div>

                {% endif %}

                {% if record.created_at %}

                    <div class="detail-item detail-full">
                        <div class="detail-label">
                            建立时间
                        </div>

                        <div class="detail-value">
                            {{ record.created_at }}
                        </div>
                    </div>

                {% endif %}

            </div>

        </div>

        <div class="detail-actions">

            <a
                class="btn-tool btn-secondary"
                href="{{ url_for('finance.records') }}"
            >
                ← 返回财政总账
            </a>

            <a
                class="btn-tool btn-primary"
                href="{{ url_for(
                    'finance.records',
                    q=display_no
                ) }}"
            >
                🔍 查询相同编号
            </a>

            <button
                class="btn-tool btn-success"
                type="button"
                onclick="window.print()"
            >
                🖨️ 打印
            </button>

        </div>

    </div>

    </body>
    </html>
    """,
        record=record,
        display_no=display_no,
        document_label=document_label,
        person_label=person_label
    )

@finance_bp.route("/dashboard")
def dashboard():

    ym = request.args.get(
        "ym",
        date.today().strftime("%Y-%m")
    )

    daily_income = db_query("""
        select
            category,
            coalesce(sum(amount), 0) as total
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
          and coalesce(status, 'confirmed') <> 'cancelled'
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
          and coalesce(status, 'confirmed') <> 'cancelled'
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
          and coalesce(status, 'confirmed') <> 'cancelled'
          and fund_account = '总会户口'
          and record_type = 'income'
        group by category
        order by category
    """, (ym,), fetchall=True)

    daily_income_total = db_query("""
        select coalesce(sum(amount), 0) as total
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
          and coalesce(status, 'confirmed') <> 'cancelled'
          and fund_account = '观音堂日常户口'
          and record_type = 'income'
    """, (ym,), fetchone=True)

    daily_expense_total = db_query("""
        select coalesce(sum(amount), 0) as total
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
          and coalesce(status, 'confirmed') <> 'cancelled'
          and fund_account = '观音堂日常户口'
          and record_type = 'expense'
    """, (ym,), fetchone=True)

    hq_income_total = db_query("""
        select coalesce(sum(amount), 0) as total
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
          and coalesce(status, 'confirmed') <> 'cancelled'
          and fund_account = '总会户口'
          and record_type = 'income'
    """, (ym,), fetchone=True)

    try:
        opening_balance = float(
            request.args.get("opening_balance") or 0
        )
    except (TypeError, ValueError):
        opening_balance = 0.0

    daily_income_value = float(
        daily_income_total["total"] or 0
    )

    daily_expense_value = float(
        daily_expense_total["total"] or 0
    )

    hq_income_value = float(
        hq_income_total["total"] or 0
    )

    daily_balance = (
        opening_balance
        + daily_income_value
        - daily_expense_value
    )

    monthly_net = (
        daily_income_value
        - daily_expense_value
    )

    return render_template_string(FINANCE_DATE_COMPONENT + """
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">

        <meta
            name="viewport"
            content="width=device-width, initial-scale=1"
        >

        <title>财政统计 Dashboard</title>

        <link
            rel="stylesheet"
            href="{{ url_for('static', filename='css/toolbox.css') }}"
        >

        <style>
            .dashboard-page {
                max-width: 1180px;
            }

            .dashboard-header {
                background:
                    linear-gradient(
                        135deg,
                        #2563eb,
                        #1d4ed8
                    );

                color: white;
                padding: 28px;
                border-radius: 22px;
                margin-bottom: 20px;
                box-shadow:
                    0 12px 30px
                    rgba(37, 99, 235, 0.18);
            }

            .dashboard-header h1 {
                margin: 0 0 8px;
                font-size: 30px;
            }

            .dashboard-header p {
                margin: 0;
                opacity: 0.92;
                line-height: 1.6;
            }

            .filter-grid {
                display: grid;
                grid-template-columns:
                    minmax(180px, 1fr)
                    minmax(220px, 1fr)
                    auto;
                gap: 14px;
                align-items: end;
            }

            .dashboard-summary {
                display: grid;
                grid-template-columns:
                    repeat(4, minmax(0, 1fr));
                gap: 16px;
                margin-bottom: 20px;
            }

            .dashboard-summary .summary-box {
                min-height: 130px;
                display: flex;
                flex-direction: column;
                justify-content: center;
                text-align: center;
            }

            .summary-icon {
                font-size: 30px;
                margin-bottom: 6px;
            }

            .summary-label {
                color: #64748b;
                font-size: 16px;
                margin-bottom: 7px;
            }

            .summary-value {
                color: #0f172a;
                font-size: 25px;
                font-weight: 800;
                white-space: nowrap;
            }

            .summary-positive {
                color: #15803d;
            }

            .summary-negative {
                color: #b91c1c;
            }

            .account-grid {
                display: grid;
                grid-template-columns:
                    minmax(0, 1fr)
                    minmax(0, 1fr);
                gap: 20px;
                align-items: start;
            }

            .account-title {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 12px;
                margin-bottom: 14px;
            }

            .account-title h2 {
                margin: 0;
                font-size: 23px;
                color: #1e293b;
            }

            .account-badge {
                display: inline-flex;
                align-items: center;
                padding: 7px 12px;
                border-radius: 999px;
                background: #dbeafe;
                color: #1d4ed8;
                font-size: 14px;
                font-weight: 700;
                white-space: nowrap;
            }

            .account-badge.hq {
                background: #ede9fe;
                color: #6d28d9;
            }

            .money {
                text-align: right;
                font-weight: 700;
                white-space: nowrap;
            }

            .money-income {
                color: #15803d;
            }

            .money-expense {
                color: #b91c1c;
            }

            .total-row {
                background: #f8fafc;
                font-weight: 800;
            }

            .balance-row {
                background: #eff6ff;
                font-weight: 800;
            }

            .dashboard-actions {
                display: flex;
                justify-content: space-between;
                gap: 12px;
                flex-wrap: wrap;
                margin-top: 20px;
            }

            .empty-row {
                text-align: center;
                padding: 28px 12px;
                color: #64748b;
            }

            .section-gap {
                margin-top: 20px;
            }

            @media (max-width: 900px) {

                .dashboard-summary {
                    grid-template-columns:
                        repeat(2, minmax(0, 1fr));
                }

                .account-grid {
                    grid-template-columns: 1fr;
                }
            }

            @media (max-width: 650px) {

                .dashboard-page {
                    padding-left: 12px;
                    padding-right: 12px;
                }

                .dashboard-header {
                    padding: 22px 18px;
                    border-radius: 18px;
                }

                .dashboard-header h1 {
                    font-size: 26px;
                }

                .filter-grid {
                    grid-template-columns: 1fr;
                }

                .filter-grid .btn-tool {
                    width: 100%;
                }

                .dashboard-summary {
                    grid-template-columns: 1fr;
                }

                .dashboard-summary .summary-box {
                    min-height: 110px;
                }

                .dashboard-actions {
                    flex-direction: column;
                }

                .dashboard-actions .btn-tool {
                    width: 100%;
                }

                .record-table {
                    min-width: 620px;
                }
            }
        </style>
    </head>

    <body>

    <div class="page dashboard-page">

        <div class="dashboard-header">

            <h1>📊 财政统计 Dashboard</h1>

            <p>
                查看每月收入、支出、户口结余及总会户口收入。
            </p>

        </div>

        <div class="card">

            <div class="section-title">
                🔎 查询条件
            </div>

            <form method="get">

                <div class="filter-grid">

                    <div class="form-group">

                        <label class="form-label">
                            月份
                        </label>

                        <input
                            class="form-input"
                            type="month"
                            name="ym"
                            value="{{ ym }}"
                            required
                        >

                    </div>

                    <div class="form-group">

                        <label class="form-label">
                            上月结余 RM
                        </label>

                        <input
                            class="form-input"
                            type="number"
                            name="opening_balance"
                            value="{{ opening_balance }}"
                            step="0.01"
                            min="0"
                            placeholder="例如 15000.00"
                        >

                    </div>

                    <button
                        class="btn-tool btn-primary"
                        type="submit"
                    >
                        🔍 查看统计
                    </button>

                </div>

            </form>

        </div>

        <div class="dashboard-summary">

            <div class="summary-box">

                <div class="summary-icon">
                    💰
                </div>

                <div class="summary-label">
                    日常户口收入
                </div>

                <div class="summary-value summary-positive">
                    RM {{ "%.2f"|format(daily_income_value) }}
                </div>

            </div>

            <div class="summary-box">

                <div class="summary-icon">
                    💸
                </div>

                <div class="summary-label">
                    日常户口支出
                </div>

                <div class="summary-value summary-negative">
                    RM {{ "%.2f"|format(daily_expense_value) }}
                </div>

            </div>

            <div class="summary-box">

                <div class="summary-icon">
                    📈
                </div>

                <div class="summary-label">
                    本月收支差额
                </div>

                <div class="
                    summary-value
                    {% if monthly_net >= 0 %}
                        summary-positive
                    {% else %}
                        summary-negative
                    {% endif %}
                ">
                    RM {{ "%.2f"|format(monthly_net) }}
                </div>

            </div>

            <div class="summary-box">

                <div class="summary-icon">
                    🏦
                </div>

                <div class="summary-label">
                    日常户口结余
                </div>

                <div class="
                    summary-value
                    {% if daily_balance >= 0 %}
                        summary-positive
                    {% else %}
                        summary-negative
                    {% endif %}
                ">
                    RM {{ "%.2f"|format(daily_balance) }}
                </div>

            </div>

        </div>

        <div class="account-grid">

            <div class="card">

                <div class="account-title">

                    <h2>
                        🏠 观音堂日常户口
                    </h2>

                    <span class="account-badge">
                        收入与支出
                    </span>

                </div>

                <div class="section-title">
                    📥 收入明细
                </div>

                <div class="table-responsive">

                    <table class="record-table">

                        <thead>
                            <tr>
                                <th>收入项目</th>
                                <th>金额 RM</th>
                            </tr>
                        </thead>

                        <tbody>

                            {% for r in daily_income %}

                            <tr>
                                <td>
                                    {{ r.category }}
                                </td>

                                <td class="money money-income">
                                    {{ "%.2f"|format(r.total) }}
                                </td>
                            </tr>

                            {% else %}

                            <tr>
                                <td
                                    colspan="2"
                                    class="empty-row"
                                >
                                    本月暂无收入记录
                                </td>
                            </tr>

                            {% endfor %}

                            <tr class="total-row">
                                <td>
                                    日常户口总收入
                                </td>

                                <td class="money money-income">
                                    {{ "%.2f"|format(daily_income_value) }}
                                </td>
                            </tr>

                        </tbody>

                    </table>

                </div>

                <div class="section-title section-gap">
                    📤 支出明细
                </div>

                <div class="table-responsive">

                    <table class="record-table">

                        <thead>
                            <tr>
                                <th>支出项目</th>
                                <th>金额 RM</th>
                            </tr>
                        </thead>

                        <tbody>

                            {% for r in daily_expense %}

                            <tr>
                                <td>
                                    {{ r.category }}
                                </td>

                                <td class="money money-expense">
                                    {{ "%.2f"|format(r.total) }}
                                </td>
                            </tr>

                            {% else %}

                            <tr>
                                <td
                                    colspan="2"
                                    class="empty-row"
                                >
                                    本月暂无支出记录
                                </td>
                            </tr>

                            {% endfor %}

                            <tr class="total-row">
                                <td>
                                    日常户口总支出
                                </td>

                                <td class="money money-expense">
                                    {{ "%.2f"|format(daily_expense_value) }}
                                </td>
                            </tr>

                            <tr class="balance-row">
                                <td>
                                    上月结余
                                </td>

                                <td class="money">
                                    {{ "%.2f"|format(opening_balance) }}
                                </td>
                            </tr>

                            <tr class="balance-row">
                                <td>
                                    本月结余
                                </td>

                                <td class="
                                    money
                                    {% if daily_balance >= 0 %}
                                        money-income
                                    {% else %}
                                        money-expense
                                    {% endif %}
                                ">
                                    {{ "%.2f"|format(daily_balance) }}
                                </td>
                            </tr>

                        </tbody>

                    </table>

                </div>

            </div>

            <div class="card">

                <div class="account-title">

                    <h2>
                        🏛️ 总会户口
                    </h2>

                    <span class="account-badge hq">
                        收入统计
                    </span>

                </div>

                <div class="section-title">
                    📥 收入明细
                </div>

                <div class="table-responsive">

                    <table class="record-table">

                        <thead>
                            <tr>
                                <th>收入项目</th>
                                <th>金额 RM</th>
                            </tr>
                        </thead>

                        <tbody>

                            {% for r in hq_income %}

                            <tr>
                                <td>
                                    {{ r.category }}
                                </td>

                                <td class="money money-income">
                                    {{ "%.2f"|format(r.total) }}
                                </td>
                            </tr>

                            {% else %}

                            <tr>
                                <td
                                    colspan="2"
                                    class="empty-row"
                                >
                                    本月暂无总会户口收入
                                </td>
                            </tr>

                            {% endfor %}

                            <tr class="total-row">
                                <td>
                                    总会户口总收入
                                </td>

                                <td class="money money-income">
                                    {{ "%.2f"|format(hq_income_value) }}
                                </td>
                            </tr>

                        </tbody>

                    </table>

                </div>

            </div>

        </div>

        <div class="dashboard-actions">

            <a
                class="btn-tool btn-secondary"
                href="{{ url_for('finance.finance_home') }}"
            >
                ← 返回财政首页
            </a>

            <a
                class="btn-tool btn-success"
                href="{{ url_for(
                    'finance.export_monthly_report',
                    ym=ym
                ) }}"
            >
                📥 下载专业版 Excel 月报
            </a>

        </div>

    </div>

    </body>
    </html>
    """,
        ym=ym,
        opening_balance=opening_balance,
        daily_income=daily_income,
        daily_expense=daily_expense,
        hq_income=hq_income,
        daily_income_value=daily_income_value,
        daily_expense_value=daily_expense_value,
        hq_income_value=hq_income_value,
        daily_balance=daily_balance,
        monthly_net=monthly_net
    )


@finance_bp.route("/export_monthly_report")
def export_monthly_report():

    ym = request.args.get(
        "ym",
        date.today().strftime("%Y-%m")
    )

    rows = db_query("""
        select
            record_date,
            receipt_date,
            receipt_no,
            payment_voucher_no,
            category,
            record_type,
            fund_account,
            member_id,
            name,
            phone,
            amount,
            payment_method,
            bank_ref,
            month_from,
            month_to,
            remarks as note
        from finance_records
        where to_char(record_date, 'YYYY-MM') = %s
          and coalesce(status, 'confirmed') <> 'cancelled'
        order by
            record_date,
            case
                when record_type = 'expense'
                then payment_voucher_no
                else receipt_no
            end,
            id
    """, (
        ym,
    ), fetchall=True)

    wb = Workbook()

    ws = wb.active
    ws.title = "财政摘要"

    ws_income = wb.create_sheet(
        "收入明细"
    )

    ws_expense = wb.create_sheet(
        "支出明细"
    )

    # =========================
    # 样式
    # =========================
    green = "1F5E20"
    light_green = "EAF4E4"

    blue = "0B3A78"
    light_blue = "EAF2FF"

    gold = "F8E7B8"
    light_gold = "FFF8E8"

    dark_text = "1F2937"
    red = "D00000"
    gray = "64748B"

    thin_green = Side(
        style="thin",
        color="8DBA7C"
    )

    thin_blue = Side(
        style="thin",
        color="7FA6D9"
    )

    thin_gold = Side(
        style="thin",
        color="C9962C"
    )

    thin_gray = Side(
        style="thin",
        color="CCCCCC"
    )

    border_green = Border(
        left=thin_green,
        right=thin_green,
        top=thin_green,
        bottom=thin_green
    )

    border_blue = Border(
        left=thin_blue,
        right=thin_blue,
        top=thin_blue,
        bottom=thin_blue
    )

    border_gold = Border(
        left=thin_gold,
        right=thin_gold,
        top=thin_gold,
        bottom=thin_gold
    )

    border_gray = Border(
        left=thin_gray,
        right=thin_gray,
        top=thin_gray,
        bottom=thin_gray
    )

    money_fmt = '"RM"#,##0.00'

    def money(value):
        return float(value or 0)

    def set_range_border(
        sheet,
        cell_range,
        border
    ):
        for row in sheet[cell_range]:
            for cell in row:
                cell.border = border

    def fill_range(
        sheet,
        cell_range,
        fill
    ):
        for row in sheet[cell_range]:
            for cell in row:
                cell.fill = fill

    def auto_width(
        sheet,
        max_width=32
    ):
        for column_cells in sheet.columns:

            max_length = 0

            column_letter = get_column_letter(
                column_cells[0].column
            )

            for cell in column_cells:

                if cell.value is None:
                    continue

                cell_text = str(cell.value)

                for line in cell_text.splitlines():
                    max_length = max(
                        max_length,
                        len(line)
                    )

            sheet.column_dimensions[
                column_letter
            ].width = min(
                max_length + 4,
                max_width
            )

    def set_print_layout(
        sheet,
        orientation="landscape"
    ):
        sheet.page_setup.orientation = orientation
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0

        sheet.sheet_properties.pageSetUpPr.fitToPage = True

        sheet.page_margins.left = 0.25
        sheet.page_margins.right = 0.25
        sheet.page_margins.top = 0.4
        sheet.page_margins.bottom = 0.4
        sheet.page_margins.header = 0.2
        sheet.page_margins.footer = 0.2

    # =========================
    # 数据整理
    # =========================
    summary = {
        "观音堂日常户口": {
            "income": 0,
            "expense": 0
        },
        "总会户口": {
            "income": 0,
            "expense": 0
        },
    }

    income_rows = []
    expense_rows = []

    income_by_category = {}
    expense_by_category = {}

    for row in rows:

        fund_account = (
            row["fund_account"]
            or "未分类"
        )

        record_type = (
            row["record_type"]
            or "income"
        )

        amount = money(
            row["amount"]
        )

        if fund_account not in summary:

            summary[fund_account] = {
                "income": 0,
                "expense": 0
            }

        summary[fund_account][
            record_type
        ] += amount

        category_name = (
            row["category"]
            or "未分类"
        )

        if record_type == "income":

            income_rows.append(row)

            income_by_category[
                category_name
            ] = (
                income_by_category.get(
                    category_name,
                    0
                )
                + amount
            )

        else:

            expense_rows.append(row)

            expense_by_category[
                category_name
            ] = (
                expense_by_category.get(
                    category_name,
                    0
                )
                + amount
            )

    daily_income = summary.get(
        "观音堂日常户口",
        {}
    ).get(
        "income",
        0
    )

    daily_expense = summary.get(
        "观音堂日常户口",
        {}
    ).get(
        "expense",
        0
    )

    daily_balance = (
        daily_income
        - daily_expense
    )

    hq_income = summary.get(
        "总会户口",
        {}
    ).get(
        "income",
        0
    )

    hq_expense = summary.get(
        "总会户口",
        {}
    ).get(
        "expense",
        0
    )

    hq_balance = (
        hq_income
        - hq_expense
    )

    total_income = sum(
        item["income"]
        for item in summary.values()
    )

    total_expense = sum(
        item["expense"]
        for item in summary.values()
    )

    total_balance = (
        total_income
        - total_expense
    )

    # =========================
    # Sheet 1：财政摘要
    # =========================
    ws.sheet_view.showGridLines = False

    for column in range(1, 16):
        ws.column_dimensions[
            get_column_letter(column)
        ].width = 14

    for row_no in range(1, 40):
        ws.row_dimensions[row_no].height = 24

    ws.row_dimensions[1].height = 14
    ws.row_dimensions[2].height = 34
    ws.row_dimensions[3].height = 26
    ws.row_dimensions[4].height = 12

    ws.merge_cells("A1:O4")

    ws["A1"] = (
        "观音堂财政月报\n"
        f"{ym}"
    )

    ws["A1"].font = Font(
        bold=True,
        size=24,
        color=dark_text
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    ws["A1"].fill = PatternFill(
        "solid",
        fgColor=light_gold
    )

    set_range_border(
        ws,
        "A1:O4",
        border_gold
    )

    # 观音堂日常户口
    ws.merge_cells("A5:G5")

    ws["A5"] = "观音堂日常户口"

    ws["A5"].fill = PatternFill(
        "solid",
        fgColor=green
    )

    ws["A5"].font = Font(
        bold=True,
        size=14,
        color="FFFFFF"
    )

    ws["A5"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    daily_table = [
        [
            "项目",
            "收入（RM）",
            "支出（RM）",
            "本月结余（RM）"
        ],
        [
            "收入",
            daily_income,
            "-",
            daily_income
        ],
        [
            "支出",
            "-",
            daily_expense,
            -daily_expense
        ],
        [
            "本月结余",
            "-",
            "-",
            daily_balance
        ],
    ]

    start_row = 6

    for row_no, row_data in enumerate(
        daily_table,
        start_row
    ):

        ws.merge_cells(
            start_row=row_no,
            start_column=1,
            end_row=row_no,
            end_column=2
        )

        ws.merge_cells(
            start_row=row_no,
            start_column=3,
            end_row=row_no,
            end_column=4
        )

        ws.merge_cells(
            start_row=row_no,
            start_column=5,
            end_row=row_no,
            end_column=5
        )

        ws.merge_cells(
            start_row=row_no,
            start_column=6,
            end_row=row_no,
            end_column=7
        )

        ws.cell(
            row_no,
            1
        ).value = row_data[0]

        ws.cell(
            row_no,
            3
        ).value = row_data[1]

        ws.cell(
            row_no,
            5
        ).value = row_data[2]

        ws.cell(
            row_no,
            6
        ).value = row_data[3]

        for column in [
            1,
            3,
            5,
            6
        ]:

            cell = ws.cell(
                row_no,
                column
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            cell.border = border_green

        if row_no == 6:

            fill_range(
                ws,
                f"A{row_no}:G{row_no}",
                PatternFill(
                    "solid",
                    fgColor=light_green
                )
            )

            for column in [
                1,
                3,
                5,
                6
            ]:

                ws.cell(
                    row_no,
                    column
                ).font = Font(
                    bold=True
                )

        elif row_no == 9:

            fill_range(
                ws,
                f"A{row_no}:G{row_no}",
                PatternFill(
                    "solid",
                    fgColor="DDEED2"
                )
            )

            ws.cell(
                row_no,
                6
            ).font = Font(
                bold=True,
                size=14,
                color=green
            )

    for cell_ref in [
        "C7",
        "E8",
        "F7",
        "F8",
        "F9"
    ]:
        ws[cell_ref].number_format = money_fmt

    # 总会户口
    ws.merge_cells("I5:O5")

    ws["I5"] = "总会户口"

    ws["I5"].fill = PatternFill(
        "solid",
        fgColor=blue
    )

    ws["I5"].font = Font(
        bold=True,
        size=14,
        color="FFFFFF"
    )

    ws["I5"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    hq_table = [
        [
            "项目",
            "收入（RM）",
            "支出（RM）",
            "本月结余（RM）"
        ],
        [
            "收入",
            hq_income,
            "-",
            hq_income
        ],
        [
            "支出",
            "-",
            hq_expense,
            -hq_expense
        ],
        [
            "本月结余",
            "-",
            "-",
            hq_balance
        ],
    ]

    for row_no, row_data in enumerate(
        hq_table,
        start_row
    ):

        ws.merge_cells(
            start_row=row_no,
            start_column=9,
            end_row=row_no,
            end_column=10
        )

        ws.merge_cells(
            start_row=row_no,
            start_column=11,
            end_row=row_no,
            end_column=12
        )

        ws.merge_cells(
            start_row=row_no,
            start_column=13,
            end_row=row_no,
            end_column=13
        )

        ws.merge_cells(
            start_row=row_no,
            start_column=14,
            end_row=row_no,
            end_column=15
        )

        ws.cell(
            row_no,
            9
        ).value = row_data[0]

        ws.cell(
            row_no,
            11
        ).value = row_data[1]

        ws.cell(
            row_no,
            13
        ).value = row_data[2]

        ws.cell(
            row_no,
            14
        ).value = row_data[3]

        for column in [
            9,
            11,
            13,
            14
        ]:

            cell = ws.cell(
                row_no,
                column
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            cell.border = border_blue

        if row_no == 6:

            fill_range(
                ws,
                f"I{row_no}:O{row_no}",
                PatternFill(
                    "solid",
                    fgColor=light_blue
                )
            )

            for column in [
                9,
                11,
                13,
                14
            ]:

                ws.cell(
                    row_no,
                    column
                ).font = Font(
                    bold=True,
                    color=blue
                )

        elif row_no == 9:

            fill_range(
                ws,
                f"I{row_no}:O{row_no}",
                PatternFill(
                    "solid",
                    fgColor="DCEAFF"
                )
            )

            ws.cell(
                row_no,
                14
            ).font = Font(
                bold=True,
                size=14,
                color=blue
            )

    for cell_ref in [
        "K7",
        "M8",
        "N7",
        "N8",
        "N9"
    ]:
        ws[cell_ref].number_format = money_fmt

    # 总计横条
    ws.merge_cells("A11:C12")

    ws["A11"] = "总计"

    ws["A11"].font = Font(
        bold=True,
        size=16,
        color="7A5200"
    )

    ws["A11"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws["A11"].fill = PatternFill(
        "solid",
        fgColor="FFF2CC"
    )

    ws.merge_cells("D11:F12")

    ws["D11"] = (
        "总收入（RM）\n"
        f"{total_income:,.2f}"
    )

    ws["D11"].font = Font(
        bold=True,
        size=12,
        color=green
    )

    ws["D11"].alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    ws["D11"].fill = PatternFill(
        "solid",
        fgColor=light_gold
    )

    ws.merge_cells("G11:I12")

    ws["G11"] = (
        "总支出（RM）\n"
        f"{total_expense:,.2f}"
    )

    ws["G11"].font = Font(
        bold=True,
        size=12,
        color=red
    )

    ws["G11"].alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    ws["G11"].fill = PatternFill(
        "solid",
        fgColor=light_gold
    )

    ws.merge_cells("J11:O12")

    ws["J11"] = (
        "总体结余（RM）\n"
        f"{total_balance:,.2f}"
    )

    ws["J11"].font = Font(
        bold=True,
        size=12,
        color=blue
    )

    ws["J11"].alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    ws["J11"].fill = PatternFill(
        "solid",
        fgColor=light_gold
    )

    set_range_border(
        ws,
        "A11:O12",
        border_gold
    )

    # 收入分类汇总
    ws.merge_cells("A14:G14")

    ws["A14"] = "收入分类汇总"

    ws["A14"].fill = PatternFill(
        "solid",
        fgColor=green
    )

    ws["A14"].font = Font(
        bold=True,
        size=14,
        color="FFFFFF"
    )

    ws["A14"].alignment = Alignment(
        horizontal="center"
    )

    ws.merge_cells("A15:E15")
    ws.merge_cells("F15:G15")

    ws["A15"] = "收入项目"
    ws["F15"] = "金额（RM）"

    for cell_ref in [
        "A15",
        "F15"
    ]:

        ws[cell_ref].fill = PatternFill(
            "solid",
            fgColor=light_green
        )

        ws[cell_ref].font = Font(
            bold=True
        )

        ws[cell_ref].alignment = Alignment(
            horizontal="center"
        )

    income_summary_start = 16
    income_summary_row = income_summary_start

    for category_name, amount in sorted(
        income_by_category.items(),
        key=lambda item: item[1],
        reverse=True
    ):

        ws.merge_cells(
            start_row=income_summary_row,
            start_column=1,
            end_row=income_summary_row,
            end_column=5
        )

        ws.merge_cells(
            start_row=income_summary_row,
            start_column=6,
            end_row=income_summary_row,
            end_column=7
        )

        ws.cell(
            income_summary_row,
            1
        ).value = category_name

        ws.cell(
            income_summary_row,
            6
        ).value = amount

        ws.cell(
            income_summary_row,
            6
        ).number_format = money_fmt

        ws.cell(
            income_summary_row,
            1
        ).alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        ws.cell(
            income_summary_row,
            6
        ).alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        income_summary_row += 1

    if not income_by_category:

        ws.merge_cells("A16:G16")

        ws["A16"] = "本月没有收入记录"

        ws["A16"].alignment = Alignment(
            horizontal="center"
        )

        income_summary_row = 17

    income_total_row = income_summary_row

    ws.merge_cells(
        start_row=income_total_row,
        start_column=1,
        end_row=income_total_row,
        end_column=5
    )

    ws.merge_cells(
        start_row=income_total_row,
        start_column=6,
        end_row=income_total_row,
        end_column=7
    )

    ws.cell(
        income_total_row,
        1
    ).value = "收入总计"

    ws.cell(
        income_total_row,
        6
    ).value = total_income

    ws.cell(
        income_total_row,
        6
    ).number_format = money_fmt

    fill_range(
        ws,
        f"A{income_total_row}:G{income_total_row}",
        PatternFill(
            "solid",
            fgColor="DDEED2"
        )
    )

    ws.cell(
        income_total_row,
        1
    ).font = Font(
        bold=True,
        color=green
    )

    ws.cell(
        income_total_row,
        6
    ).font = Font(
        bold=True,
        color=green
    )

    ws.cell(
        income_total_row,
        1
    ).alignment = Alignment(
        horizontal="right"
    )

    ws.cell(
        income_total_row,
        6
    ).alignment = Alignment(
        horizontal="center"
    )

    set_range_border(
        ws,
        f"A15:G{income_total_row}",
        border_green
    )

    # 支出分类汇总
    ws.merge_cells("I14:O14")

    ws["I14"] = "支出分类汇总"

    ws["I14"].fill = PatternFill(
        "solid",
        fgColor=blue
    )

    ws["I14"].font = Font(
        bold=True,
        size=14,
        color="FFFFFF"
    )

    ws["I14"].alignment = Alignment(
        horizontal="center"
    )

    ws.merge_cells("I15:M15")
    ws.merge_cells("N15:O15")

    ws["I15"] = "支出项目"
    ws["N15"] = "金额（RM）"

    for cell_ref in [
        "I15",
        "N15"
    ]:

        ws[cell_ref].fill = PatternFill(
            "solid",
            fgColor=light_blue
        )

        ws[cell_ref].font = Font(
            bold=True,
            color=blue
        )

        ws[cell_ref].alignment = Alignment(
            horizontal="center"
        )

    expense_summary_start = 16
    expense_summary_row = expense_summary_start

    for category_name, amount in sorted(
        expense_by_category.items(),
        key=lambda item: item[1],
        reverse=True
    ):

        ws.merge_cells(
            start_row=expense_summary_row,
            start_column=9,
            end_row=expense_summary_row,
            end_column=13
        )

        ws.merge_cells(
            start_row=expense_summary_row,
            start_column=14,
            end_row=expense_summary_row,
            end_column=15
        )

        ws.cell(
            expense_summary_row,
            9
        ).value = category_name

        ws.cell(
            expense_summary_row,
            14
        ).value = amount

        ws.cell(
            expense_summary_row,
            14
        ).number_format = money_fmt

        ws.cell(
            expense_summary_row,
            9
        ).alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        ws.cell(
            expense_summary_row,
            14
        ).alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        expense_summary_row += 1

    if not expense_by_category:

        ws.merge_cells("I16:O16")

        ws["I16"] = "本月没有支出记录"

        ws["I16"].alignment = Alignment(
            horizontal="center"
        )

        expense_summary_row = 17

    expense_total_row = expense_summary_row

    ws.merge_cells(
        start_row=expense_total_row,
        start_column=9,
        end_row=expense_total_row,
        end_column=13
    )

    ws.merge_cells(
        start_row=expense_total_row,
        start_column=14,
        end_row=expense_total_row,
        end_column=15
    )

    ws.cell(
        expense_total_row,
        9
    ).value = "支出总计"

    ws.cell(
        expense_total_row,
        14
    ).value = total_expense

    ws.cell(
        expense_total_row,
        14
    ).number_format = money_fmt

    fill_range(
        ws,
        f"I{expense_total_row}:O{expense_total_row}",
        PatternFill(
            "solid",
            fgColor="DCEAFF"
        )
    )

    ws.cell(
        expense_total_row,
        9
    ).font = Font(
        bold=True,
        color=blue
    )

    ws.cell(
        expense_total_row,
        14
    ).font = Font(
        bold=True,
        color=red
    )

    ws.cell(
        expense_total_row,
        9
    ).alignment = Alignment(
        horizontal="right"
    )

    ws.cell(
        expense_total_row,
        14
    ).alignment = Alignment(
        horizontal="center"
    )

    set_range_border(
        ws,
        f"I15:O{expense_total_row}",
        border_blue
    )

    max_summary_row = max(
        income_total_row,
        expense_total_row
    )

    ws.freeze_panes = "A14"

    ws.print_area = (
        f"A1:O{max_summary_row}"
    )

    set_print_layout(
        ws,
        orientation="landscape"
    )

    # =========================
    # Sheet 2：收入明细
    # =========================
    def write_income_sheet(
        sheet,
        data_rows
    ):

        sheet.sheet_view.showGridLines = False

        sheet.merge_cells("A1:N2")

        sheet["A1"] = (
            f"收入明细 - {ym}"
        )

        sheet["A1"].font = Font(
            bold=True,
            size=18,
            color="FFFFFF"
        )

        sheet["A1"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        sheet["A1"].fill = PatternFill(
            "solid",
            fgColor=green
        )

        headers = [
            "付款日期",
            "开收条日期",
            "Receipt No.",
            "类别",
            "基金户口",
            "会员编号",
            "姓名",
            "电话",
            "金额（RM）",
            "付款方式",
            "银行 Reference",
            "开始月份",
            "缴费至",
            "备注",
        ]

        for column, header in enumerate(
            headers,
            1
        ):

            cell = sheet.cell(
                4,
                column
            )

            cell.value = header

            cell.fill = PatternFill(
                "solid",
                fgColor=light_green
            )

            cell.font = Font(
                bold=True,
                color=dark_text
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            cell.border = border_gray

        row_no = 5

        for row in data_rows:

            values = [
                row["record_date"],
                row["receipt_date"],
                row["receipt_no"],
                row["category"],
                row["fund_account"],
                row["member_id"],
                row["name"],
                row["phone"],
                money(row["amount"]),
                row["payment_method"],
                row["bank_ref"],
                row["month_from"],
                row["month_to"],
                row["note"],
            ]

            for column, value in enumerate(
                values,
                1
            ):

                cell = sheet.cell(
                    row_no,
                    column
                )

                cell.value = value

                cell.border = border_gray

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=(
                        column in [
                            7,
                            14
                        ]
                    )
                )

                if column == 9:
                    cell.number_format = money_fmt

            row_no += 1

        if not data_rows:

            sheet.merge_cells(
                start_row=5,
                start_column=1,
                end_row=5,
                end_column=14
            )

            sheet.cell(
                5,
                1
            ).value = "本月没有收入记录"

            sheet.cell(
                5,
                1
            ).alignment = Alignment(
                horizontal="center"
            )

            row_no = 6

        total_row = row_no + 1

        sheet.merge_cells(
            start_row=total_row,
            start_column=1,
            end_row=total_row,
            end_column=8
        )

        sheet.cell(
            total_row,
            1
        ).value = "收入总计"

        sheet.cell(
            total_row,
            1
        ).font = Font(
            bold=True
        )

        sheet.cell(
            total_row,
            1
        ).alignment = Alignment(
            horizontal="right"
        )

        sheet.cell(
            total_row,
            9
        ).value = sum(
            money(row["amount"])
            for row in data_rows
        )

        sheet.cell(
            total_row,
            9
        ).number_format = money_fmt

        sheet.cell(
            total_row,
            9
        ).font = Font(
            bold=True,
            color=green
        )

        fill_range(
            sheet,
            f"A{total_row}:N{total_row}",
            PatternFill(
                "solid",
                fgColor=light_green
            )
        )

        set_range_border(
            sheet,
            f"A{total_row}:N{total_row}",
            border_gray
        )

        sheet.freeze_panes = "A5"
        sheet.auto_filter.ref = (
            f"A4:N{max(4, row_no - 1)}"
        )

        sheet.print_title_rows = "1:4"
        sheet.print_area = (
            f"A1:N{total_row}"
        )

        auto_width(
            sheet,
            max_width=30
        )

        sheet.column_dimensions["N"].width = 34
        sheet.column_dimensions["G"].width = 18

        set_print_layout(
            sheet,
            orientation="landscape"
        )

    # =========================
    # Sheet 3：支出明细
    # =========================
    def write_expense_sheet(
        sheet,
        data_rows
    ):

        sheet.sheet_view.showGridLines = False

        sheet.merge_cells("A1:I2")

        sheet["A1"] = (
            f"支出明细 - {ym}"
        )

        sheet["A1"].font = Font(
            bold=True,
            size=18,
            color="FFFFFF"
        )

        sheet["A1"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        sheet["A1"].fill = PatternFill(
            "solid",
            fgColor=blue
        )

        headers = [
            "支出日期",
            "Payment Voucher No.",
            "类别",
            "基金户口",
            "付款对象",
            "金额（RM）",
            "付款方式",
            "银行 Reference",
            "用途／备注",
        ]

        for column, header in enumerate(
            headers,
            1
        ):

            cell = sheet.cell(
                4,
                column
            )

            cell.value = header

            cell.fill = PatternFill(
                "solid",
                fgColor=light_blue
            )

            cell.font = Font(
                bold=True,
                color=dark_text
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            cell.border = border_gray

        row_no = 5

        for row in data_rows:

            values = [
                row["record_date"],
                row["payment_voucher_no"],
                row["category"],
                row["fund_account"],
                row["name"],
                money(row["amount"]),
                row["payment_method"],
                row["bank_ref"],
                row["note"],
            ]

            for column, value in enumerate(
                values,
                1
            ):

                cell = sheet.cell(
                    row_no,
                    column
                )

                cell.value = value

                cell.border = border_gray

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=(
                        column in [
                            5,
                            9
                        ]
                    )
                )

                if column == 6:
                    cell.number_format = money_fmt

            row_no += 1

        if not data_rows:

            sheet.merge_cells(
                start_row=5,
                start_column=1,
                end_row=5,
                end_column=9
            )

            sheet.cell(
                5,
                1
            ).value = "本月没有支出记录"

            sheet.cell(
                5,
                1
            ).alignment = Alignment(
                horizontal="center"
            )

            row_no = 6

        total_row = row_no + 1

        sheet.merge_cells(
            start_row=total_row,
            start_column=1,
            end_row=total_row,
            end_column=5
        )

        sheet.cell(
            total_row,
            1
        ).value = "支出总计"

        sheet.cell(
            total_row,
            1
        ).font = Font(
            bold=True
        )

        sheet.cell(
            total_row,
            1
        ).alignment = Alignment(
            horizontal="right"
        )

        sheet.cell(
            total_row,
            6
        ).value = sum(
            money(row["amount"])
            for row in data_rows
        )

        sheet.cell(
            total_row,
            6
        ).number_format = money_fmt

        sheet.cell(
            total_row,
            6
        ).font = Font(
            bold=True,
            color=red
        )

        fill_range(
            sheet,
            f"A{total_row}:I{total_row}",
            PatternFill(
                "solid",
                fgColor=light_blue
            )
        )

        set_range_border(
            sheet,
            f"A{total_row}:I{total_row}",
            border_gray
        )

        sheet.freeze_panes = "A5"
        sheet.auto_filter.ref = (
            f"A4:I{max(4, row_no - 1)}"
        )

        sheet.print_title_rows = "1:4"
        sheet.print_area = (
            f"A1:I{total_row}"
        )

        auto_width(
            sheet,
            max_width=32
        )

        sheet.column_dimensions["I"].width = 36
        sheet.column_dimensions["E"].width = 22

        set_print_layout(
            sheet,
            orientation="landscape"
        )

    write_income_sheet(
        ws_income,
        income_rows
    )

    write_expense_sheet(
        ws_expense,
        expense_rows
    )

    # =========================
    # 输出
    # =========================
    output = BytesIO()

    wb.save(output)

    output.seek(0)

    filename = (
        f"观音堂财政月报_{ym}.xlsx"
    )

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        )
    )

def get_next_payment_voucher():

    row = db_query("""
        select payment_voucher_no
        from finance_records
        where payment_voucher_no is not null
          and payment_voucher_no <> ''
        order by payment_voucher_no desc
        limit 1
    """, fetchone=True)

    if not row:
        return "PV000001"

    m = re.search(r"(\d+)$", row["payment_voucher_no"])

    if not m:
        return "PV000001"

    return f"PV{int(m.group(1))+1:06d}"

@finance_bp.route(
    "/vendors",
    methods=["GET", "POST"]
)
def finance_vendors():

    message = ""

    if request.method == "POST":

        vendor_name = request.form.get(
            "vendor_name",
            ""
        ).strip()

        sort_order_raw = request.form.get(
            "sort_order",
            "0"
        ).strip()

        try:
            sort_order = int(
                sort_order_raw or 0
            )
        except ValueError:
            sort_order = 0

        if not vendor_name:

            message = "请填写付款对象名称。"

        else:

            existing = db_query("""
                select id
                from finance_vendors
                where lower(vendor_name) = lower(%s)
                limit 1
            """, (
                vendor_name,
            ), fetchone=True)

            if existing:

                message = "这个付款对象已经存在。"

            else:

                db_query("""
                    insert into finance_vendors
                    (
                        vendor_name,
                        is_active,
                        sort_order
                    )
                    values
                    (
                        %s,
                        true,
                        %s
                    )
                """, (
                    vendor_name,
                    sort_order
                ))

                return redirect(
                    url_for(
                        "finance.finance_vendors"
                    )
                )

    rows = db_query("""
        select
            id,
            vendor_name,
            is_active,
            sort_order,
            created_at
        from finance_vendors
        order by
            is_active desc,
            sort_order,
            vendor_name
    """, fetchall=True)

    return render_template_string("""
    <!doctype html>
    <html lang="zh">
    <head>

        <meta charset="utf-8">

        <meta
            name="viewport"
            content="width=device-width, initial-scale=1"
        >

        <title>付款对象管理</title>

        <link
            rel="stylesheet"
            href="{{ url_for(
                'static',
                filename='css/toolbox.css'
            ) }}"
        >

        <style>

            .vendor-page{
                max-width:980px;
            }

            .vendor-header{
                background:linear-gradient(
                    135deg,
                    #7c3aed,
                    #5b21b6
                );
                color:#fff;
                padding:28px;
                border-radius:22px;
                margin-bottom:20px;
            }

            .vendor-header h1{
                margin:0 0 8px;
            }

            .vendor-header p{
                margin:0;
                opacity:.92;
            }

            .vendor-form-grid{
                display:grid;
                grid-template-columns:1fr 180px auto;
                gap:12px;
                align-items:end;
            }

            .vendor-status{
                display:inline-flex;
                align-items:center;
                justify-content:center;
                padding:7px 11px;
                border-radius:999px;
                font-weight:800;
                font-size:14px;
            }

            .vendor-active{
                background:#dcfce7;
                color:#166534;
            }

            .vendor-inactive{
                background:#fee2e2;
                color:#991b1b;
            }

            @media(max-width:700px){

                .vendor-form-grid{
                    grid-template-columns:1fr;
                }

                .vendor-form-grid .btn-tool{
                    width:100%;
                }
            }

        </style>

    </head>

    <body>

    <div class="page vendor-page">

        <div class="vendor-header">

            <h1>🏢 付款对象管理</h1>

            <p>
                统一维护常用 Vendor，避免同一个对象出现不同写法。
            </p>

        </div>

        {% if message %}

            <div class="alert alert-danger">
                ⚠️ {{ message }}
            </div>

        {% endif %}

        <div class="card">

            <div class="section-title">
                ➕ 新增付款对象
            </div>

            <form method="post">

                <div class="vendor-form-grid">

                    <div class="form-group">

                        <label class="form-label">
                            付款对象名称
                        </label>

                        <input
                            class="form-input"
                            name="vendor_name"
                            placeholder="例如 TM Unifi"
                            required
                        >

                    </div>

                    <div class="form-group">

                        <label class="form-label">
                            排序
                        </label>

                        <input
                            class="form-input"
                            name="sort_order"
                            type="number"
                            value="0"
                        >

                    </div>

                    <button
                        class="btn-tool btn-primary"
                        type="submit"
                    >
                        保存
                    </button>

                </div>

            </form>

        </div>

        <div class="card">

            <div class="section-title">
                📋 付款对象名单
            </div>

            <div class="table-responsive">

                <table class="record-table">

                    <thead>
                        <tr>
                            <th>付款对象</th>
                            <th>排序</th>
                            <th>状态</th>
                            <th>操作</th>
                        </tr>
                    </thead>

                    <tbody>

                        {% for r in rows %}

                            <tr>

                                <td>
                                    <strong>
                                        {{ r.vendor_name }}
                                    </strong>
                                </td>

                                <td>
                                    {{ r.sort_order }}
                                </td>

                                <td>

                                    {% if r.is_active %}

                                        <span class="
                                            vendor-status
                                            vendor-active
                                        ">
                                            启用
                                        </span>

                                    {% else %}

                                        <span class="
                                            vendor-status
                                            vendor-inactive
                                        ">
                                            停用
                                        </span>

                                    {% endif %}

                                </td>

                                <td>

                                    <form
                                        method="post"
                                        action="{{ url_for(
                                            'finance.toggle_finance_vendor',
                                            vendor_id=r.id
                                        ) }}"
                                    >

                                        <button
                                            class="
                                                btn-tool
                                                btn-secondary
                                            "
                                            type="submit"
                                        >
                                            {% if r.is_active %}
                                                停用
                                            {% else %}
                                                恢复
                                            {% endif %}
                                        </button>

                                    </form>

                                </td>

                            </tr>

                        {% else %}

                            <tr>
                                <td
                                    colspan="4"
                                    style="text-align:center;"
                                >
                                    还没有付款对象
                                </td>
                            </tr>

                        {% endfor %}

                    </tbody>

                </table>

            </div>

        </div>

        <div class="btn-row">

            <a
                class="btn-tool btn-secondary"
                href="{{ url_for(
                    'finance.finance_expense_menu'
                ) }}"
            >
                ← 返回支出项目
            </a>

        </div>

    </div>

    </body>
    </html>
    """,
        rows=rows,
        message=message
    )


@finance_bp.route(
    "/vendors/<int:vendor_id>/toggle",
    methods=["POST"]
)
def toggle_finance_vendor(vendor_id):

    db_query("""
        update finance_vendors
        set is_active = not is_active
        where id = %s
    """, (
        vendor_id,
    ))

    return redirect(
        url_for(
            "finance.finance_vendors"
        )
    )


@finance_bp.route(
    "/expense/<category>",
    methods=["GET", "POST"]
)
def expense(category):

    message = ""

    fund_account = "观音堂日常户口"

    suggested_payment_voucher_no = (
        get_next_payment_voucher()
    )

    vendors = get_active_finance_vendors()

    sub_category_options = (
        EXPENSE_SUB_CATEGORY_OPTIONS.get(
            category,
            ["其它"]
        )
    )

    form_data = {
        "payment_voucher_no": suggested_payment_voucher_no,
        "record_date": date.today().isoformat(),
        "sub_category": "",
        "sub_category_custom": "",
        "vendor": "",
        "vendor_custom": "",
        "amount": "",
        "payment_method": "现金",
        "remarks": "",
    }

    if request.method == "POST":

        payment_voucher_no = request.form.get(
            "payment_voucher_no",
            ""
        ).strip().upper()

        record_date = (
            request.form.get("record_date")
            or date.today().isoformat()
        )

        sub_category_selected = request.form.get(
            "sub_category",
            ""
        ).strip()

        sub_category_custom = request.form.get(
            "sub_category_custom",
            ""
        ).strip()

        if sub_category_selected == "__custom__":
            sub_category = sub_category_custom
        else:
            sub_category = sub_category_selected

        vendor_selected = request.form.get(
            "vendor",
            ""
        ).strip()

        vendor_custom = request.form.get(
            "vendor_custom",
            ""
        ).strip()

        # 电费、水费、电话网络费：
        # 根据费用明细自动决定付款对象
        if category in AUTO_VENDOR_CATEGORIES:

            vendor = EXPENSE_VENDOR_BY_SUB_CATEGORY.get(
                sub_category,
                ""
            )

            # “其它”明细仍允许手动填写付款对象
            if not vendor:
                vendor = vendor_custom

        # 其它类别才读取付款对象
        elif vendor_selected == "__custom__":

            vendor = vendor_custom

        else:

            vendor = vendor_selected

        amount_raw = request.form.get(
            "amount",
            ""
        ).strip()

        amount = money(amount_raw)

        payment_method = request.form.get(
            "payment_method",
            "现金"
        ).strip()

        remarks = request.form.get(
            "remarks",
            ""
        ).strip()

        form_data = {
            "payment_voucher_no": payment_voucher_no,
            "record_date": str(record_date),
            "sub_category": sub_category_selected,
            "sub_category_custom": sub_category_custom,
            "vendor": vendor_selected,
            "vendor_custom": vendor_custom,
            "amount": amount_raw,
            "payment_method": payment_method,
            "remarks": remarks,
        }

        # =================================================
        # V6：检查该财政月份是否已经月结
        # =================================================
        month_lock_error = require_finance_month_open(
            record_date,
            fund_account
        )

        existing_voucher = None

        if payment_voucher_no:

            existing_voucher = db_query("""
                select id
                from finance_records
                where payment_voucher_no = %s
                limit 1
            """, (
                payment_voucher_no,
            ), fetchone=True)

        # 月结锁定必须放在所有验证的最前面
        if month_lock_error:

            message = month_lock_error

        elif not payment_voucher_no:

            message = "请填写 Payment Voucher 编号。"

        elif not re.match(
            r"^PV\d+$",
            payment_voucher_no
        ):

            message = (
                "Payment Voucher 格式错误，"
                "例如：PV000001。"
            )

        elif existing_voucher:

            message = (
                "这个 Payment Voucher 编号已经存在，"
                "请检查是否重复。"
            )

        elif not sub_category:

            message = "请选择或填写费用明细／单位。"

        elif category in AUTO_VENDOR_CATEGORIES and not vendor:

            message = "系统无法根据费用明细判断付款对象，请检查选项。"

        elif category not in AUTO_VENDOR_CATEGORIES and not vendor:

            message = "请选择或填写付款对象。"

        elif amount <= 0:

            message = "请输入正确的支出金额。"

        elif not remarks:

            message = "请填写支出用途或说明。"

        else:

            # 保存前再次检查，防止页面打开后才进行月结
            month_lock_error = require_finance_month_open(
                record_date,
                fund_account
            )

            if month_lock_error:

                message = month_lock_error

            else:

                db_query("""
                    insert into finance_records
                    (
                        record_type,
                        payment_voucher_no,
                        record_date,
                        category,
                        sub_category,
                        vendor,
                        name,
                        amount,
                        payment_method,
                        fund_account,
                        remarks
                    )
                    values
                    (
                        'expense',
                        %s, %s, %s, %s, %s, %s,
                        %s, %s, %s, %s
                    )
                """, (
                    payment_voucher_no,
                    record_date,
                    category,
                    sub_category,
                    vendor,
                    vendor,
                    amount,
                    payment_method,
                    fund_account,
                    remarks
                ))

                return redirect(
                    url_for(
                        "finance.records"
                    )
                )

    return render_template_string("""
    <!doctype html>
    <html lang="zh">
    <head>

        <meta charset="utf-8">

        <meta
            name="viewport"
            content="width=device-width, initial-scale=1"
        >

        <title>{{ category }}支出记录</title>

        <link
            rel="stylesheet"
            href="{{ url_for(
                'static',
                filename='css/toolbox.css'
            ) }}"
        >

        <style>

            .expense-page{
                max-width:820px;
            }

            .expense-header{
                background:linear-gradient(
                    135deg,
                    #dc2626,
                    #b91c1c
                );
                color:white;
                padding:28px;
                border-radius:22px;
                margin-bottom:20px;
            }

            .expense-header h1{
                margin:0 0 8px;
            }

            .expense-header p{
                margin:0;
                opacity:.92;
            }

            .expense-form-grid{
                display:grid;
                grid-template-columns:
                    repeat(2,minmax(0,1fr));
                gap:16px;
            }

            .full-width{
                grid-column:1 / -1;
            }

            .custom-box{
                display:none;
                margin-top:10px;
            }

            .voucher-box{
                background:#eff6ff;
                border:1px solid #bfdbfe;
                border-radius:14px;
                padding:16px;
                margin-bottom:18px;
            }

            .voucher-input{
                font-weight:800;
                color:#1d4ed8;
            }

            .expense-actions{
                display:flex;
                justify-content:space-between;
                gap:12px;
                flex-wrap:wrap;
                margin-top:22px;
            }

            @media(max-width:700px){

                .expense-form-grid{
                    grid-template-columns:1fr;
                }

                .full-width{
                    grid-column:auto;
                }

                .expense-actions{
                    display:grid;
                }

                .expense-actions .btn-tool{
                    width:100%;
                }
            }

        </style>

    </head>

    <body>

    <div class="page expense-page">

        <div class="expense-header">

            <h1>💸 {{ category }}</h1>

            <p>
                填写支出资料并保存 Payment Voucher。
            </p>

        </div>

        {% if message %}

            <div class="alert alert-danger">
                ⚠️ {{ message }}
            </div>

        {% endif %}

        <div class="card">

            <div class="section-title">
                🧾 支出资料
            </div>

            <form method="post">

                <div class="voucher-box">

                    <div class="form-group">

                        <label class="form-label">
                            Payment Voucher No.
                        </label>

                        <input
                            class="
                                form-input
                                voucher-input
                            "
                            name="payment_voucher_no"
                            value="{{ form_data.payment_voucher_no }}"
                            required
                        >

                        <div class="form-help">
                            系统建议：
                            {{ suggested_payment_voucher_no }}
                        </div>

                    </div>

                </div>

                <div class="expense-form-grid">

                    <div class="form-group">

                        <label class="form-label">
                            支出日期
                        </label>

                        <input
                            class="form-input"
                            name="record_date"
                            type="date"
                            value="{{ form_data.record_date }}"
                            required
                        >

                    </div>

                    <div class="form-group">

                        <label class="form-label">
                            付款方式
                        </label>

                        <select
                            class="form-input"
                            name="payment_method"
                            required
                        >

                            {% for method in [
                                '现金',
                                '银行过账',
                                '支票'
                            ] %}

                                <option
                                    value="{{ method }}"
                                    {% if
                                        form_data.payment_method
                                        == method
                                    %}
                                        selected
                                    {% endif %}
                                >
                                    {{ method }}
                                </option>

                            {% endfor %}

                        </select>

                    </div>

                    <div class="
                        form-group
                        full-width
                    ">

                        <label class="form-label">
                            费用明细／单位
                        </label>

                        <select
                            class="form-input"
                            name="sub_category"
                            id="sub_category"
                            onchange="toggleCustomSubCategory()"
                            required
                        >

                            <option value="">
                                请选择费用明细
                            </option>

                            {% for item in sub_category_options %}

                                <option
                                    value="{{ item }}"
                                    {% if form_data.sub_category == item %}
                                        selected
                                    {% endif %}
                                >
                                    {{ item }}
                                </option>

                            {% endfor %}

                            <option
                                value="__custom__"
                                {% if form_data.sub_category == '__custom__' %}
                                    selected
                                {% endif %}
                            >
                                其它／手动输入
                            </option>

                        </select>

                        <div
                            class="custom-box"
                            id="custom_sub_category_box"
                        >

                            <input
                                class="form-input"
                                name="sub_category_custom"
                                value="{{ form_data.sub_category_custom }}"
                                placeholder="填写费用明细／单位"
                            >

                        </div>

                        <div class="form-help">
                            例如：TNB 20-1、Air Selangor 20-2、
                            Celcom Reload 或佛具。
                        </div>

                    </div>
                                  
                    {% if category not in auto_vendor_categories %}

                    <div class="
                        form-group
                        full-width
                    ">

                        <label class="form-label">
                            付款对象
                        </label>

                        <select
                            class="form-input"
                            name="vendor"
                            id="vendor"
                            onchange="toggleCustomVendor()"
                            required
                        >

                            <option value="">
                                请选择付款对象
                            </option>

                            {% for vendor in vendors %}

                                <option
                                    value="{{ vendor.vendor_name }}"
                                    {% if
                                        form_data.vendor
                                        == vendor.vendor_name
                                    %}
                                        selected
                                    {% endif %}
                                >
                                    {{ vendor.vendor_name }}
                                </option>

                            {% endfor %}

                            <option
                                value="__custom__"
                                {% if
                                    form_data.vendor
                                    == '__custom__'
                                %}
                                    selected
                                {% endif %}
                            >
                                其它／手动输入
                            </option>

                        </select>

                        <div
                            class="custom-box"
                            id="custom_vendor_box"
                        >

                            <input
                                class="form-input"
                                name="vendor_custom"
                                value="{{ form_data.vendor_custom }}"
                                placeholder="填写其它付款对象"
                            >

                        </div>
                                  
                    </div>
                                  
                    {% else %}

                    <div class="form-group full-width">

                        <label class="form-label">
                            付款对象
                        </label>

                        <div class="alert alert-info">
                            系统会根据费用明细自动填写付款对象。
                        </div>

                    </div>

                    {% endif %}

                        <div class="form-help">
                            没有合适选项时，可选择“其它／手动输入”。
                        </div>

                    </div>

                    <div class="
                        form-group
                        full-width
                    ">

                        <label class="form-label">
                            支出金额 RM
                        </label>

                        <input
                            class="form-input"
                            name="amount"
                            type="number"
                            step="0.01"
                            min="0.01"
                            value="{{ form_data.amount }}"
                            required
                        >

                    </div>

                    <div class="
                        form-group
                        full-width
                    ">

                        <label class="form-label">
                            支出用途／备注
                        </label>

                        <textarea
                            class="form-input"
                            name="remarks"
                            rows="5"
                            required
                        >{{ form_data.remarks }}</textarea>

                    </div>

                </div>

                <div class="expense-actions">

                    <a
                        class="btn-tool btn-secondary"
                        href="{{ url_for(
                            'finance.finance_expense_menu'
                        ) }}"
                    >
                        ← 返回支出项目
                    </a>

                    <button
                        class="btn-tool btn-danger"
                        type="submit"
                        onclick="
                            return confirm(
                                '确定保存这笔支出？'
                            );
                        "
                    >
                        💾 保存支出
                    </button>

                </div>

            </form>

        </div>

        <div class="btn-row">

            <a
                class="btn-tool btn-primary"
                href="{{ url_for(
                    'finance.finance_vendors'
                ) }}"
            >
                🏢 管理付款对象
            </a>

        </div>

    </div>

    <script>

    function toggleCustomVendor(){

        const select = document.getElementById(
            "vendor"
        );

        const box = document.getElementById(
            "custom_vendor_box"
        );

        if(select.value === "__custom__"){
            box.style.display = "block";
        }else{
            box.style.display = "none";
        }
    }

    function toggleCustomSubCategory(){

        const select = document.getElementById(
            "sub_category"
        );

        const box = document.getElementById(
            "custom_sub_category_box"
        );

        if(select.value === "__custom__"){
            box.style.display = "block";
        }else{
            box.style.display = "none";
        }
    }

    toggleCustomVendor();
    toggleCustomSubCategory();

    </script>

    </body>
    </html>
    """,
        category=category,
        message=message,
        form_data=form_data,
        vendors=vendors,
        sub_category_options=sub_category_options,
        auto_vendor_categories=AUTO_VENDOR_CATEGORIES,
        suggested_payment_voucher_no=(
            suggested_payment_voucher_no
        )
    )

