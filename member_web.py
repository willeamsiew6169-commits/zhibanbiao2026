# member_web.py

import os
import psycopg2
import pandas as pd

from db import db_query
from opencc import OpenCC
from datetime import datetime
from psycopg2.extras import RealDictCursor
from flask import Blueprint, request, render_template_string, redirect, url_for, flash, session


cc = OpenCC('t2s')  # 繁 → 简

member_bp = Blueprint("member", __name__, url_prefix="/member")

DATABASE_URL = os.environ.get("DATABASE_URL")

MEMBER_ADMIN_PIN = os.environ.get("MEMBER_ADMIN_PIN", "1234")
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

def get_conn():
    return psycopg2.connect(DATABASE_URL)


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

@member_bp.route("/", methods=["GET", "POST"])
def member_home():
    member = None
    error = None
    paid_until = None
    summary = None
    payments = []

    if request.method == "POST":
        raw_member_id = request.form.get("member_id", "").strip()
        pin = request.form.get("pin", "").strip()

        member_id = normalize_member_id(raw_member_id)

        try:
            with get_conn() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:
                    keyword = raw_member_id.strip()
                    member_id = normalize_member_id(keyword)

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
                            limit 1
                        """, (
                            member_id,
                            f"%{keyword}%",
                            f"%{keyword}%"
                        ))

                    member = cur.fetchone()

                    print("找到会员:", member)

                    if not member:
                        error = "找不到这个会员，请检查编号 / 姓名 / 电话"
                    else:
                        if not verify_member_pin(member, pin):
                            error = "PIN 不正确"
                            member = None
                        else:
                            real_member_id = member["member_id"]

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

                            if summary and summary["paid_until"]:
                                paid_until = summary["paid_until"].strftime("%Y年%m月")

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
                            payments = cur.fetchall()

        except Exception as e:
            error = f"系统错误：{e}"

    return render_template_string(
        MEMBER_HTML,
        member=member,
        error=error,
        paid_until=paid_until,
        summary=summary,
        payments=payments
    )

@member_bp.route("/admin", methods=["GET", "POST"])
def member_admin():
    error = None
    member = None
    payments = []
    summary = None

    admin_pin = request.form.get("admin_pin", "").strip()
    raw_member_id = request.form.get("member_id", "").strip()

    if request.method == "POST":

        if not session.get("member_admin"):
            if admin_pin != MEMBER_ADMIN_PIN:
                error = "管理员 PIN 不正确"
            else:
                session["member_admin"] = True

        if not error:
            keyword = raw_member_id.strip()
            member_id = normalize_member_id(keyword)

            try:
                with get_conn() as conn:
                    with conn.cursor(cursor_factory=RealDictCursor) as cur:

                        cur.execute("""
                            select *
                            from members
                            where member_id = %s
                            or name ilike %s
                            or english_name ilike %s
                            limit 1
                        """, (
                            member_id,
                            f"%{keyword}%",
                            f"%{keyword}%"
                        ))
                        member = cur.fetchone()

                        if not member:
                            error = "找不到这个月费编号"
                        else:
                            real_member_id = member["member_id"]

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

                            cur.execute("""
                                select
                                    payment_date,
                                    receipt_no,
                                    start_month,
                                    end_month,
                                    month_count,
                                    amount,
                                    name
                                from member_payments
                                where member_id = %s
                                order by payment_date desc, receipt_no desc
                            """, (real_member_id,))
                            payments = cur.fetchall()

            except Exception as e:
                error = f"系统错误：{e}"

    return render_template_string(
        MEMBER_ADMIN_HTML,
        error=error,
        member=member,
        payments=payments,
        summary=summary,
        admin_pin=admin_pin,
        raw_member_id=raw_member_id
    )

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

@member_bp.route("/finance-upload", methods=["GET", "POST"])
def finance_upload():
    error = None
    msg = None

    if request.method == "POST":
        pin = request.form.get("pin", "").strip()

        if pin != FINANCE_PIN:
            error = "财政 PIN 不正确"
        else:
            file = request.files.get("file")

            if not file or file.filename == "":
                error = "请选择 Excel 文件"
            else:
                try:
                    df = pd.read_excel(file)

                    df = df[[
                        "日期\nDate",
                        "收据编号 \nOfficial Receipt No",
                        "编号 No/",
                        "捐款人\n姓名\nName",
                        "START MONTH",
                        "END MONTH",
                        "No/ of Mth",
                        "Total Amt"
                    ]]

                    # 只要求会员编号不为空；收据编号可以为空
                    df = df.dropna(subset=[
                        "编号 No/"
                    ])

                    inserted = 0
                    skipped = 0

                    for _, row in df.iterrows():
                        try:
                            receipt_raw = row["收据编号 \nOfficial Receipt No"]

                            if pd.isna(receipt_raw):
                                receipt_no = None
                            else:
                                receipt_no = str(receipt_raw).strip().replace(" ", "")
                                if receipt_no == "":
                                    receipt_no = None

                            member_no = int(row["编号 No/"])
                            member_id = f"CHE-{member_no}"

                            name = str(row["捐款人\n姓名\nName"]).strip()

                            payment_date = pd.to_datetime(row["日期\nDate"]).date()
                            start_month = pd.to_datetime(row["START MONTH"]).date()
                            end_month = pd.to_datetime(row["END MONTH"]).date()

                            month_count = int(row["No/ of Mth"])
                            amount = float(row["Total Amt"])

                            result = db_query("""
                                insert into member_payments
                                (
                                    receipt_no,
                                    member_id,
                                    name,
                                    payment_date,
                                    start_month,
                                    end_month,
                                    month_count,
                                    amount
                                )
                                values
                                (
                                    %s, %s, %s, %s,
                                    %s, %s, %s, %s
                                )
                                on conflict (receipt_no) do nothing
                                returning id
                            """, (
                                receipt_no,
                                member_id,
                                name,
                                payment_date,
                                start_month,
                                end_month,
                                month_count,
                                amount
                            ))

                            if result:
                                inserted += 1
                            else:
                                skipped += 1

                        except Exception as e:
                            print("导入失败:", row.to_dict(), e)
                            skipped += 1

                    msg = f"上传完成：读取 {len(df)} 行，新增 {inserted} 行，跳过 {skipped} 行。"

                except Exception as e:
                    print("上传失败:", e)
                    error = f"上传失败：{e}"

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>财政上传月费 Excel</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body{
    font-family:Arial,"Microsoft YaHei",sans-serif;
    background:#f3f4f6;
    margin:0;
    padding:18px;
}
.box{
    max-width:720px;
    margin:auto;
    background:white;
    padding:24px;
    border-radius:18px;
    box-shadow:0 3px 12px rgba(0,0,0,.08);
}
h1{
    text-align:center;
    font-size:42px;
    margin:20px 0 28px;
}
label{
    font-size:26px;
    font-weight:bold;
    display:block;
    margin-top:22px;
}
input{
    width:100%;
    font-size:28px;
    padding:18px;
    border:1px solid #ccc;
    border-radius:12px;
    box-sizing:border-box;
    margin-top:8px;
}
input[type=file]{
    background:#fafafa;
}
button{
    width:100%;
    font-size:34px;
    font-weight:bold;
    padding:22px;
    margin-top:28px;
    border:0;
    border-radius:14px;
    background:#2d7ff9;
    color:white;
}
.error{
    background:#ffe5e5;
    color:#a10000;
    font-size:22px;
    padding:16px;
    border-radius:12px;
    margin-bottom:16px;
}
.ok{
    background:#e8f7e8;
    color:#176b2c;
    font-size:22px;
    padding:16px;
    border-radius:12px;
    margin-bottom:16px;
}
.link{
    font-size:20px;
    margin-bottom:16px;
}
.link a{
    color:#555;
    text-decoration:none;
}
.note{
    background:#fff7d6;
    color:#6b4b00;
    padding:14px;
    border-radius:12px;
    font-size:20px;
    margin-top:18px;
}
@media(max-width:600px){
    body{padding:10px;}
    .box{padding:18px;border-radius:14px;}
    h1{font-size:34px;}
    label{font-size:24px;}
    input{font-size:26px;}
    button{font-size:30px;}
}
</style>
</head>
<body>

<div class="box">

    <div class="link">
        <a href="/member/admin">← 返回月费管理员</a>
    </div>

    <h1>财政上传月费 Excel</h1>

    {% if error %}
        <div class="error">❌ {{ error }}</div>
    {% endif %}

    {% if msg %}
        <div class="ok">✅ {{ msg }}</div>
    {% endif %}

    <form method="post" enctype="multipart/form-data">
        <label>财政 PIN</label>
        <input
            name="pin"
            type="password"
            inputmode="numeric"
            autocomplete="new-password"
            placeholder="请输入财政 PIN"
            required
        >

        <label>选择 Excel 文件</label>
        <input
            name="file"
            type="file"
            accept=".xlsx,.xls"
            required
        >

        <button type="submit">上传 Excel</button>
    </form>

    <div class="note">
        上传后系统会自动导入月费记录；收据编号空白的记录也会导入，重复收据编号会自动跳过。
    </div>

</div>

</body>
</html>
""", error=error, msg=msg)

MEMBER_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>月费查询</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body{
    font-family: Arial, "Microsoft YaHei", sans-serif;
    background:#f3f4f6;
    margin:0;
    padding:14px;
}
.box{
    max-width:780px;
    margin:auto;
    background:white;
    padding:22px;
    border-radius:18px;
    box-shadow:0 3px 12px rgba(0,0,0,.08);
}
.back{
    display:inline-block;
    margin-bottom:12px;
    text-decoration:none;
    color:#555;
    font-size:22px;
}
.admin-link{
    text-align:right;
    margin-bottom:8px;
}
.admin-link a{
    font-size:13px;
    color:#aaa;
    text-decoration:none;
}
h1{
    text-align:center;
    margin:8px 0 24px;
    font-size:46px;
}
label{
    font-weight:bold;
    font-size:28px;
    display:block;
    margin-top:18px;
}
input{
    width:100%;
    font-size:32px;
    padding:18px;
    box-sizing:border-box;
    border:1px solid #ccc;
    border-radius:12px;
    margin-top:8px;
}
button{
    width:100%;
    margin-top:24px;
    font-size:36px;
    padding:22px;
    border:0;
    border-radius:14px;
    background:#2d7ff9;
    color:white;
    font-weight:bold;
}
.error{
    background:#ffe5e5;
    color:#a10000;
    padding:16px;
    border-radius:12px;
    margin-bottom:16px;
    font-size:22px;
}
.change-pin-btn{
    display:block;
    text-align:center;
    margin-top:16px;
    padding:16px;
    background:#f0f0f0;
    border-radius:12px;
    color:#333;
    text-decoration:none;
    font-size:22px;
    font-weight:bold;
}
.result{
    margin-top:24px;
    background:#eef8ee;
    padding:20px;
    border-radius:16px;
}
.name{
    font-size:34px;
    font-weight:bold;
    text-align:center;
    margin-bottom:8px;
}
.info{
    font-size:22px;
    line-height:1.7;
}
.big-status{
    margin-top:18px;
    background:white;
    border-radius:16px;
    padding:22px;
    text-align:center;
    border:2px solid #bde5c8;
}
.status-title{
    font-size:24px;
    color:#555;
}
.status-month{
    font-size:42px;
    font-weight:bold;
    color:#1b7f3a;
    margin-top:8px;
}
.summary-grid{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:12px;
    margin-top:16px;
}
.summary-box{
    background:white;
    border-radius:14px;
    padding:16px;
    text-align:center;
    border:1px solid #ddd;
}
.summary-title{
    font-size:18px;
    color:#666;
}
.summary-value{
    font-size:28px;
    font-weight:bold;
    margin-top:6px;
}
.record-title{
    font-size:26px;
    font-weight:bold;
    margin-top:24px;
    margin-bottom:10px;
}
.record-card{
    background:white;
    border-radius:14px;
    padding:14px;
    margin-top:10px;
    border:1px solid #ddd;
    font-size:20px;
}
.record-date{
    font-weight:bold;
    font-size:22px;
}
.record-amount{
    font-weight:bold;
    color:#1b7f3a;
}
.no-record{
    margin-top:16px;
    background:#fff4d6;
    padding:14px;
    border-radius:12px;
    color:#7a4b00;
    font-size:20px;
}
@media(max-width:600px){
    body{padding:10px;}
    .box{padding:18px;}
    h1{font-size:38px;}
    label{font-size:25px;}
    input{font-size:28px;}
    button{font-size:32px;}
    .name{font-size:30px;}
    .status-month{font-size:36px;}
    .summary-grid{grid-template-columns:1fr;}
}
</style>
</head>
<body>

<div class="box">
    <a class="back" href="/">← 返回签到首页</a>

    <div class="admin-link">
        <a href="/member/admin">⚙ 管理员入口</a>
    </div>

    <h1>月费查询</h1>

    {% if error %}
    <div class="error">❌ {{ error }}</div>
    {% endif %}

    <form method="post">
        <label>月费编号 / 姓名 / 英文名 / 电话</label>
        <input name="member_id"
            placeholder="例如：CHE-108 / 0108 / 张三 / 0123456789"
            autocomplete="off"
            required>

        <label>PIN</label>
        <input
            id="member_pin"
            name="pin"
            type="password"
            inputmode="numeric"
            placeholder="请输入 PIN"
            autocomplete="new-password"
            autocorrect="off"
            autocapitalize="off"
            spellcheck="false"
            readonly
            onfocus="this.removeAttribute('readonly');"
            required
        >

        <button type="submit">查询</button>
    </form>

    <a href="/member/change-pin" class="change-pin-btn">
        🔒 更改月费密码
    </a>

    {% if member %}
    <div class="result">
        <div class="name">{{ member.name }}</div>

        <div class="info">
            {% if member.english_name %}
            英文名：{{ member.english_name }}<br>
            {% endif %}
            月费编号：{{ member.member_id }}
        </div>

        <div class="big-status">
            <div class="status-title">✅ 已供养至</div>
            <div class="status-month">{{ paid_until or "暂无记录" }}</div>
        </div>

        {% if summary %}
        <div class="summary-grid">
            <div class="summary-box">
                <div class="summary-title">总供养金额</div>
                <div class="summary-value">RM {{ "%.2f"|format(summary.total_payment or 0) }}</div>
            </div>

            <div class="summary-box">
                <div class="summary-title">总供养月数</div>
                <div class="summary-value">{{ summary.total_months or 0 }} 个月</div>
            </div>

            <div class="summary-box">
                <div class="summary-title">最近付款</div>
                <div class="summary-value">
                    {% if summary.last_payment_date %}
                        {{ summary.last_payment_date.strftime("%Y-%m-%d") }}
                    {% else %}
                        -
                    {% endif %}
                </div>
            </div>
        </div>
        {% endif %}

        <div class="record-title">最近付款记录</div>

        {% if payments %}
            {% for p in payments[:5] %}
            <div class="record-card">
                <div class="record-date">
                    {{ p.payment_date.strftime("%Y-%m-%d") if p.payment_date else "-" }}
                </div>
                <div>
                    供养月份：
                    {{ p.start_month.strftime("%Y-%m") if p.start_month else "-" }}
                    ~
                    {{ p.end_month.strftime("%Y-%m") if p.end_month else "-" }}
                </div>
                <div class="record-amount">
                    金额：RM {{ "%.2f"|format(p.amount or 0) }}
                </div>
            </div>
            {% endfor %}
        {% else %}
            <div class="no-record">暂无付款记录。</div>
        {% endif %}
    </div>
    {% endif %}
</div>

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
<style>
body{
    font-family:Arial,"Microsoft YaHei",sans-serif;
    background:#f5f5f5;
    margin:0;
    padding:20px;
}
.box{
    max-width:1000px;
    margin:auto;
    background:white;
    padding:24px;
    border-radius:16px;
    box-shadow:0 2px 10px rgba(0,0,0,.08);
}
a{
    text-decoration:none;
    color:#555;
}
h1{
    text-align:center;
    font-size:48px;
}
label{
    font-weight:bold;
    font-size:28px;
    display:block;
    margin-top:14px;
}
input{
    width:100%;
    font-size:30px;
    padding:16px;
    box-sizing:border-box;
    border:1px solid #ccc;
    border-radius:10px;
    margin-top:6px;
}
button{
    width:100%;
    margin-top:20px;
    font-size:32px;
    padding:18px;
    border:0;
    border-radius:12px;
    background:#2d7ff9;
    color:white;
    font-weight:bold;
}
.error{
    background:#ffe5e5;
    color:#a10000;
    padding:12px;
    border-radius:10px;
    margin-bottom:12px;
}
.member-card{
    margin-top:20px;
    padding:18px;
    background:#f7f7f7;
    border-radius:12px;
}
.summary-grid{
    display:grid;
    grid-template-columns:repeat(2, 1fr);
    gap:12px;
    margin-top:18px;
}
.summary-box{
    background:white;
    border:1px solid #ddd;
    border-radius:12px;
    padding:16px;
}
.summary-title{
    color:#666;
    font-size:18px;
}
.summary-value{
    font-size:30px;
    font-weight:bold;
    margin-top:6px;
}
table{
    width:100%;
    border-collapse:collapse;
    margin-top:18px;
    background:white;
}
th,td{
    border:1px solid #ddd;
    padding:10px;
    text-align:center;
    font-size:16px;
}
th{
    background:#eeeeee;
}
.no-record{
    margin-top:18px;
    background:#fff4d6;
    padding:14px;
    border-radius:10px;
    color:#7a4b00;
}
@media(max-width:700px){
    h1{font-size:36px;}
    .summary-grid{grid-template-columns:1fr;}
    th,td{font-size:13px;padding:6px;}
}
</style>
</head>
<body>

<div class="box">
    <a href="/member">← 返回月费查询</a>
    &nbsp; | &nbsp;
    <a href="/">返回签到首页</a>
    &nbsp; | &nbsp;
    <a href="/member/finance-upload">财政上传Excel</a>

    <h1>月费管理员</h1>

    <div style="margin-bottom:15px;">
        <a href="/member/admin/logout"
        style="color:red;font-size:18px;text-decoration:none;">
        🚪 退出管理员
        </a>
    </div>

    {% if error %}
    <div class="error">{{ error }}</div>
    {% endif %}

    <form method="post">
        <label>管理员 PIN</label>
        <input
            id="member_admin_pin"
            name="admin_pin"
            type="password"
            inputmode="numeric"
            autocomplete="new-password"
            autocorrect="off"
            autocapitalize="off"
            spellcheck="false"
            value=""
            readonly
            onfocus="this.removeAttribute('readonly');"
            required
        >

        <label>月费编号 / 姓名</label>
        <input name="member_id" value="{{ raw_member_id }}" placeholder="例如：CHE-3 / Anna" required>

        <button type="submit">查找会员</button>
    </form>

    {% if member %}
    <div class="member-card">
        <h2>{{ member.name }}</h2>

        {% if member.english_name %}
        <div>英文名：{{ member.english_name }}</div>
        {% endif %}

        <div>电话：{{ member.phone or "-" }}</div>
        <div>月费编号：{{ member.member_id }}</div>

        {% if summary %}
        <div class="summary-grid">
            <div class="summary-box">
                <div class="summary-title">总供养金额</div>
                <div class="summary-value">RM {{ "%.2f"|format(summary.total_payment or 0) }}</div>
            </div>

            <div class="summary-box">
                <div class="summary-title">总供养月数</div>
                <div class="summary-value">{{ summary.total_months or 0 }} 个月</div>
            </div>

            <div class="summary-box">
                <div class="summary-title">已供养至</div>
                <div class="summary-value">
                    {% if summary.paid_until %}
                        {{ summary.paid_until.strftime("%Y-%m") }}
                    {% else %}
                        暂无记录
                    {% endif %}
                </div>
            </div>

            <div class="summary-box">
                <div class="summary-title">最后付款日期</div>
                <div class="summary-value">
                    {% if summary.last_payment_date %}
                        {{ summary.last_payment_date.strftime("%Y-%m-%d") }}
                    {% else %}
                        暂无记录
                    {% endif %}
                </div>
            </div>
        </div>
        {% endif %}

        {% if payments %}
        <h3>付款记录</h3>
        <table>
            <tr>
                <th>付款日期</th>
                <th>收据编号</th>
                <th>Start Month</th>
                <th>End Month</th>
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
        {% else %}
        <div class="no-record">
            这个会员目前没有财政 Excel 付款记录。
        </div>
        {% endif %}
    </div>
    {% endif %}
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
<style>
body{font-family:Arial,"Microsoft YaHei",sans-serif;background:#f5f5f5;margin:0;padding:20px;}
.box{max-width:480px;margin:auto;background:white;padding:24px;border-radius:16px;box-shadow:0 2px 10px rgba(0,0,0,.08);}
a{text-decoration:none;color:#555;}
h1{text-align:center;}
label{font-weight:bold;display:block;margin-top:14px;}
input{width:100%;font-size:20px;padding:12px;box-sizing:border-box;border:1px solid #ccc;border-radius:10px;margin-top:6px;}
button{width:100%;margin-top:22px;font-size:22px;padding:14px;border:0;border-radius:12px;background:#2d7ff9;color:white;font-weight:bold;}
.error{background:#ffe5e5;color:#a10000;padding:12px;border-radius:10px;margin-bottom:15px;}
.ok{background:#e8f7e8;color:#176b2c;padding:12px;border-radius:10px;margin-bottom:15px;}
</style>
</head>
<body>
<div class="box">
    <a href="/member">← 返回月费查询</a>
    <h1>更改月费密码</h1>

    {% if error %}
    <div class="error">{{ error }}</div>
    {% endif %}

    {% if ok %}
    <div class="ok">{{ ok }}</div>
    {% endif %}

    <form method="post">

    <label>月费编号</label>
    <input
        name="member_id"
        placeholder="例如：108 / CHE-108 / 0108"
        autocomplete="off"
        required
    >

    <label>旧密码</label>
    <input
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

    <label>新密码</label>
    <input
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

    <label>确认新密码</label>
    <input
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

    <button type="submit">确认更改</button>

</form>
</div>
</body>
</html>
"""