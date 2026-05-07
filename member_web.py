# member_web.py

import os
import psycopg2
from opencc import OpenCC
from psycopg2.extras import RealDictCursor
from flask import Blueprint, request, render_template_string, redirect, url_for, flash


cc = OpenCC('t2s')  # 繁 → 简

member_bp = Blueprint("member", __name__, url_prefix="/member")

DATABASE_URL = os.environ.get("DATABASE_URL")

MEMBER_ADMIN_PIN = os.environ.get("MEMBER_ADMIN_PIN", "1234")

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

    if request.method == "POST":
        raw_member_id = request.form.get("member_id")
        pin = request.form.get("pin")

        member_id = normalize_member_id(raw_member_id)

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
                    else:
                        if not verify_member_pin(member, pin):
                            error = "PIN 不正确"
                            member = None
                        else:
                            cur.execute("""
                                select paid_month
                                from member_payments
                                where member_id = %s
                                order by paid_month
                            """, (member_id,))

                            paid_rows = cur.fetchall()
                            paid_months = [r["paid_month"] for r in paid_rows]

                            if paid_months:
                                latest = max(paid_months)
                                y, m = latest.split("-")
                                paid_until = f"{y}年{int(m)}月"

        except Exception as e:
            error = f"系统错误：{e}"

    return render_template_string(MEMBER_HTML, member=member, error=error, paid_until=paid_until)

@member_bp.route("/admin", methods=["GET", "POST"])
def member_admin():
    error = None
    member = None
    paid_months = []

    admin_pin = request.form.get("admin_pin", "").strip()
    raw_member_id = request.form.get("member_id", "").strip()
    action = request.form.get("action", "")

    if request.method == "POST":
        if admin_pin != MEMBER_ADMIN_PIN:
            error = "管理员 PIN 不正确"
        else:
            member_id = normalize_member_id(raw_member_id)

            try:
                with get_conn() as conn:
                    with conn.cursor(cursor_factory=RealDictCursor) as cur:

                        keyword = raw_member_id.strip()
                        member_id = normalize_member_id(keyword)

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

                        real_member_id = member["member_id"] if member else member_id

                        if not member:
                            error = "找不到这个月费编号"
                        else:
                            if action == "save":
                                selected_months = request.form.getlist("paid_months")
                                remark = request.form.get("remark", "").strip()

                                cur.execute("""
                                    delete from member_payments
                                    where member_id = %s
                                    and paid_month like %s
                                """, (real_member_id, f"{PAYMENT_YEAR}-%"))

                                for month in selected_months:
                                    cur.execute("""
                                        insert into member_payments
                                        (member_id, paid_month, remark)
                                        values (%s, %s, %s)
                                    """, (real_member_id, month, remark))

                                conn.commit()
                                paid_months = selected_months

                                if paid_months:
                                    latest = max(paid_months)
                                    y, m = latest.split("-")
                                    paid_until = f"{y}年{int(m)}月"
                                else:
                                    paid_until = "暂无记录"

                                missing_months = check_missing_months(paid_months)

                                if missing_months:
                                    missing_text = "、".join([f"{int(m.split('-')[1])}月" for m in missing_months])
                                    flash(f"已保存，但注意：中间漏了 {missing_text}", "error")
                                else:
                                    flash(f"已保存月费记录，已供养至：{paid_until}", "ok")

                            else:
                                cur.execute("""
                                    select paid_month
                                    from member_payments
                                    where member_id = %s
                                    and paid_month like %s
                                    order by paid_month
                                """, (real_member_id, f"{PAYMENT_YEAR}-%"))

                                paid_months = [r["paid_month"] for r in cur.fetchall()]

            except Exception as e:
                error = f"系统错误：{e}"

    return render_template_string(
        MEMBER_ADMIN_HTML,
        error=error,
        member=member,
        paid_months=paid_months,
        months=MONTHS,
        admin_pin=admin_pin,
        raw_member_id=raw_member_id,
        year=PAYMENT_YEAR
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
    background:#f5f5f5;
    margin:0;
    padding:20px;
}
.box{
    max-width:480px;
    margin:auto;
    background:white;
    padding:24px;
    border-radius:16px;
    box-shadow:0 2px 10px rgba(0,0,0,.08);
}
.back{
    display:inline-block;
    margin-bottom:18px;
    text-decoration:none;
    color:#555;
    font-size:16px;
}
h1{
    text-align:center;
    margin:10px 0 24px;
}
label{
    font-weight:bold;
    display:block;
    margin-top:14px;
}
input{
    width:100%;
    font-size:20px;
    padding:12px;
    box-sizing:border-box;
    border:1px solid #ccc;
    border-radius:10px;
    margin-top:6px;
}
button{
    width:100%;
    margin-top:22px;
    font-size:22px;
    padding:14px;
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
    margin-bottom:15px;
}
.result{
    margin-top:20px;
    background:#eef8ee;
    padding:16px;
    border-radius:12px;
    font-size:18px;
}
.small{
    color:#777;
    font-size:14px;
}
.name{
    font-size:26px;
    font-weight:bold;
    text-align:center;
    margin-bottom:12px;
}
.paid-title{
    text-align:center;
    color:#666;
    font-size:16px;
}
.paid-month{
    text-align:center;
    font-size:30px;
    font-weight:bold;
    color:#1b7f3a;
    margin-top:6px;
}
.change-pin-btn{
    display:block;
    text-align:center;
    margin-top:15px;
    padding:12px;
    background:#f0f0f0;
    border-radius:10px;
    color:#333;
    text-decoration:none;
    font-weight:bold;
}
hr{
    border:0;
    border-top:1px solid #ddd;
    margin:16px 0;
}
</style>
</head>
<body>

<div class="box">
    <a class="back" href="/">← 返回签到首页</a>

    <div style="text-align:right; margin-bottom:10px;">
        <a href="/member/admin" style="font-size:12px; color:#aaa;">
            ⚙ 管理员入口
        </a>
    </div>

    <h1>月费查询</h1>

    {% if error %}
    <div class="error">{{ error }}</div>
    {% endif %}

    <form method="post">
        <label>月费编号</label>
        <input name="member_id"
            placeholder="例如：208 / CHE-208 / 0208"
            autocomplete="off"
            required>

        <label>PIN</label>
        <input
            id="member_pin"
            name="pin"
            type="password"
            placeholder="请输入 PIN"
            autocomplete="new-password"
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

        {% if member.english_name %}
        <div>英文名：{{ member.english_name }}</div>
        {% endif %}

        <div>月费编号：{{ member.member_id }}</div>
        <div>电话：{{ member.phone or "-" }}</div>

        <hr>

        <div class="paid-title">已供养至</div>
        <div class="paid-month">{{ paid_until or "暂无记录" }}</div>
        
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
    max-width:620px;
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
}
label{
    font-weight:bold;
    display:block;
    margin-top:14px;
}
input{
    width:100%;
    font-size:18px;
    padding:11px;
    box-sizing:border-box;
    border:1px solid #ccc;
    border-radius:10px;
    margin-top:6px;
}
button{
    width:100%;
    margin-top:20px;
    font-size:20px;
    padding:13px;
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
}
.ok{
    background:#e8f7e8;
    color:#176b2c;
    padding:12px;
    border-radius:10px;
}
.member-card{
    margin-top:20px;
    padding:15px;
    background:#f7f7f7;
    border-radius:12px;
}
.month-grid{
    display:grid;
    grid-template-columns:repeat(3, 1fr);
    gap:10px;
    margin-top:15px;
}
.month{
    background:white;
    border:1px solid #ddd;
    border-radius:10px;
    padding:12px;
    font-size:18px;
}
.month input{
    width:auto;
    transform:scale(1.3);
    margin-right:8px;
}
</style>
</head>
<body>

<div class="box">
    <a href="/member">← 返回月费查询</a>
    &nbsp; | &nbsp;
    <a href="/">返回签到首页</a>

    <h1>月费管理员</h1>

    {% if error %}
    <div class="error">{{ error }}</div>
    {% endif %}

    {% with messages = get_flashed_messages(with_categories=true) %}
      {% for category, message in messages %}
        <div class="{{ category }}">{{ message }}</div>
      {% endfor %}
    {% endwith %}

    <form method="post">
        <input type="hidden" name="action" value="search">

        <label>管理员 PIN</label>
        <input name="admin_pin" type="password" value="{{ admin_pin }}" required>

        <label>月费编号 / 姓名 / 电话</label>
        <input name="member_id" value="{{ raw_member_id }}" placeholder="例如：输入编号 / 姓名" required>

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

        <form method="post">
            <input type="hidden" name="action" value="save">
            <input type="hidden" name="admin_pin" value="{{ admin_pin }}">
            <input type="hidden" name="member_id" value="{{ member.member_id }}">

            <h3>{{ year }} 年供养月份</h3>

            <div class="month-grid">
            {% for m in months %}
                <label class="month">
                    <input type="checkbox" name="paid_months" value="{{ m }}"
                    {% if m in paid_months %}checked{% endif %}>
                    {{ m[5:7] }}月
                </label>
            {% endfor %}
            </div>

            <label>备注</label>
            <input name="remark" placeholder="可空，例如：现金 / 转账 / 财政补录">

            <button type="submit">保存月费记录</button>
        </form>
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
        <input name="member_id" placeholder="例如：208 / CHE-208 / 0208" required>

        <label>旧密码</label>
        <input name="old_pin" type="password" required>

        <label>新密码</label>
        <input name="new_pin" type="password" required>

        <label>确认新密码</label>
        <input name="confirm_pin" type="password" required>

        <button type="submit">确认更改</button>
    </form>
</div>
</body>
</html>
"""