# member_web.py

from flask import Blueprint, request, render_template_string, redirect, url_for, flash
import os
import psycopg2
from psycopg2.extras import RealDictCursor

member_bp = Blueprint("member", __name__, url_prefix="/member")

DATABASE_URL = os.environ.get("DATABASE_URL")


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


def verify_member_pin(member, pin):
    pin = str(pin or "").strip()

    db_pin = str(member.get("pin") or "").strip()
    phone = str(member.get("phone") or "").strip()

    default_pin = phone[-4:] if len(phone) >= 4 else ""

    return pin and (pin == db_pin or pin == default_pin)


@member_bp.route("/", methods=["GET", "POST"])
def member_home():
    member = None
    error = None

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
            elif not verify_member_pin(member, pin):
                error = "PIN 不正确"
                member = None

        except Exception as e:
            error = f"系统错误：{e}"

    return render_template_string(MEMBER_HTML, member=member, error=error)


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
</style>
</head>
<body>

<div class="box">
    <a class="back" href="/">← 返回签到首页</a>

    <h1>月费查询</h1>

    {% if error %}
    <div class="error">{{ error }}</div>
    {% endif %}

    <form method="post">
        <label>月费编号</label>
        <input name="member_id" placeholder="例如：208 / CHE-208 / 0208" autocomplete="off" required>

        <label>PIN</label>
        <input name="pin" type="password" placeholder="请输入 PIN" required>

        <button type="submit">查询</button>
    </form>

    {% if member %}
    <div class="result">
        <b>{{ member.name }}</b><br>
        编号：{{ member.member_id }}<br>
        电话：{{ member.phone or "-" }}<br>
        状态：{{ member.status or "-" }}<br>
        <br>
        <span class="small">下一步这里会显示：已供养到几月</span>
    </div>
    {% endif %}
</div>

</body>
</html>
"""