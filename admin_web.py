# admin_web.py

from utils import get_today_code
from flask import Blueprint, render_template_string

admin_bp = Blueprint("admin", __name__)

ADMIN_HOME_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<link rel="manifest" href="/manifest.json">

<title>观音堂管理员入口</title>

<style>

body{
    margin:0;
    padding:0;
    background:#f5efe3;
    font-family:"Microsoft YaHei";
    text-align:center;
}

.container{
    max-width:500px;
    margin:auto;
    padding:30px 20px;
}

h1{
    color:#8b5a2b;
}

.card{
    background:white;
    border-radius:20px;
    padding:25px;
    margin-bottom:25px;
    box-shadow:0 4px 12px rgba(0,0,0,0.08);
}

.btn{
    display:block;
    background:#b67b3d;
    color:white;
    text-decoration:none;
    padding:16px;
    border-radius:14px;
    font-size:18px;
    font-weight:bold;
}

</style>
</head>

<body>

<div class="container">

<h1>🙏 观音堂管理员入口</h1>

<div class="card">
<h2>📋 值班表生成系统</h2>
<a class="btn" href="/schedule">
进入系统
</a>
</div>

<div class="card">
<h2>💰 月费管理员系统</h2>
<a class="btn" href="/member">
进入系统
</a>
</div>

<div class="card">
<h2>✅ 义工签到系统</h2>
<a class="btn" href="/">
进入系统
</a>
</div>

<div class="card">
    <h2>📅 今日签到码</h2>

    <div style="
        font-size:48px;
        font-weight:bold;
        text-align:center;
        padding:20px;
        background:#fff3cd;
        border-radius:16px;
        color:#856404;
    ">
        {{ today_code }}
    </div>
</div>

</div>

</body>
</html>
"""

@admin_bp.route("/admin-home")
def admin_home():
    return render_template_string(
        ADMIN_HOME_HTML,
        today_code=get_today_code()
    )