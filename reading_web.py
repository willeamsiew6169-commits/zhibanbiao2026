# reading_web.py

import os
import re
import pandas as pd

from db import db_query, get_conn
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from flask import (
    Blueprint,
    request,
    redirect,
    url_for,
    render_template_string
)

from datetime import datetime
from utils import get_text, now_date_str, MY_TZ

reading_bp = Blueprint("reading", __name__)
READING_FILE = "reading.xlsx"

def ensure_reading_file():
    required_cols = ["日期", "姓名", "主题", "场次", "时间"]

    if not os.path.exists(READING_FILE):
        df = pd.DataFrame(columns=required_cols)
        df.to_excel(READING_FILE, index=False)
        return

    df = pd.read_excel(READING_FILE)

    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    df = df[required_cols]
    df.to_excel(READING_FILE, index=False)
    beautify_reading_excel()

def beautify_reading_excel():
    if not os.path.exists(READING_FILE):
        return
    
    print("🔥 正在执行 beautify")

    wb = load_workbook(READING_FILE)
    ws = wb.active

    header_fill = PatternFill("solid", fgColor="D9EAD3")
    header_font = Font(name="Microsoft YaHei", size=16, bold=True)
    body_font = Font(name="Microsoft YaHei", size=15)

    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows():
        for cell in row:
            cell.font = header_font if cell.row == 1 else body_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = border
            if cell.row == 1:
                cell.fill = header_fill

    # ✅ 行高：重点
    ws.row_dimensions[1].height = 32

    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 30

    # ✅ 列宽：放大
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 18

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{ws.max_row}"

    wb.save(READING_FILE)

def get_today_reading_rows():
    rows = db_query("""
        select *
        from reading
        where date = %s
        order by id desc
    """, (now_date_str(),), fetchall=True)

    return rows or []



def add_reading_record(name, identity, topic, session, time_text):
    db_query("""
        insert into reading
        (date, name, identity, topic, session, time)
        values (%s, %s, %s, %s, %s, %s)
    """, (
        now_date_str(),
        name,
        identity,
        topic,
        session,
        time_text
    ))

def add_reading_once_today(name, identity, topic, time_text):
    """同一天同一姓名只记一次；事务锁避免两台设备同时重复提交。"""
    from psycopg2.extras import RealDictCursor
    day = now_date_str()
    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            # 相同日期与姓名的提交串行处理，不需要修改数据库表结构。
            cur.execute("SELECT pg_advisory_xact_lock(hashtext(%s), hashtext(%s))", (day, name))
            cur.execute("SELECT 1 FROM reading WHERE date=%s AND name=%s LIMIT 1", (day, name))
            if cur.fetchone():
                return False
            cur.execute("""
                INSERT INTO reading (date, name, identity, topic, session, time)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (day, name, identity, topic, "", time_text))
    return True


def get_reading_volunteer_by_number(number):
    """按完整 CHE/STW 编号查义工，避免仅凭数字误认其他分会。"""
    number = str(number or "").strip().upper().replace(" ", "")
    match = re.fullmatch(r"(CHE|STW)-?(\d+)", number)
    if not match:
        return None
    branch, digits = match.groups()
    normalized = f"{branch}-{int(digits)}"
    rows = db_query("""
        select id, name from volunteers
        where upper(id) = %s or upper(coalesce(member_id, '')) = %s
        limit 2
    """, (normalized, normalized), fetchall=True) or []
    names = {str(r.get("name") or "").strip() for r in rows if r.get("name")}
    return next(iter(names)) if len(names) == 1 else None


def get_reading_topics():
    rows = db_query("""
        select distinct topic
        from reading
        where topic is not null and topic <> ''
        order by topic
    """, fetchall=True)

    return [r["topic"] for r in rows if r.get("topic")]

def get_today_attendees():
    rows = db_query("""
        select distinct name
        from attendance
        where date = %s and signin = 1
    """, (now_date_str(),), fetchall=True)

    return sorted([r["name"] for r in rows if r.get("name")])


@reading_bp.route("/reading", methods=["GET", "POST"])
def reading():
    t = get_text()
    attendees = get_today_attendees()
    today = now_date_str()

    if request.method == "POST":
        topic = request.form.get("topic", "").strip()
        names = request.form.getlist("names")
        extra_text = request.form.get("extra_names", "").strip()
        volunteer_number = request.form.get("volunteer_number", "").strip()

        extra_names = []
        if extra_text:
            for sep in ["，", "、", ",", "\n"]:
                extra_text = extra_text.replace(sep, " ")
            extra_names = [x.strip() for x in extra_text.split(" ") if x.strip()]

        if not names and not extra_names and not volunteer_number:
            return "❌ 请选择至少一位共修者。<br><a href='/reading'>返回</a>"

        if not topic:
            return f"❌ {t['enter_topic']}<br><a href='/reading'>{t['back_home']}</a>"

        now_time = datetime.now(MY_TZ).strftime("%I:%M %p").lstrip("0")

        # 同一表单重复勾选／填写同名，也只尝试一次。
        unique_people = {}
        for name in names:
            clean_name = str(name).strip()
            if clean_name:
                unique_people[clean_name] = t.get("volunteer_identity", "Volunteer" if t.get("html_lang", "zh").startswith("en") else "义工")
        for name in extra_names:
            clean_name = str(name).strip()
            if clean_name and clean_name not in unique_people:
                unique_people[clean_name] = t.get("friend_identity", "Friend" if t.get("html_lang", "zh").startswith("en") else "佛友")
        if volunteer_number:
            volunteer_name = get_reading_volunteer_by_number(volunteer_number)
            if not volunteer_name:
                return "❌ 找不到此 CHE/STW 义工编号，请检查编号。<br><a href='/reading'>返回</a>", 400
            unique_people[volunteer_name] = t.get("volunteer_identity", "Volunteer" if t.get("html_lang", "zh").startswith("en") else "义工")
        for name, identity in unique_people.items():
            add_reading_once_today(name, identity, topic, now_time)

        return redirect(url_for("reading.reading"))

    rows = get_today_reading_rows()

    today_records = []
    for r in rows:
        today_records.append({
            "id": r.get("id"),
            "日期": r.get("date"),
            "姓名": r.get("name"),
            "身份": r.get("identity"),
            "主题": r.get("topic"),
            "场次": r.get("session"),
            "时间": r.get("time"),
        })

    summary_map = {}
    for r in today_records:
        key = (r["姓名"], r["身份"])
        summary_map[key] = summary_map.get(key, 0) + 1

    today_summary_records = [
        {
            "姓名": name,
            "身份": identity,
            "共修次数": count
        }
        for (name, identity), count in summary_map.items()
    ]

    # 已点名者从待选名单隐藏，其他设备旧页面提交也由数据库事务再次检查。
    recorded_names = {str(r.get("name") or "").strip() for r in rows}
    attendees = [name for name in attendees if name not in recorded_names]
    topic_options = get_reading_topics()

    html = """
    <!doctype html>
    <html lang="{{ t.html_lang }}">
    <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
          *{box-sizing:border-box}
          body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#f6f7f2;margin:0;padding:16px;color:#26352a;font-size:17px}
          .wrap{max-width:950px;margin:auto}
          .card{background:#fff;padding:22px;border-radius:16px;margin:14px 0;box-shadow:0 2px 10px #00000012}
          h2{font-size:23px;margin:0 0 18px} h3{font-size:19px}
          .field-label{display:block;margin:16px 0 8px;font-weight:700}
          input[type=text]{width:100%;font-size:18px;padding:13px;border:1px solid #cbd5cc;border-radius:10px}
          .branch-number-row{display:flex;gap:10px;align-items:stretch}
          .branch-switch{flex:0 0 88px;font-weight:700;font-size:18px;border-radius:10px}
          .branch-switch.che{background:#399d55;color:#fff}
          .branch-switch.stw{background:#dc3545;color:#fff}
          .branch-number-row input[type=text]{flex:1;min-width:0}
          .name-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(145px,1fr));gap:9px;margin:12px 0}
          .name-item{display:flex;align-items:center;gap:8px;padding:12px;background:#f4f8f3;border-radius:10px;font-size:17px}
          .name-item input{width:19px;height:19px}
          button,.button-link{display:inline-flex;align-items:center;justify-content:center;gap:5px;border:0;border-radius:10px;padding:11px 16px;background:#399d55;color:white;font-size:16px;cursor:pointer;text-decoration:none}
          .button-link{margin:0 8px 8px 0}.secondary{background:#4778bb}.edit{background:#e6a23c}.delete{background:#d9534f}
          .actions{display:flex;gap:7px;flex-wrap:wrap}.actions button,.actions .button-link{font-size:14px;padding:8px 10px;margin:0}
          .record{border:1px solid #e2e9e2;border-radius:12px;padding:13px;margin:10px 0}
          .record-head{display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:8px}
          .record-meta{font-size:14px;color:#66736a;margin-bottom:10px;overflow-wrap:anywhere}
          .muted{color:#718074;font-size:15px}
          details summary{cursor:pointer;font-size:19px;font-weight:700;padding:5px 0}
          @media(max-width:600px){body{padding:9px;font-size:16px}.card{padding:15px;border-radius:13px}h2{font-size:20px}.name-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.name-item{padding:10px;font-size:16px}.record-head{align-items:flex-start}.actions{width:100%}.actions .button-link{flex:1}}
        </style>
    </head>
    <body><main class="wrap">
      <nav>
        <a class="button-link" href="/">⬅ {{ t.back_home }}</a>
        <a class="button-link secondary" href="/download_reading">{{ t.download_report }}</a>
      </nav>
      <section class="card">
        <h2>{{ t.bhff_record_title }}</h2>
        <form method="post">
          <label class="field-label">{{ t.today_topic }}</label>
          <input type="text" name="topic" list="topicList" placeholder="{{ t.topic_placeholder }}" required>
          <datalist id="topicList">{% for topic in topic_options %}<option value="{{ topic }}">{% endfor %}</datalist>
          <label class="field-label">{{ t.volunteer_list }}（{{ attendees|length }}）</label>
          <button type="button" onclick="selectAllNames()">{{ t.select_all_volunteers }}</button>
          {% if attendees %}
          <div class="name-grid">
          {% for name in attendees %}<label class="name-item"><input type="checkbox" name="names" value="{{ name }}">{{ name }}</label>{% endfor %}
          </div>
          {% else %}<p class="muted">今日签到义工均已点名，或暂无可点名义工。</p>{% endif %}
          <label class="field-label">{{ '未签到的义工：输入 CHE/STW 编号' if not t.html_lang.startswith('en') else 'Volunteer not yet checked in: enter CHE/STW number' }}</label>
          <div class="branch-number-row">
            <button type="button" id="readingBranchBtn" class="branch-switch che" onclick="toggleReadingBranch()" aria-label="切换 CHE 或 STW">CHE</button>
            <input type="text" id="readingNumberDigits" inputmode="numeric" autocomplete="off" pattern="[0-9]+" placeholder="例如：208" title="只需输入数字编号">
          </div>
          <input type="hidden" name="volunteer_number" id="readingVolunteerNumber">
          <p class="muted">{{ '不必填写姓名；保存时自动从义工资料读取姓名。' if not t.html_lang.startswith('en') else 'No name needed; the volunteer name is looked up when saving.' }}</p>
          <label class="field-label">{{ t.extra_friend_name }}</label>
          <input type="text" name="extra_names" placeholder="{{ t.extra_friend_placeholder }}">
          <p class="muted">{{ t.extra_friend_tip }}</p>
          <button type="submit">{{ t.record_study }}</button>
        </form>
      </section>
      <section class="card">
        <details>
          <summary>📋 {{ t.today_recorded }}（{{ today_records|length }}）</summary>
          {% for r in today_records %}
          <div class="record">
            <div class="record-head"><strong>{{ r["姓名"] }}</strong><span class="muted">{{ r["时间"] }}</span></div>
            <div class="record-meta">{{ r["身份"] }} · {{ r["主题"] }}</div>
            <div class="actions">
              <a class="button-link edit" href="/reading_edit/{{ r['id'] }}">{{ t.edit }}</a>
              <a class="button-link delete" href="/reading_delete/{{ r['id'] }}" onclick="return confirm({{ t.get("delete_confirm_simple", "确定删除这笔共修记录吗？" if not t.get("html_lang", "zh").startswith("en") else "Delete this study record?")|tojson }})">{{ t.delete_record }}</a>
            </div>
          </div>
          {% else %}<p class="muted">今日暂无记录。</p>{% endfor %}
        </details>
      </section>
    </main>

    <script>
    let readingBranch = 'CHE';
    function toggleReadingBranch() {
        readingBranch = readingBranch === 'CHE' ? 'STW' : 'CHE';
        const btn = document.getElementById('readingBranchBtn');
        btn.textContent = readingBranch;
        btn.className = 'branch-switch ' + readingBranch.toLowerCase();
    }
    function syncReadingVolunteerNumber() {
        const digits = document.getElementById('readingNumberDigits').value.trim();
        document.getElementById('readingVolunteerNumber').value = digits ? readingBranch + '-' + digits : '';
    }
    document.getElementById('readingNumberDigits').addEventListener('input', syncReadingVolunteerNumber);
    document.getElementById('readingBranchBtn').addEventListener('click', syncReadingVolunteerNumber);
    document.getElementById('readingNumberDigits').closest('form').addEventListener('submit', syncReadingVolunteerNumber);
    function selectAllNames() {
        const boxes = document.querySelectorAll('input[name="names"]');
        boxes.forEach(b => b.checked = true);
    }
    </script>

    </body>
    </html>
    """

    return render_template_string(
        html,
        t=t,
        attendees=attendees,
        today_records=today_records,
        today_summary=today_summary_records,
        topic_options=topic_options
    )

    
@reading_bp.route("/reading_delete/<int:record_id>")
def reading_delete(record_id):
    db_query("""
        delete from reading
        where id = %s
    """, (record_id,))

    return redirect(url_for("reading.reading"))

@reading_bp.route("/reading_edit/<int:record_id>", methods=["GET", "POST"])
def reading_edit(record_id):
    row = db_query("""
        select *
        from reading
        where id = %s
    """, (record_id,), fetchone=True)

    if not row:
        return "找不到这笔记录<br><a href='/reading'>返回</a>"

    if request.method == "POST":
        new_topic = request.form.get("topic", "").strip()
        new_session = row.get("session") or ""

        db_query("""
            update reading
            set topic = %s,
                session = %s
            where id = %s
        """, (new_topic, new_session, record_id))

        return redirect(url_for("reading.reading"))

    html = """
    <!doctype html>
    <html lang="zh">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">

        <title>修改白话佛法记录</title>

        <link rel="stylesheet"
            href="{{ url_for('static', filename='css/toolbox.css') }}">
    </head>

    <body>

    <div class="page">

        <div class="card">

            <h1 class="page-title">
                ✏️ 修改白话佛法记录
            </h1>

            <div class="summary-grid">

                <div class="summary-box">
                    <div class="summary-title">
                        姓名
                    </div>

                    <div class="summary-value">
                        {{ name }}
                    </div>
                </div>

                <div class="summary-box">
                    <div class="summary-title">
                        身份
                    </div>

                    <div class="summary-value">
                        {{ identity }}
                    </div>
                </div>

            </div>

            <form method="post">

                <div class="form-group">

                    <label class="form-label">
                        共修主题
                    </label>

                    <input
                        class="form-input"
                        name="topic"
                        value="{{ topic }}">

                </div>



                <div class="btn-row">

                    <button
                        class="btn-tool btn-success"
                        type="submit">

                        保存修改

                    </button>

                </div>

            </form>

            <div class="btn-row">

                <a
                    class="btn-tool btn-secondary"
                    href="{{ url_for('reading.reading') }}">

                    返回共修记录

                </a>

            </div>

        </div>

    </div>

    </body>
    </html>
    """

    return render_template_string(
        html,
        name=row.get("name"),
        identity=row.get("identity"),
        topic=row.get("topic") or "",
        session=row.get("session") or ""
    )