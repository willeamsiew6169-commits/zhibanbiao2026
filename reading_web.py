# reading_web.py

import os
import re
import html as html_lib
import ipaddress
import socket
from html.parser import HTMLParser
from urllib.parse import urlparse, urljoin
from datetime import timedelta

import requests
import pandas as pd

from db import db_query, get_conn
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from flask import (
    Blueprint,
    request,
    redirect,
    url_for,
    render_template_string,
    abort
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


def get_reading_person_by_number(number):
    """按完整 CHE/STW 编号查询月费佛友或义工，保留原分会核对。"""
    number = str(number or "").strip().upper().replace(" ", "")
    match = re.fullmatch(r"(CHE|STW)-?(\d+)", number)
    if not match:
        return None
    branch, digits = match.groups()
    normalized = f"{branch}-{int(digits)}"

    # 月费佛友可能不在 volunteers 表中；优先使用 members 的正式资料。
    members = db_query("""
        select name from members
        where upper(member_id) = %s
        limit 2
    """, (normalized,), fetchall=True) or []
    member_names = {str(r.get("name") or "").strip() for r in members if r.get("name")}
    if len(member_names) == 1:
        return next(iter(member_names)), "member"
    if len(member_names) > 1:
        return None

    rows = db_query("""
        select id, name from volunteers
        where upper(id) = %s or upper(coalesce(member_id, '')) = %s
        limit 2
    """, (normalized, normalized), fetchall=True) or []
    names = {str(r.get("name") or "").strip() for r in rows if r.get("name")}
    return (next(iter(names)), "volunteer") if len(names) == 1 else None


def reading_number_error_page(t):
    """仅用于编号找不到时的友善提示；不更改共修保存流程。"""
    is_en = str(t.get("html_lang", "zh")).startswith("en")
    return render_template_string("""
    <!doctype html><html lang="{{ 'en' if is_en else 'zh' }}"><head>
    <meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <title>{{ 'Number not found' if is_en else '编号查询提示' }}</title>
    <style>
      *{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;
      padding:20px;background:#f5f7f3;color:#24372a;font-family:'Microsoft YaHei',Arial,sans-serif}
      .card{width:min(100%,520px);padding:36px 26px;text-align:center;background:white;
      border-radius:20px;box-shadow:0 12px 36px #15331d1c}
      .icon{font-size:54px;margin-bottom:12px}h1{font-size:27px;margin:0 0 18px}
      p{font-size:18px;line-height:1.8;color:#59685c;margin:0 0 26px}
      a{display:block;padding:15px 20px;background:#249451;color:white;text-decoration:none;
      border-radius:12px;font-size:19px;font-weight:bold}
      a:hover{background:#18773e}
      @media(max-width:480px){.card{padding:30px 18px}h1{font-size:23px}p{font-size:17px}}
    </style></head><body><main class="card">
      <div class="icon">⚠️</div>
      <h1>{{ 'Number not found' if is_en else '找不到此编号' }}</h1>
      <p>{{ 'Please check the CHE/STW branch and number, then try again.' if is_en else '请检查 CHE/STW 分会及编号是否正确，然后重新提交。' }}</p>
      <a href="/reading">{{ '← Back to study record' if is_en else '← 返回共修记录' }}</a>
    </main></body></html>
    """, is_en=is_en), 400


# 每日共修链接：独立于历史 reading 记录，不自动改写旧记录。
# 部署前请先执行随附的 reading_daily_topics.sql。
class _ReadingTitleParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_title = False
        self.title = ""
        self.og_title = ""
        self.visible_parts = []
        self._hidden_depth = 0

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag in ("script", "style", "noscript"):
            self._hidden_depth += 1
        if tag == "title":
            self.in_title = True
        if tag == "meta" and attrs.get("property", "").lower() == "og:title":
            self.og_title = attrs.get("content", "")

    def handle_endtag(self, tag):
        if tag == "title":
            self.in_title = False
        if tag in ("script", "style", "noscript"):
            self._hidden_depth = max(0, self._hidden_depth - 1)

    def handle_data(self, data):
        if self.in_title:
            self.title += data
        elif not self._hidden_depth and data.strip():
            self.visible_parts.append(data.strip())


def _is_google_form_url(url):
    parsed = urlparse(url)
    host = (parsed.hostname or "").lower()
    return (parsed.scheme == "https" and not parsed.username and not parsed.password
            and parsed.port in (None, 443) and host in ("forms.gle", "docs.google.com"))


def _normalize_reading_heading(value):
    """仅整理已识别标题的重复前缀；其他书籍标题保持原文。"""
    value = str(value).strip()
    bracket = re.search(r"[【\[]([^】\]]{5,100})[】\]]", value)
    if bracket:
        inner = bracket.group(1).strip()
        if ("白话佛法广播讲座" in inner and
                re.search(r"20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日", inner)):
            value = inner
    match = re.search(r"白话佛法广播讲座\s*[-－—：:]?\s*(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", value)
    if match:
        year, month, day = map(int, match.groups())
        return f"白话佛法广播讲座－{year}年{month}月{day}日"
    return value


def _extract_reading_heading(candidates):
    """从正文提取书名/篇章，排除表单提醒；不限制固定书名。"""
    for item in candidates:
        value = re.sub(r"\s+", " ", str(item)).strip(" ：:【】[]-—\t")
        # 表单通知、作答说明不能作为主题。
        if not value or len(value) > 120 or any(word in value for word in
                ("考试前", "考试后", "增强记忆", "填写", "提交", "请选择", "建议时间")):
            continue
        # 有书名并有明确的篇章/日期/编号，才认为可靠。
        if re.search(r"(?:白话佛法|广播讲座|视频开示)", value) and re.search(
                r"(?:20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\b\d{6}\b|第[一二三四五六七八九十百\d]+[册篇章集])", value):
            # 去掉材料标题前常见的说明性前缀。
            bracket = re.search(r"[【\[]([^】\]]{5,100})[】\]]", value)
            if bracket and re.search(r"(?:20\d{2}年|\d{6}|第.+[篇章集])", bracket.group(1)):
                value = bracket.group(1).strip()
            return _normalize_reading_heading(value[:200])
        # 换书后的新书名可能不含白话佛法：须有明确「第N册第N篇」等结构。
        if re.search(r"第[一二三四五六七八九十百\d]+册", value) and re.search(
                r"第[一二三四五六七八九十百\d]+[篇章]", value):
            return _normalize_reading_heading(value[:200])
    return None


def _read_google_form_title(url):
    """只允许 Google Forms 域名及受限重定向；失败时让义工手动填主题。"""
    current = url
    with requests.Session() as session:
        session.trust_env = False
        for _ in range(5):
            if not _is_google_form_url(current):
                raise ValueError("只接受 Google Forms 链接")
            # Google 官方域名可能在网络环境中解析异常；拒绝私网目标。
            host = urlparse(current).hostname
            for result in socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM):
                address = ipaddress.ip_address(result[4][0])
                if not address.is_global:
                    raise ValueError("链接地址不可访问")
            with session.get(current, timeout=(4, 7), allow_redirects=False,
                             stream=True, headers={"User-Agent": "Mozilla/5.0"}) as response:
                if response.status_code in (301, 302, 303, 307, 308):
                    location = response.headers.get("Location")
                    if not location:
                        raise ValueError("链接跳转无效")
                    current = urljoin(current, location)
                    continue
                response.raise_for_status()
                if "text/html" not in response.headers.get("Content-Type", "").lower():
                    raise ValueError("不是可读取的表单页面")
                chunks, total = [], 0
                for chunk in response.iter_content(chunk_size=8192):
                    total += len(chunk)
                    if total > 250000:
                        break
                    chunks.append(chunk)
                page = b"".join(chunks).decode("utf-8", errors="replace")
                parser = _ReadingTitleParser()
                parser.feed(page)
                # 优先找正文中的阅读篇章标题；绝不将表单顶部提醒当成主题。
                candidates = [html_lib.unescape(x).strip() for x in parser.visible_parts]
                # Google Forms 部分正文放在内嵌数据中，尝试读取转义后的中文。
                embedded = page.replace('\\u003c', '<').replace('\\u003e', '>')
                try:
                    embedded = re.sub(r'\\u([0-9a-fA-F]{4})',
                                      lambda m: chr(int(m.group(1), 16)), embedded)
                except (ValueError, OverflowError):
                    pass
                candidates += re.findall(r'[^\"\[\]<>\n]{0,85}(?:白话佛法|广播讲座|视频开示|第[一二三四五六七八九十百\d]+[册篇章集])[^\"\[\]<>\n]{0,85}', embedded)
                found = _extract_reading_heading(candidates)
                if not found:
                    raise ValueError("未能识别阅读材料的篇章标题")
                return found
    raise ValueError("链接跳转次数过多")


def _daily_reading_topic(day):
    return db_query("SELECT date, topic, url FROM reading_daily_topics WHERE date=%s",
                    (day,), fetchone=True)


@reading_bp.route("/reading/daily-link", methods=["POST"])
def submit_reading_daily_link():
    t = get_text()
    is_en = str(t.get("html_lang", "zh")).startswith("en")
    day_choice = request.form.get("day", "today")
    if day_choice not in ("today", "tomorrow"):
        abort(400)
    day = datetime.now(MY_TZ).date() + timedelta(days=1 if day_choice == "tomorrow" else 0)
    url = request.form.get("daily_url", "").strip()
    fallback_topic = request.form.get("manual_topic", "").strip()
    if len(url) > 2048 or not _is_google_form_url(url):
        return _daily_link_message("只接受有效的 Google Forms HTTPS 链接。" if not is_en else "Please enter a valid Google Forms HTTPS link.", False)
    existing = _daily_reading_topic(day)
    if existing:
        if existing["url"] == url:
            return _daily_link_message("该日期的链接已经提交，无需重复提交。" if not is_en else "This link was already submitted for that date.", True)
        return _daily_link_message("该日期已有不同链接，请联系负责人更正。" if not is_en else "A different link is already saved for this date. Please contact an administrator.", False)
    if request.form.get("stage") != "confirm":
        try:
            topic = _read_google_form_title(url)
        except (ValueError, requests.RequestException, OSError):
            topic = ""
        return _daily_link_preview(day_choice, url, topic or fallback_topic)
    topic = request.form.get("confirmed_topic", "").strip()
    if not topic:
        return _daily_link_message("请先确认或填写正确的共修主题。", False)
    if len(topic) > 200:
        return _daily_link_message("主题最多200字。" if not is_en else "Topic must be at most 200 characters.", False)
    # 并发提交只允许第一笔成功；不覆盖任何已有日期的链接。
    inserted = db_query("""
        INSERT INTO reading_daily_topics (date, topic, url)
        VALUES (%s, %s, %s)
        ON CONFLICT (date) DO NOTHING RETURNING date
    """, (day, topic, url), fetchone=True)
    if not inserted:
        return _daily_link_message("该日期已有链接，请刷新共修页面查看。" if not is_en else "A link was already saved. Refresh the study page.", False)
    return redirect(url_for("reading.reading"))


def _daily_link_preview(day_choice, url, topic):
    """先预览，义工确认后才写入数据库；允许更正识别结果。"""
    return render_template_string("""<!doctype html><html lang="zh"><head>
    <meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <style>*{box-sizing:border-box}body{font-family:'Microsoft YaHei',Arial,sans-serif;background:#f6f7f2;margin:0;padding:18px;color:#24372a}
    main{max-width:620px;margin:25px auto;background:#fff;border-radius:16px;padding:24px;box-shadow:0 3px 15px #0001}
    h2{font-size:24px}p{font-size:17px;line-height:1.6}label{display:block;font-size:18px;font-weight:bold;margin:18px 0 8px}
    input{width:100%;padding:14px;border:1px solid #b9cdbb;border-radius:10px;font-size:19px}
    button,.back{display:block;text-align:center;width:100%;padding:15px;border:0;border-radius:11px;background:#309b51;color:white;font-size:19px;font-weight:bold;text-decoration:none;margin-top:16px;cursor:pointer}
    .back{background:#697a70}.tip{color:#a15b0a;background:#fff7e7;padding:12px;border-radius:9px}</style></head>
    <body><main><h2>📖 确认共修主题</h2>
    <p>共修日期：<strong>{{ '今天' if day == 'today' else '明天' }}</strong></p>
    <p class="tip">{{ '请核对识别结果；如果不正确，可以直接修改。' if topic else '未能可靠识别阅读标题，请手动填写后确认。' }}</p>
    <form action="{{ url_for('reading.submit_reading_daily_link') }}" method="post">
    <input type="hidden" name="stage" value="confirm"><input type="hidden" name="day" value="{{ day }}">
    <input type="hidden" name="daily_url" value="{{ url }}">
    <label>今日共修主题</label><input name="confirmed_topic" maxlength="200" value="{{ topic }}" required autofocus>
    <button type="submit">✅ 确认并保存</button></form>
    <a class="back" href="/reading">← 返回修改链接</a></main></body></html>""", day=day_choice, url=url, topic=topic)


def _daily_link_message(message, success):
    return render_template_string("""<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1">
    <style>body{background:#f6f7f2;font-family:Arial,'Microsoft YaHei',sans-serif;padding:25px}
    main{max-width:520px;margin:12vh auto;background:white;padding:30px;border-radius:18px;text-align:center;box-shadow:0 6px 25px #0001}
    p{font-size:20px;line-height:1.7}a{display:block;background:#399d55;color:white;padding:15px;border-radius:10px;text-decoration:none;font-size:18px}</style>
    </head><body><main><h1>{{ '✅' if success else '⚠️' }}</h1><p>{{ message }}</p><a href="/reading">{{ '返回共修记录 / Back to study' }}</a></main></body></html>""", message=message, success=success), (200 if success else 400)


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
        daily = _daily_reading_topic(datetime.now(MY_TZ).date())
        if daily:
            topic = daily["topic"]
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
            person = get_reading_person_by_number(volunteer_number)
            if not person:
                return reading_number_error_page(t)
            person_name, person_type = person
            if person_type == "member":
                identity = t.get("friend_identity", "Friend" if t.get("html_lang", "zh").startswith("en") else "佛友")
            else:
                identity = t.get("volunteer_identity", "Volunteer" if t.get("html_lang", "zh").startswith("en") else "义工")
            unique_people[person_name] = identity
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
    daily_today = _daily_reading_topic(datetime.now(MY_TZ).date())
    daily_tomorrow = _daily_reading_topic(datetime.now(MY_TZ).date() + timedelta(days=1))

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
          .daily-grid{display:flex;gap:10px;flex-wrap:wrap;align-items:center}.daily-grid label{display:inline-flex;align-items:center;gap:6px;padding:9px 12px;background:#eef7ed;border-radius:9px}.daily-link{display:block;width:100%;max-width:460px;text-align:center;background:#2563c4;color:white;text-decoration:none;font-size:20px;font-weight:700;padding:17px 18px;border-radius:12px;margin:15px 0;overflow-wrap:anywhere}.daily-note{color:#66736a;font-size:15px;line-height:1.5}
          .name-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(145px,1fr));gap:9px;margin:12px 0}
          .name-item{display:flex;align-items:center;gap:8px;padding:12px;background:#f4f8f3;border-radius:10px;font-size:17px}
          .name-item input{width:19px;height:19px}
          button,.button-link{display:inline-flex;align-items:center;justify-content:center;gap:5px;border:0;border-radius:10px;padding:11px 16px;background:#399d55;color:white;font-size:16px;cursor:pointer;text-decoration:none}
          .button-link{margin:0 8px 8px 0}.secondary{background:#4778bb}.edit{background:#e6a23c}.delete{background:#d9534f}
          .actions{display:flex;gap:7px;flex-wrap:wrap}.actions button,.actions .button-link{font-size:14px;padding:8px 10px;margin:0}
          .record{border:1px solid #e2e9e2;border-radius:12px;padding:13px;margin:10px 0}
          .record-head{display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:8px}
          .record-meta{font-size:14px;color:#66736a;margin-bottom:10px;overflow-wrap:anywhere}
          .record-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(145px,1fr));gap:9px;margin-top:12px}
          .record-grid .record{margin:0;padding:0;background:#f4f8f3;min-width:0;overflow:hidden}
          .record-grid .record summary{padding:11px 9px;font-size:17px;font-weight:700;list-style:none;text-align:center;overflow-wrap:anywhere}
          .record-grid .record summary::-webkit-details-marker{display:none}
          .record-grid .record[open]{grid-column:span 2;background:#fff}
          .record-grid .record[open] summary{background:#eef7ed}
          .record-grid .record-meta{padding:9px 11px 0;margin-bottom:7px}
          .record-grid .actions{padding:0 10px 10px}
          .record-grid .actions .button-link{flex:1;font-size:13px;padding:7px 5px}
          .muted{color:#718074;font-size:15px}
          details summary{cursor:pointer;font-size:19px;font-weight:700;padding:5px 0}
          @media(max-width:600px){.record-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.record-grid .record[open]{grid-column:span 2}body{padding:9px;font-size:16px}.card{padding:15px;border-radius:13px}h2{font-size:20px}.name-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.name-item{padding:10px;font-size:16px}.record-head{align-items:flex-start}.actions{width:100%}.actions .button-link{flex:1}}
        </style>
    </head>
    <body><main class="wrap">
      <nav>
        <a class="button-link" href="/">⬅ {{ t.back_home }}</a>
        <a class="button-link secondary" href="/download_reading">📥 {{ t.download_report or ('下载共修报表' if not t.html_lang.startswith('en') else 'Download report') }}</a>
      </nav>
      <section class="card">
        <h2>{{ '📖 每日共修链接' if not t.html_lang.startswith('en') else '📖 Daily study link' }}</h2>
        {% if daily_today %}
          <p><strong>{{ '今日主题：' if not t.html_lang.startswith('en') else 'Today: ' }}</strong>{{ daily_today.topic }}</p>
          <a class="daily-link" href="{{ daily_today.url }}" target="_blank" rel="noopener noreferrer">📖 {{ '打开今日共修内容' if not t.html_lang.startswith('en') else 'Open today’s form' }}</a>
        {% else %}<p class="daily-note">{{ '今天尚未提交共修链接。' if not t.html_lang.startswith('en') else 'No link submitted for today.' }}</p>{% endif %}
        {% if daily_tomorrow %}<p class="daily-note">{{ '明日主题：' if not t.html_lang.startswith('en') else 'Tomorrow: ' }}{{ daily_tomorrow.topic }}</p>{% endif %}
        <form action="{{ url_for('reading.submit_reading_daily_link') }}" method="post">
          <div class="daily-grid">
            <label><input type="radio" name="day" value="today" checked> {{ '今天' if not t.html_lang.startswith('en') else 'Today' }}</label>
            <label><input type="radio" name="day" value="tomorrow"> {{ '明天' if not t.html_lang.startswith('en') else 'Tomorrow' }}</label>
          </div>
          <label class="field-label">{{ '粘贴 Google Forms 共修链接' if not t.html_lang.startswith('en') else 'Paste Google Forms link' }}</label>
          <input type="url" name="daily_url" maxlength="2048" placeholder="https://forms.gle/..." required style="width:100%;font-size:17px;padding:13px;border:1px solid #cbd5cc;border-radius:10px">
          <details style="margin:12px 0"><summary style="font-size:15px">{{ '可选：预先填写主题' if not t.html_lang.startswith('en') else 'Fallback: enter topic if detection fails' }}</summary>
            <input type="text" name="manual_topic" maxlength="200" placeholder="{{ '例如：白话佛法广播讲座-170131' if not t.html_lang.startswith('en') else 'Study topic' }}">
          </details>
          <button type="submit">{{ '识别并预览主题' if not t.html_lang.startswith('en') else 'Submit study link' }}</button>
          <p class="daily-note">{{ '无需编号或 PIN；同一天已有不同链接时不能直接覆盖。' if not t.html_lang.startswith('en') else 'No ID or PIN required. Existing links cannot be overwritten.' }}</p>
        </form>
      </section>
      <section class="card">
        <h2>{{ t.bhff_record_title }}</h2>
        <form method="post">
          <label class="field-label">{{ t.today_topic }}</label>
          <input type="text" name="topic" list="topicList" value="{{ daily_today.topic if daily_today else '' }}" placeholder="{{ t.topic_placeholder }}" {% if daily_today %}readonly{% else %}required{% endif %}>
          {% if daily_today %}<p class="daily-note">{{ '已自动采用今日主题。' if not t.html_lang.startswith('en') else 'Today’s topic is filled automatically.' }}</p>{% endif %}
          <datalist id="topicList">{% for topic in topic_options %}<option value="{{ topic }}">{% endfor %}</datalist>
          <label class="field-label">{{ t.volunteer_list }}（{{ attendees|length }}）</label>
          <button type="button" onclick="selectAllNames()">{{ t.select_all_volunteers }}</button>
          {% if attendees %}
          <div class="name-grid">
          {% for name in attendees %}<label class="name-item"><input type="checkbox" name="names" value="{{ name }}">{{ name }}</label>{% endfor %}
          </div>
          {% else %}<p class="muted">今日签到义工均已点名，或暂无可点名义工。</p>{% endif %}
          <label class="field-label">{{ '未签到义工／月费佛友：输入 CHE/STW 编号' if not t.html_lang.startswith('en') else 'Volunteer / monthly member: enter CHE/STW number' }}</label>
          <div class="branch-number-row">
            <button type="button" id="readingBranchBtn" class="branch-switch che" onclick="toggleReadingBranch()" aria-label="切换 CHE 或 STW">CHE</button>
            <input type="text" id="readingNumberDigits" inputmode="numeric" autocomplete="off" pattern="[0-9]+" placeholder="例如：208" title="只需输入数字编号">
          </div>
          <input type="hidden" name="volunteer_number" id="readingVolunteerNumber">
          <p class="muted">{{ '不必填写姓名；保存时自动从义工或月费佛友资料读取姓名。' if not t.html_lang.startswith('en') else 'No name needed; the name is looked up from volunteer or member records when saving.' }}</p>
          <label class="field-label">{{ t.extra_friend_name }}</label>
          <input type="text" name="extra_names" placeholder="{{ t.extra_friend_placeholder }}">
          <p class="muted">{{ t.extra_friend_tip }}</p>
          <button type="submit">{{ t.record_study }}</button>
        </form>
      </section>
      <section class="card">
        <details>
          <summary>📋 {{ t.today_recorded }}（{{ today_records|length }}）</summary>
          <div class="record-grid">
          {% for r in today_records %}
          <details class="record">
            <summary title="{{ r['身份'] }} · {{ r['主题'] }}">{{ r["姓名"] }}</summary>
            <div class="record-meta">{{ r["时间"] }}<br>{{ r["身份"] }} · {{ r["主题"] }}</div>
            <div class="actions">
              <a class="button-link edit" href="/reading_edit/{{ r['id'] }}">{{ t.edit }}</a>
              <a class="button-link delete" href="/reading_delete/{{ r['id'] }}" onclick="return confirm({{ t.get("delete_confirm_simple", "确定删除这笔共修记录吗？" if not t.get("html_lang", "zh").startswith("en") else "Delete this study record?")|tojson }})">{{ t.delete_record }}</a>
            </div>
          </details>
          {% else %}<p class="muted">今日暂无记录。</p>{% endfor %}
          </div>
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
        topic_options=topic_options,
        daily_today=daily_today,
        daily_tomorrow=daily_tomorrow
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