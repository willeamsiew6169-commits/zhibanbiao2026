# volunteer_templates.py

VOLUNTEER_SIGNUP_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>蕉赖观音堂义工报名</title>

<style>
body {
    font-family:"Microsoft YaHei", Arial;
    background:#f8f3ea;
    padding:20px;
    font-size:24px;
}

.box {
    background:white;
    max-width:1200px;
    margin:auto;
    padding:35px;
    border-radius:20px;
    box-shadow:0 4px 15px rgba(0,0,0,0.08);
}

.hero {
    text-align:center;
    margin-bottom:25px;
}

.hero .lotus {
    font-size:60px;
}

.hero h1 {
    color:#8b5a2b;
    margin:10px 0;
    font-size:42px;
}

.hero p {
    color:#666;
    line-height:1.8;
    font-size:22px;
}

.top-actions {
    display:flex;
    justify-content:center;
    gap:20px;
    flex-wrap:wrap;
    margin-bottom:20px;
}

small {
    color:#777;
    font-size:18px;
    font-weight:normal;
}


.top-btn {
    display:flex;
    align-items:center;
    justify-content:center;
    width:280px;
    min-height:100px;
    border-radius:18px;
    color:white;
    text-decoration:none;
    font-size:28px;
    font-weight:bold;
    text-align:center;
    padding:10px;
    box-sizing:border-box;
}

.green  { background:#4CAF50; }
.blue   { background:#5B8DEF; }
.orange { background:#c58b39; }
.yellow { background:#F4D35E; }

.notice {
    background:#fff8e6;
    color:#8b5a2b;
    border:2px solid #f3d9a5;
    padding:18px;
    border-radius:15px;
    margin:0 auto 25px auto;
    max-width:900px;
    font-size:22px;
    line-height:1.8;
    text-align:center;
}

.admin {
    text-align:right;
    font-size:18px;
    margin-bottom:20px;
}

.signup-card {
    background:#fafafa;
    border:2px solid #eee;
    border-radius:20px;
    padding:25px;
}

label {
    font-weight:bold;
    color:#5c3b1e;
}

input, select, button {
    font-size:24px;
    padding:12px;
    margin:8px 0 18px 0;
    width:100%;
    box-sizing:border-box;
}

input, select {
    border:1px solid #d6c7b0;
    border-radius:12px;
    background:white;
}

button {
    cursor:pointer;
    background:#4CAF50;
    color:white;
    border:0;
    border-radius:15px;
    font-weight:bold;
    padding:20px;
    font-size:28px;
}

@media (max-width:700px) {
    body {
        padding:10px;
        font-size:22px;
    }

    .box {
        padding:20px;
    }

    .hero h1 {
        font-size:34px;
    }

    .top-btn {
        width:100%;
        min-height:85px;
        font-size:24px;
    }
}
</style>
</head>

<body>
<div class="box">

<div class="hero">
    <div class="lotus">🪷</div>
    <h1>蕉赖观音堂义工报名</h1>
    <p>
        感恩师兄们发心护持观音堂 🙏<br>
        随缘报名，共同护持道场
    </p>
</div>

<div class="top-actions">
    <a class="top-btn green" href="/volunteer/today_schedule">
        📋 {{ schedule_label }}
    </a>

    <a class="top-btn blue" href="/volunteer/prebook">
        📅 多日报名
    </a>

    <a class="top-btn yellow" href="/volunteer/monthly_signup_list?year={{ prebook_year }}&month={{ prebook_month }}">
        📖 查看预报名名单
    </a>
    
    <a class="top-btn orange" href="/volunteer/my_schedule_search">
        🔍 我的报名
    </a>
</div>

<div class="notice">
    🌸 报名后将由负责人统一安排岗位。<br>
    📢 请以最终公布的值班表为准。<br>
    🙏 感恩您的发心与护持。
</div>

<div class="admin">
    <a href="/schedule/admin">🔐 管理员入口</a>
</div>

<div class="signup-card">
<form method="post" action="/volunteer/signup">

<label>
义工编号 / 电话 / 姓名<br>
<small>Volunteer ID / Phone / Name</small>
</label>
<input name="keyword" required placeholder="例如 CHE-108 / 108 / 姓名">

<label>
日期<br>
<small>Date</small>
</label>
<input type="date" name="signup_date" value="{{ default_date }}" required>

<label>
岗位<br>
<small>Duty Type</small>
</label>
<select name="role" id="role_select" required onchange="toggleTimeSection()">
    <option value="值班">值班 Duty</option>
    <option value="卫生">卫生 Cleaning</option>
    <option value="供台">供台 Offering Table</option>
    </select>

<div id="time_section">

<label>
开始时间<br>
<small>Start Time</small>
</label>
<select
    name="start_time"
    id="start_time"
    onchange="updateTimeOptions()"
>
{% for t in times %}
<option value="{{ t }}">{{ t }}</option>
{% endfor %}
</select>

<label>
结束时间<br>
<small>End Time</small>
</label>
<select
    name="end_time"
    id="end_time"
    onchange="updateTimeOptions()"
>
{% for t in times %}
<option value="{{ t }}">{{ t }}</option>
{% endfor %}
</select>

</div>

<button type="submit">
🙏 提交报名<br>
<small>Submit Signup</small>
</button>

</form>
</div>

</div>

<script>
const ALL_TIMES = {{ times|tojson }};

function parseTimeToMinutes(t) {
    t = String(t).trim().toLowerCase();

    let match = t.match(/^(\d{1,2}):(\d{2})(am|pm)$/);
    if (!match) return null;

    let hour = parseInt(match[1]);
    let minute = parseInt(match[2]);
    let ap = match[3];

    if (ap === "pm" && hour !== 12) hour += 12;
    if (ap === "am" && hour === 12) hour = 0;

    return hour * 60 + minute;
}

function getCurrentRoundedTimeMinutes() {
    const now = new Date();
    let m = now.getHours() * 60 + now.getMinutes();

    // 向下取最近半小时：12:05 -> 12:00，12:35 -> 12:30
    return Math.floor(m / 30) * 30;
}

function isSignupDateToday() {
    const dateInput = document.querySelector("input[name='signup_date']");
    if (!dateInput) return false;

    const selected = dateInput.value;

    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const dd = String(today.getDate()).padStart(2, "0");

    return selected === `${yyyy}-${mm}-${dd}`;
}

function toggleTimeSection() {
    const role = document.getElementById("role_select").value;
    const timeSection = document.getElementById("time_section");

    if (role === "卫生" || role === "供台") {
        timeSection.style.display = "none";
    } else {
        timeSection.style.display = "block";
        updateTimeOptions();
    }
}

function updateTimeOptions() {
    const role = document.getElementById("role_select").value;

    if (role !== "值班") return;

    const today = isSignupDateToday();
    const nowMin = getCurrentRoundedTimeMinutes();
    const dateInput = document.querySelector("input[name='signup_date']");

    let dutyStartMin = 600; // 默认平时 10:00am

    if (dateInput && dateInput.value) {
        fetch("/volunteer/day_info?date=" + dateInput.value)
            .then(r => r.json())
            .then(data => {
                if (data.duty_start) {
                    dutyStartMin = parseTimeToMinutes(data.duty_start);
                }
                applyTimeFilter(today, nowMin, dutyStartMin);
            });
    } else {
        applyTimeFilter(today, nowMin, dutyStartMin);
    }
}

function applyTimeFilter(today, nowMin, dutyStartMin) {
    const start = document.getElementById("start_time");
    const end = document.getElementById("end_time");

    if (!start || !end) return;

    for (let i = 0; i < start.options.length; i++) {
        const tMin = parseTimeToMinutes(start.options[i].value);

        let hide = false;

        if (tMin < dutyStartMin) hide = true;
        if (today && tMin < nowMin) hide = true;

        start.options[i].hidden = hide;
        start.options[i].disabled = hide;
    }

    // 只有当前选择已经失效时才重选

    if (start.options[start.selectedIndex]?.disabled) {

        for (let i = 0; i < start.options.length; i++) {

            if (!start.options[i].disabled) {
                start.selectedIndex = i;
                break;
            }
        }
    }

    const startMin = parseTimeToMinutes(start.value);

    for (let i = 0; i < end.options.length; i++) {
        const eMin = parseTimeToMinutes(end.options[i].value);

        let hide = false;

        if (eMin <= startMin) hide = true;

        end.options[i].hidden = hide;
        end.options[i].disabled = hide;
    }

    if (end.options[end.selectedIndex]?.disabled) {

        for (let i = 0; i < end.options.length; i++) {

            if (!end.options[i].disabled) {
                end.selectedIndex = i;
                break;
            }
        }
    }

    }

document.addEventListener("DOMContentLoaded", function () {
    toggleTimeSection();

    const roleSelect = document.getElementById("role_select");
    const dateInput = document.querySelector("input[name='signup_date']");
    const startSelect = document.getElementById("start_time");
    const endSelect = document.getElementById("end_time");

    if (roleSelect) {
        roleSelect.addEventListener("change", toggleTimeSection);
    }

    if (dateInput) {
        dateInput.addEventListener("change", function () {
            const startSelect = document.getElementById("start_time");
            const endSelect = document.getElementById("end_time");

            if (startSelect) startSelect.selectedIndex = 0;
            if (endSelect) endSelect.selectedIndex = 0;

            updateTimeOptions();
        });
    }

    if (startSelect) {
        startSelect.addEventListener("change", updateTimeOptions);
    }

    if (endSelect) {
        endSelect.addEventListener("change", updateTimeOptions);
    }

    updateTimeOptions();
});
</script>

</body>
</html>
"""

VOLUNTEER_PREBOOK_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>多日报名</title>

<style>
body {
    font-family:"Microsoft YaHei", Arial;
    background:#f5f5f5;
    padding:20px;
    font-size:24px;
}

.box {
    background:white;
    max-width:750px;
    margin:auto;
    padding:25px;
    border-radius:15px;
}

input, select, button {
    font-size:24px;
    padding:12px;
    margin:8px 0;
    width:100%;
    box-sizing:border-box;
}

.calendar {
    display:grid;
    grid-template-columns:repeat(7, 1fr);
    gap:10px;
    margin:15px 0;
}

.week-title {
    text-align:center;
    font-weight:bold;
    font-size:20px;
    color:#555;
}

.day-card {
    display:block;
    text-align:center;
    border:2px solid #ccc;
    border-radius:12px;
    padding:14px 0;
    font-size:24px;
    background:white;
    cursor:pointer;
}

.day-card input {
    display:none;
}

.day-card span {
    display:block;
}

.day-card.lunar {
    background:#fff2a8;
    border-color:#e6c84f;
}

.day-card.festival {
    background:#ffd1d1;
    border-color:#e58a8a;
}

.day-card.checked {
    background:#2196F3;
    color:white;
    border-color:#2196F3;
    font-weight:bold;
}

.empty-day {
    height:55px;
}

.legend {
    display:flex;
    gap:20px;
    margin:10px 0 20px 0;
    font-size:20px;
    flex-wrap:wrap;
}

.legend-box {
    display:inline-block;
    width:25px;
    height:18px;
    border-radius:4px;
    margin-right:6px;
    vertical-align:middle;
}

.legend-lunar {
    background:#fff2a8;
    border:1px solid #e6c84f;
}

.legend-festival {
    background:#ffd1d1;
    border:1px solid #e58a8a;
}

a {
    font-size:22px;
}
</style>

<script>
const SPECIAL_DAYS = {{ special_days|tojson }};

function updateCalendar() {
    const year = parseInt(document.getElementById("year").value);
    const month = parseInt(document.getElementById("month").value);
    const calendarDays = document.getElementById("calendar-days");

    calendarDays.innerHTML = "";

    if (!year || !month) return;

    const firstDay = new Date(year, month - 1, 1).getDay();
    const lastDate = new Date(year, month, 0).getDate();

    for (let i = 0; i < firstDay; i++) {
        const empty = document.createElement("div");
        empty.className = "empty-day";
        calendarDays.appendChild(empty);
    }

    for (let d = 1; d <= lastDate; d++) {
        const label = document.createElement("label");
        label.className = "day-card";

        if (SPECIAL_DAYS[d] === "lunar") {
            label.classList.add("lunar");
        }

        if (SPECIAL_DAYS[d] === "festival") {
            label.classList.add("festival");
        }

        const input = document.createElement("input");
        input.type = "checkbox";
        input.name = "days";
        input.value = d;

        const span = document.createElement("span");
        span.innerText = d;

        input.addEventListener("change", function () {
            if (input.checked) {
                label.classList.add("checked");
            } else {
                label.classList.remove("checked");
            }
        });

        label.appendChild(input);
        label.appendChild(span);
        calendarDays.appendChild(label);
    }
}

function reloadPrebookPage() {
    const year = document.getElementById("year").value;
    const month = document.getElementById("month").value;

    window.location.href = "/volunteer/prebook?year=" + year + "&month=" + month;
}

function toggleTimeFields() {
    const role = document.querySelector("select[name='role']").value;

    const startGroup = document.getElementById("start-time-group");
    const endGroup = document.getElementById("end-time-group");

    if (role === "卫生" || role === "供台" || role === "整理佛台") {
        startGroup.style.display = "none";
        endGroup.style.display = "none";
    } else {
        startGroup.style.display = "block";
        endGroup.style.display = "block";
        updateEndTimes();
    }
}

function updateEndTimes() {
    const start = document.getElementById("start_time");
    const end = document.getElementById("end_time");

    if (!start || !end) return;

    const startIndex = start.selectedIndex;

    for (let i = 0; i < end.options.length; i++) {
        if (i <= startIndex) {
            end.options[i].hidden = true;
            end.options[i].disabled = true;
        } else {
            end.options[i].hidden = false;
            end.options[i].disabled = false;
        }
    }

    if (end.selectedIndex <= startIndex) {
        end.selectedIndex = startIndex + 1;
    }
}

function goMonthlySignupList() {
    const year = document.getElementById("year").value;
    const month = document.getElementById("month").value;

    window.location.href = "/volunteer/monthly_signup_list?year=" + year + "&month=" + month;
}

document.addEventListener("DOMContentLoaded", function () {
    updateCalendar();
    toggleTimeFields();

    document.getElementById("year").addEventListener("change", reloadPrebookPage);
    document.getElementById("month").addEventListener("change", reloadPrebookPage);

    const roleSelect = document.querySelector("select[name='role']");
    const startSelect = document.getElementById("start_time");

    if (roleSelect) {
        roleSelect.addEventListener("change", toggleTimeFields);
    }

    if (startSelect) {
        startSelect.addEventListener("change", updateEndTimes);
    }
});
</script>

</head>

<body>
<div class="box">

<h1>📅 多日报名</h1>

<form method="post" action="/volunteer/prebook">

<label>义工编号 / 电话 / 姓名</label>
<input name="keyword" required placeholder="例如 CHE-108 / 108 / 姓名">

<label>年份</label>
<input id="year" name="year" value="{{ default_year }}" required>

<label>月份</label>
<select id="month" name="month">
{% for m in range(1, 13) %}
<option value="{{ m }}" {% if m == default_month %}selected{% endif %}>{{ m }}月</option>
{% endfor %}
</select>

<h3>选择日期</h3>

<div class="calendar">
    <div class="week-title">日</div>
    <div class="week-title">一</div>
    <div class="week-title">二</div>
    <div class="week-title">三</div>
    <div class="week-title">四</div>
    <div class="week-title">五</div>
    <div class="week-title">六</div>
</div>

<div id="calendar-days" class="calendar"></div>

<div class="legend">
    <span><span class="legend-box legend-lunar"></span>农历初一 / 十五</span>
    <span><span class="legend-box legend-festival"></span>佛诞日</span>
</div>

<label>岗位</label>
<select name="role" onchange="toggleTimeFields()">
    <option value="值班">值班</option>
    <option value="卫生">卫生</option>
    <option value="供台">供台</option>
</select>

<div id="start-time-group">
    <label>开始时间（值班才需要）</label>
    <select name="start_time" id="start_time" onchange="updateEndTimes()">
        {% for t in times %}
        <option value="{{ t }}">{{ t }}</option>
        {% endfor %}
    </select>
</div>

<div id="end-time-group">
    <label>结束时间（值班才需要）</label>
    <select name="end_time" id="end_time">
        {% for t in times %}
        <option value="{{ t }}">{{ t }}</option>
        {% endfor %}
    </select>
</div>

<button type="submit">提交多日报名</button>

<button type="button" onclick="goMonthlySignupList()">
    📖 查看这个月份预报名名单 / 复制 WhatsApp
</button>

</form>

<br>
<a href="/volunteer">返回</a>

</div>
</body>
</html>
"""