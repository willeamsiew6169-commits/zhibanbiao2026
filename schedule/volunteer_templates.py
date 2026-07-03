# volunteer_templates.py

VOLUNTEER_SIGNUP_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="manifest" href="/volunteer-manifest.json">
<link rel="icon" href="/static/volunteer_icon.png?v=1">

<title>蕉赖共修会义工报名</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
/* =========================
   Volunteer Signup V4
   只保留本页专用样式，其余交给 toolbox.css
========================= */

.volunteer-wrap{
    max-width:980px;
    margin:auto;
    padding:20px;
}

.volunteer-hero{
    text-align:center;
    padding:24px 18px 12px;
}

.volunteer-lotus{
    font-size:58px;
    line-height:1;
    margin-bottom:10px;
}

.volunteer-title{
    font-size:46px;
    font-weight:800;
    margin:8px 0;
    color:#8b5a2b;
}

.volunteer-subtitle{
    font-size:26px;
    color:#6b7280;
    line-height:1.7;
    margin:10px 0 0;
}

.volunteer-actions{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:18px;
    margin:24px 0;
}

.volunteer-actions .btn-tool{
    min-height:92px;
    flex-direction:column;
    line-height:1.35;
    text-align:center;
}

.action-main-text{
    font-size:30px;
    font-weight:bold;
}

.action-sub-text{
    font-size:21px;
    font-weight:bold;
    opacity:.95;
    margin-top:6px;
}

.signup-section{
    display:none;
    margin-top:24px;
}

.branch-search-row{
    display:grid;
    grid-template-columns:120px 1fr;
    gap:14px;
    align-items:center;
}

.branch-toggle-btn{
    height:72px;
    font-size:28px;
    font-weight:bold;
    background:#28a745;
    color:white;
    border:none;
    border-radius:16px;
    cursor:pointer;
}

.signup-form-grid{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:20px;
}

.time-grid{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:20px;
}

.lookup-result{
    font-size:24px;
    margin-top:10px;
    min-height:32px;
}

.form-card-title{
    font-size:36px;
    font-weight:bold;
    text-align:center;
    margin:0 0 18px;
}

.form-card-desc{
    font-size:24px;
    text-align:center;
    color:#6b7280;
    line-height:1.6;
    margin-bottom:26px;
}

.admin-link{
    text-align:right;
    font-size:22px;
    margin:16px 0 0;
}

.confirm-overlay{
    position:fixed;
    inset:0;
    background:rgba(0,0,0,.45);
    display:flex;
    justify-content:center;
    align-items:center;
    z-index:9999;
    padding:18px;
}

.confirm-box{
    background:white;
    width:100%;
    max-width:560px;
    border-radius:22px;
    padding:26px;
    box-shadow:0 8px 25px rgba(0,0,0,.25);
}

.confirm-box h2{
    margin-top:0;
    color:#8b5a2b;
    text-align:center;
    font-size:34px;
}

.confirm-content{
    font-size:26px;
    line-height:1.7;
    color:#333;
}

.confirm-actions{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:14px;
    margin-top:24px;
}

#keyword::placeholder{
    color:#b79b6b;
}

#keyword:focus{
    background:#fff6d8;
    border-color:#d6a100;
    outline:none;
    box-shadow:0 0 0 4px rgba(214,161,0,.18);
}

@media (max-width:700px){
    .volunteer-wrap{
        padding:12px;
    }

    .volunteer-title{
        font-size:36px;
    }

    .volunteer-subtitle{
        font-size:22px;
    }

    .volunteer-actions,
    .signup-form-grid,
    .time-grid,
    .branch-search-row,
    .confirm-actions{
        grid-template-columns:1fr;
    }

    .volunteer-actions .btn-tool{
        min-height:82px;
    }

    .action-main-text{
        font-size:26px;
    }

    .action-sub-text{
        font-size:19px;
    }
}
</style>
</head>

<body>
<div class="volunteer-wrap">

    <div class="card">

        <div class="volunteer-hero">
            <div class="volunteer-lotus">🪷</div>
            <div class="volunteer-title">蕉赖共修会义工报名</div>
            <p class="volunteer-subtitle">
                感恩师兄们发心护持观音堂 🙏<br>
                随缘报名，共同护持道场
            </p>
        </div>

        <div class="volunteer-actions">

            <button
                type="button"
                class="btn-tool btn-green"
                onclick="showDailySignup()">
                <span class="action-main-text">📝 每日报名</span>
                <span class="action-sub-text">今天 / 明天值班报名</span>
            </button>

            <a
                class="btn-tool btn-purple"
                href="/volunteer/my_schedule_search">
                <span class="action-main-text">📋 我的报名</span>
                <span class="action-sub-text">查看 / 修改 / 取消</span>
            </a>

            {% if multi_day_signup_open %}
            <a
                class="btn-tool btn-blue"
                href="/volunteer/prebook">
                <span class="action-main-text">📅 下个月多日报名</span>
                <span class="action-sub-text">一次报名多个日期</span>
            </a>
            {% endif %}

            {% if meal_signup_open %}
            <a
                class="btn-tool btn-red"
                href="/volunteer/meal_status?date={{ meal_date }}">
                <span class="action-main-text">🍱 派餐义工名单</span>
                <span class="action-sub-text">{{ meal_button_date }}　👥 {{ meal_count }}/9</span>
            </a>
            {% endif %}

        </div>

        <div class="alert alert-warning">
            🌸 报名后将由系统安排岗位，必要时负责人会调整。<br>
            📢 10:00pm 正式发布前，可自行修改或取消。<br>
            🚫 值班当天如需更改或取消，请必须通知负责人。<br>
            🙏 请以最新正式值班表为准，感恩护持。
        </div>

        <div class="btn-row" style="margin-top:18px;">
            <a class="btn-tool btn-blue" href="/volunteer/guide">
                📖 义工须知
            </a>
        </div>

        <div class="admin-link">
            <a href="/schedule/admin">🔐 管理员入口</a>
        </div>

    </div>

    <div
        id="daily_signup_box"
        class="card signup-section">

        <h2 class="form-card-title">📝 每日报名</h2>
        <div class="form-card-desc">
            请填写义工资料、日期、岗位和时间。提交前系统会让您再次确认。
        </div>

        <form
            id="daily_signup_form"
            method="post"
            action="/volunteer/signup">

            <div class="form-group">
                <label class="form-label">
                    义工编号 / 姓名 / 电话
                </label>

                <div class="branch-search-row">

                    <button
                        type="button"
                        id="branch_btn"
                        onclick="toggleBranch()"
                        class="branch-toggle-btn">
                        CHE
                    </button>

                    <input
                        type="hidden"
                        id="branch"
                        name="branch"
                        value="CHE">

                    <input
                        class="form-input"
                        name="keyword"
                        id="keyword"
                        required
                        placeholder="例如：108 / 张三 / 0123456789"
                        oninput="lookupVolunteer()">

                </div>

                <div id="volunteer_lookup_result" class="lookup-result"></div>
            </div>

            <div class="signup-form-grid">

                <div class="form-group">
                    <label class="form-label">
                        日期
                        <span class="form-help">Date</span>
                    </label>

                    <input
                        class="form-input"
                        type="date"
                        name="signup_date"
                        value="{{ default_date }}"
                        required>
                </div>

                <div class="form-group">
                    <label class="form-label">
                        岗位
                        <span class="form-help">Duty Type</span>
                    </label>

                    <select
                        class="form-select"
                        name="role"
                        id="role_select"
                        required
                        onchange="toggleTimeSection()">

                        <option value="值班">值班 Duty</option>
                        <option value="卫生">卫生 Cleaning</option>
                        <option value="供台">供台 Offering Table</option>
                        <option value="膳食">膳食组 Meal Team</option>

                    </select>
                </div>

            </div>

            <div id="time_section" class="time-grid">

                <div class="form-group">
                    <label class="form-label">
                        开始时间
                        <span class="form-help">Start Time</span>
                    </label>

                    <select
                        class="form-select"
                        name="start_time"
                        id="start_time"
                        onchange="updateTimeOptions()">

                        {% for t in times %}
                        <option value="{{ t }}">{{ t }}</option>
                        {% endfor %}

                    </select>
                </div>

                <div class="form-group">
                    <label class="form-label">
                        结束时间
                        <span class="form-help">End Time</span>
                    </label>

                    <select
                        class="form-select"
                        name="end_time"
                        id="end_time"
                        onchange="updateTimeOptions()">

                        {% for t in times %}
                        <option value="{{ t }}">{{ t }}</option>
                        {% endfor %}

                    </select>
                </div>

            </div>

            <div class="btn-row">
                <button
                    type="button"
                    class="btn-tool btn-green btn-full"
                    onclick="openSignupConfirm(event)">
                    🙏 提交报名
                </button>
            </div>

        </form>

        <div id="confirm_modal" class="confirm-overlay" style="display:none;">
            <div class="confirm-box">
                <h2>📋 确认报名资料</h2>

                <div id="confirm_content" class="confirm-content"></div>

                <div class="confirm-actions">
                    <button
                        type="button"
                        class="btn-tool btn-gray"
                        onclick="closeSignupConfirm()">
                        取消
                    </button>

                    <button
                        type="button"
                        class="btn-tool btn-green"
                        onclick="submitSignupForm()">
                        确认报名
                    </button>
                </div>
            </div>
        </div>

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

    if (role === "卫生" || role === "供台" || role === "膳食") {
        timeSection.style.display = "none";
    } else {
        timeSection.style.display = "grid";
        updateTimeOptions();
    }
}

window.addEventListener("DOMContentLoaded", toggleTimeSection);

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

            if (startSelect) startSelect.value = "10:00am";
            if (endSelect) endSelect.value = "6:00pm";

            updateTimeOptions();
        });
    }

    if (startSelect) {
        startSelect.addEventListener("change", updateTimeOptions);
    }

    if (endSelect) {
        endSelect.addEventListener("change", updateTimeOptions);
    }

    if (startSelect) startSelect.value = "10:00am";
    if (endSelect) endSelect.value = "6:00pm";

    updateTimeOptions();
});

let signupFormSubmitting = false;
let currentVolunteerText = "";

async function openSignupConfirm(event){
    if(event){
        event.preventDefault();
    }

    if(signupFormSubmitting){
        return false;
    }

    const keyword = document.getElementById("keyword").value.trim();
    const branch = document.getElementById("branch").value;
    const role = document.getElementById("role_select").value;
    const date = document.querySelector("input[name='signup_date']").value;

    let volunteerText = keyword;

    try {
        const r = await fetch(`/volunteer/query_volunteer?keyword=${encodeURIComponent(keyword)}&branch=${encodeURIComponent(branch)}`);
        const data = await r.json();

        if(data.ok){
            volunteerText = `${data.volunteer_id}　${data.name}`;
            currentVolunteerText = volunteerText;
        }
    } catch(e) {}

    let html = `
        <p><b>义工：</b><br>${volunteerText}</p>
        <p><b>日期：</b><br>${date}</p>
        <p><b>岗位：</b><br>${role}</p>
    `;

    if(role === "值班"){
        const start = document.getElementById("start_time").value;
        const end = document.getElementById("end_time").value;

        html += `
            <p><b>时间：</b><br>${start} ～ ${end}</p>
        `;
    }

    document.getElementById("confirm_content").innerHTML = html;
    document.getElementById("confirm_modal").style.display = "flex";

    return false;
}

function closeSignupConfirm(){
    document.getElementById("confirm_modal").style.display = "none";
}

function submitSignupForm(){
    signupFormSubmitting = true;
    document.getElementById("daily_signup_form").submit();
}
</script>

<script>
function toggleBranch(){
    const btn=document.getElementById("branch_btn");
    const branch=document.getElementById("branch");

    if(branch.value==="CHE"){
        branch.value="STW";
        btn.innerText="STW";
        btn.style.background="#dc3545";
    }else{
        branch.value="CHE";
        btn.innerText="CHE";
        btn.style.background="#28a745";
    }
}

function showDailySignup(){
    const box = document.getElementById("daily_signup_box");

    box.style.display = "block";

    setTimeout(function(){
        box.scrollIntoView({
            behavior:"smooth",
            block:"start"
        });
    },100);
}

let lookupTimer = null;

function lookupVolunteer(){
    clearTimeout(lookupTimer);

    lookupTimer = setTimeout(function(){
        const keyword = document.getElementById("keyword").value.trim();
        const branch = document.getElementById("branch").value;
        const result = document.getElementById("volunteer_lookup_result");

        currentVolunteerText = "";

        if(!keyword){
            result.innerHTML = "";
            return;
        }

        result.innerHTML = "🔎 查询中...";

        fetch(`/volunteer/query_volunteer?keyword=${encodeURIComponent(keyword)}&branch=${encodeURIComponent(branch)}`)
            .then(r => r.json())
            .then(data => {
                if(data.ok){
                    currentVolunteerText = `${data.volunteer_id}　${data.name}`;
                    result.innerHTML = `<span style="color:green;font-weight:bold;">✅ ${currentVolunteerText}</span>`;
                }else{
                    result.innerHTML = `<span style="color:#c0392b;font-weight:bold;">⚠️ ${data.message}</span>`;
                }
            })
            .catch(() => {
                result.innerHTML = `<span style="color:#c0392b;font-weight:bold;">⚠️ 查询失败</span>`;
            });
    }, 400);
}
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

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.prebook-wrap{
    max-width:900px;
    margin:auto;
    padding:20px;
}

.prebook-card{
    margin-top:20px;
}

.calendar{
    display:grid;
    grid-template-columns:repeat(7, 1fr);
    gap:10px;
    margin:15px 0;
}

.week-title{
    text-align:center;
    font-weight:bold;
    font-size:22px;
    color:#555;
}

.day-card{
    display:block;
    text-align:center;
    border:2px solid #d1d5db;
    border-radius:14px;
    padding:16px 0;
    font-size:28px;
    background:white;
    cursor:pointer;
}

.day-card input{
    display:none;
}

.day-card span{
    display:block;
}

.day-card.lunar{
    background:#fff2a8;
    border-color:#e6c84f;
}

.day-card.festival{
    background:#ffd1d1;
    border-color:#e58a8a;
}

.day-card.checked{
    background:#2563eb;
    color:white;
    border-color:#2563eb;
    font-weight:bold;
}

.empty-day{
    height:60px;
}

.legend{
    display:flex;
    gap:20px;
    margin:15px 0 25px;
    font-size:22px;
    flex-wrap:wrap;
}

.legend-box{
    display:inline-block;
    width:28px;
    height:20px;
    border-radius:6px;
    margin-right:8px;
    vertical-align:middle;
}

.legend-lunar{
    background:#fff2a8;
    border:1px solid #e6c84f;
}

.legend-festival{
    background:#ffd1d1;
    border:1px solid #e58a8a;
}

.branch-search-row{
    display:grid;
    grid-template-columns:110px 1fr;
    gap:12px;
    align-items:center;
}

.branch-toggle-btn{
    height:66px;
    font-size:26px;
    font-weight:bold;
    color:white;
    border:none;
    border-radius:16px;
    cursor:pointer;
}

.month-grid{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:18px;
}

@media (max-width:700px){
    .month-grid,
    .branch-search-row{
        grid-template-columns:1fr;
    }

    .day-card{
        font-size:24px;
        padding:14px 0;
    }
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
    const keyword = document.getElementById("keyword").value;
    const branch = document.getElementById("multi_branch").value;

    window.location.href =
        "/volunteer/prebook?"
        + "year=" + encodeURIComponent(year)
        + "&month=" + encodeURIComponent(month)
        + "&keyword=" + encodeURIComponent(keyword)
        + "&branch=" + encodeURIComponent(branch);
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

    window.location.href =
        "/volunteer/monthly_signup_list?year=" + year + "&month=" + month;
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

<div class="prebook-wrap">

<div class="card prebook-card">

    <h1 class="page-title">📅 多日报名</h1>

    <p class="page-subtitle">
        可一次报名整个月份，减少每天重复填写。
    </p>

    <form method="post" action="/volunteer/prebook">

        <div class="form-group">
            <label class="form-label">
                义工编号 / 姓名 / 电话
            </label>

            <div class="branch-search-row">

                <button
                    type="button"
                    id="multi_branch_btn"
                    onclick="toggleMultiBranch()"
                    class="branch-toggle-btn"
                    style="background:{{ '#dc3545' if branch == 'STW' else '#28a745' }};">
                    {{ branch }}
                </button>

                <input
                    type="hidden"
                    id="multi_branch"
                    name="branch"
                    value="{{ branch }}">

                <input
                    class="form-input"
                    name="keyword"
                    id="keyword"
                    value="{{ keyword }}"
                    required
                    placeholder="例如：108 / 张三 / 0123456789">

            </div>
        </div>

        <div class="month-grid">

            <div class="form-group">
                <label class="form-label">年份</label>

                <input
                    class="form-input"
                    id="year"
                    name="year"
                    value="{{ default_year }}"
                    required>
            </div>

            <div class="form-group">
                <label class="form-label">月份</label>

                <select
                    class="form-select"
                    id="month"
                    name="month">

                    {% for m in range(1, 13) %}
                    <option value="{{ m }}" {% if m == default_month %}selected{% endif %}>
                        {{ m }}月
                    </option>
                    {% endfor %}

                </select>
            </div>

        </div>

        <div class="divider"></div>

        <h2 class="section-title">选择日期</h2>

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
            <span>
                <span class="legend-box legend-lunar"></span>
                农历初一 / 十五
            </span>

            <span>
                <span class="legend-box legend-festival"></span>
                佛诞日
            </span>
        </div>

        <div class="form-group">
            <label class="form-label">岗位</label>

            <select
                class="form-select"
                name="role"
                onchange="toggleTimeFields()">

                <option value="值班">值班</option>
                <option value="卫生">卫生</option>
                <option value="供台">供台</option>

            </select>
        </div>

        <div class="form-grid">

            <div class="form-group" id="start-time-group">
                <label class="form-label">
                    开始时间
                    <span class="form-help">值班才需要</span>
                </label>

                <select
                    class="form-select"
                    name="start_time"
                    id="start_time"
                    onchange="updateEndTimes()">

                    {% for t in times %}
                    <option value="{{ t }}">{{ t }}</option>
                    {% endfor %}

                </select>
            </div>

            <div class="form-group" id="end-time-group">
                <label class="form-label">
                    结束时间
                    <span class="form-help">值班才需要</span>
                </label>

                <select
                    class="form-select"
                    name="end_time"
                    id="end_time">

                    {% for t in times %}
                    <option value="{{ t }}">{{ t }}</option>
                    {% endfor %}

                </select>
            </div>

        </div>

        <div class="btn-row">

            <button
                class="btn-tool btn-green btn-full"
                type="submit">
                ✅ 提交多日报名
            </button>

            <button
                class="btn-tool btn-blue btn-full"
                type="button"
                onclick="goMonthlySignupList()">
                📖 查看本月预报名 / 复制 WhatsApp
            </button>

            <a
                class="btn-tool btn-gray btn-full"
                href="/volunteer">
                ⬅ 返回义工首页
            </a>

        </div>

    </form>

</div>

</div>

<script>
function toggleMultiBranch(){

    const btn = document.getElementById("multi_branch_btn");
    const branch = document.getElementById("multi_branch");

    if(branch.value === "CHE"){
        branch.value = "STW";
        btn.innerText = "STW";
        btn.style.background = "#dc3545";
    }else{
        branch.value = "CHE";
        btn.innerText = "CHE";
        btn.style.background = "#28a745";
    }
}
</script>

</body>
</html>
"""