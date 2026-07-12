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
    display:flex;
    flex-direction:column;
    gap:22px;
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

            {% if food_offering_open %}

            <a
                class="btn-tool btn-teal"
                href="/volunteer/food-offering">

                <span class="action-main-text">
                    🥗 素食结缘报名
                </span>

                <span class="action-sub-text">
                    填写食物名称及份量
                </span>

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

                <!-- 日期 -->

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


                <!-- 报名项目 -->

                <div class="form-group">

                    <label class="form-label">
                        报名项目
                        <span class="form-help">Signup Type</span>
                    </label>

                    <input
                        type="hidden"
                        name="role"
                        id="role_select"
                        value="值班">

                    <div class="choice-grid">

                        <button
                            type="button"
                            class="choice-btn active"
                            data-role="值班"
                            onclick="selectRole(this)">

                            <span class="choice-icon">🤝</span>
                            <span class="choice-label">值班</span>
                            <span class="choice-help">Duty</span>

                        </button>

                        <button
                            type="button"
                            class="choice-btn"
                            data-role="卫生"
                            onclick="selectRole(this)">

                            <span class="choice-icon">🧹</span>
                            <span class="choice-label">卫生</span>
                            <span class="choice-help">Cleaning</span>

                        </button>

                        {% if supply_signup_available %}

                        <button
                            type="button"
                            class="choice-btn"
                            data-role="供台"
                            onclick="selectRole(this)">

                            <span class="choice-icon">🛕</span>
                            <span class="choice-label">供台</span>
                            <span class="choice-help">Offering Table</span>

                        </button>

                        {% endif %}

                        {% if meal_signup_open %}

                        <button
                            type="button"
                            class="choice-btn"
                            data-role="膳食"
                            onclick="selectRole(this)">

                            <span class="choice-icon">🍱</span>
                            <span class="choice-label">膳食组</span>
                            <span class="choice-help">Meal Team</span>

                        </button>

                        {% endif %}

                        {# 以后开启即可 #}

                        
                    </div>

                </div>


                <!-- 动态展开区域 -->

                <div
                    id="role-detail-area">

                    <!--
                    以后这里自动展开：

                    🛕 供台
                        ○ 设供台
                        ○ 收供台

                    🍱 膳食组
                        ○ 派餐义工
                        ○ 组长

                    🥗 素食结缘
                        分类报名
                    -->

                </div>

                <div
                    id="supply_task_section"
                    class="form-group"
                    style="display:none;">

                    <label class="form-label">
                        供台工作
                        <span class="form-help">Offering Task</span>
                    </label>

                    <input
                        type="hidden"
                        name="supply_task"
                        id="supply_task"
                        value="">

                    <div class="choice-grid">

                        {% if supply_setup_open %}

                        <button
                            type="button"
                            class="choice-btn supply-task-btn"
                            data-supply-task="setup"
                            onclick="selectSupplyTask(this)">

                            <span class="choice-icon">🛕</span>
                            <span class="choice-label">设供台</span>

                            <span class="choice-help">
                                {{ supply_summary.setup.count }}
                                /
                                {{ supply_summary.setup.limit }}
                            </span>

                        </button>

                        {% endif %}


                        {% if supply_remove_open %}

                        <button
                            type="button"
                            class="choice-btn supply-task-btn"
                            data-supply-task="remove"
                            onclick="selectSupplyTask(this)">

                            <span class="choice-icon">📦</span>
                            <span class="choice-label">收供台</span>

                            <span class="choice-help">
                                {{ supply_summary.remove.count }}
                                /
                                {{ supply_summary.remove.limit }}
                            </span>

                        </button>

                        {% endif %}

                    </div>

                    <div class="form-help" style="margin-top:12px;">
                        设供台：大日子早上布置供台<br>
                        收供台：供奉结束后收供品及整理供台
                    </div>

                </div>

                <div id="food_offering_task_section" class="form-group" style="display:none;">

                    <div class="section-title">
                        请选择结缘项目
                    </div>

                    <input
                        type="hidden"
                        name="food_offering_task"
                        id="food_offering_task_input">

                    <div class="choice-grid">

                        <button
                            type="button"
                            class="choice-btn food-offering-task-btn"
                            data-food-offering-task="main_food"
                            onclick="selectFoodOfferingTask(this)">

                            <span class="choice-icon">🍚</span>
                            <span class="choice-label">主食</span>
                            <span class="choice-help">Main Food</span>
                        </button>

                        <button
                            type="button"
                            class="choice-btn food-offering-task-btn"
                            data-food-offering-task="vegetable"
                            onclick="selectFoodOfferingTask(this)">

                            <span class="choice-icon">🥬</span>
                            <span class="choice-label">菜肴</span>
                            <span class="choice-help">Dishes</span>
                        </button>

                        <button
                            type="button"
                            class="choice-btn food-offering-task-btn"
                            data-food-offering-task="dessert"
                            onclick="selectFoodOfferingTask(this)">

                            <span class="choice-icon">🍰</span>
                            <span class="choice-label">甜品</span>
                            <span class="choice-help">Dessert</span>
                        </button>

                        <button
                            type="button"
                            class="choice-btn food-offering-task-btn"
                            data-food-offering-task="fruit"
                            onclick="selectFoodOfferingTask(this)">

                            <span class="choice-icon">🍉</span>
                            <span class="choice-label">水果</span>
                            <span class="choice-help">Fruit</span>
                        </button>

                        <button
                            type="button"
                            class="choice-btn food-offering-task-btn"
                            data-food-offering-task="drink"
                            onclick="selectFoodOfferingTask(this)">

                            <span class="choice-icon">🥤</span>
                            <span class="choice-label">饮料</span>
                            <span class="choice-help">Drinks</span>
                        </button>

                    </div>

                    <div
                        id="food_offering_item_section"
                        class="form-group"
                        style="display:none; margin-top:18px;">

                        <label class="form-label">
                            结缘食物名称
                            <span class="form-help">Food Name</span>
                        </label>

                        <input
                            type="text"
                            class="form-input"
                            name="food_offering_item"
                            id="food_offering_item"
                            maxlength="100"
                            placeholder="例如：炒饭、西瓜、豆浆、红豆汤">

                        <div class="form-help">
                            请写明准备结缘的食物或饮料。
                        </div>

                    </div>

                </div>


                <!-- 时间 -->

                <!-- 只有值班才显示的时间区域 -->
                <div id="duty_section" class="time-grid">

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

                <!-- 所有报名项目都需要这个按钮，所以放在 duty_section 外面 -->
                <div class="notice-card">
                    <button
                        type="button"
                        class="btn-tool btn-green btn-full"
                        onclick="openSignupConfirm(event)">
                        🙏 提交报名
                    </button>
                </div>

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
const ALL_TIMES = {{ times | tojson }};

let signupFormSubmitting = false;
let currentVolunteerText = "";
let lookupTimer = null;


/* =========================================================
   时间工具
   ========================================================= */

function parseTimeToMinutes(value) {
    const text = String(value || "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, "");

    const match = text.match(/^(\d{1,2}):(\d{2})(am|pm)$/);

    if (!match) {
        return null;
    }

    let hour = parseInt(match[1], 10);
    const minute = parseInt(match[2], 10);
    const period = match[3];

    if (period === "pm" && hour !== 12) {
        hour += 12;
    }

    if (period === "am" && hour === 12) {
        hour = 0;
    }

    return hour * 60 + minute;
}


function getCurrentRoundedTimeMinutes() {
    const now = new Date();
    const minutes = now.getHours() * 60 + now.getMinutes();

    // 向下取最近半小时
    return Math.floor(minutes / 30) * 30;
}


function getMalaysiaTodayString() {
    const formatter = new Intl.DateTimeFormat("en-CA", {
        timeZone: "Asia/Kuala_Lumpur",
        year: "numeric",
        month: "2-digit",
        day: "2-digit"
    });

    return formatter.format(new Date());
}


function isSignupDateToday() {
    const dateInput =
        document.querySelector("input[name='signup_date']");

    if (!dateInput) {
        return false;
    }

    return dateInput.value === getMalaysiaTodayString();
}


/* =========================================================
   岗位选择
   ========================================================= */

function selectRole(button) {

    const role = button.dataset.role;

    // 保存岗位到隐藏字段
    const roleInput = document.getElementById("role_select");

    if (roleInput) {
        roleInput.value = role;
    }

    // 只处理最上层的报名项目按钮
    document
        .querySelectorAll('.choice-btn[data-role]')
        .forEach(function (btn) {
            btn.classList.remove("active");
            btn.classList.remove("selected");
        });

    button.classList.add("active");

    const dutySection =
        document.getElementById("duty_section");

    const supplySection =
        document.getElementById("supply_task_section");

    const foodOfferingSection =
    document.getElementById("food_offering_task_section");

    // 先全部隐藏
    if (dutySection) {
        dutySection.style.display = "none";
    }

    if (supplySection) {
        supplySection.style.display = "none";
    }

    if (foodOfferingSection) {
        foodOfferingSection.style.display = "none";
    }

    // 再根据岗位显示对应区域
    if (role === "值班") {

        if (dutySection) {
            dutySection.style.display = "grid";
        }

    } else if (role === "供台") {

        if (supplySection) {
            supplySection.style.display = "block";
        }

    } else if (role === "素食结缘") {

        if (foodOfferingSection) {
            foodOfferingSection.style.display = "block";
        }
    }

    // 离开供台时，清除供台细项
    if (role !== "供台") {

        const supplyInput =
            document.getElementById("supply_task");

        if (supplyInput) {
            supplyInput.value = "";
        }

        document
            .querySelectorAll(".supply-task-btn")
            .forEach(function (btn) {
                btn.classList.remove("active");
                btn.classList.remove("selected");
            });
    }

    // 离开素食结缘时，清除素食细项
    if (role !== "素食结缘") {

        const foodInput =
            document.getElementById(
                "food_offering_task_input"
            );

        if (foodInput) {
            foodInput.value = "";
        }

        const foodItemInput =
            document.getElementById(
                "food_offering_item"
            );

        if (foodItemInput) {
            foodItemInput.value = "";
        }

        const foodItemSection =
            document.getElementById(
                "food_offering_item_section"
            );

        if (foodItemSection) {
            foodItemSection.style.display = "none";
        }

        document
            .querySelectorAll(".food-offering-task-btn")
            .forEach(function (btn) {
                btn.classList.remove("active");
                btn.classList.remove("selected");
            });
    }
}

function selectFoodOfferingTask(button) {

    document
        .querySelectorAll(".food-offering-task-btn")
        .forEach(function (btn) {
            btn.classList.remove("active");
            btn.classList.remove("selected");
        });

    button.classList.add("active");
    button.classList.add("selected");

    const input =
        document.getElementById(
            "food_offering_task_input"
        );

    if (input) {
        input.value =
            button.dataset.foodOfferingTask;
    }

    const itemSection =
        document.getElementById(
            "food_offering_item_section"
        );

    if (itemSection) {
        itemSection.style.display = "block";
    }

    const itemInput =
        document.getElementById(
            "food_offering_item"
        );

    if (itemInput) {
        itemInput.focus();
    }
}


function updateRoleUI(selectedRole) {
    const dutySection =
        document.getElementById("duty_section");

    const supplyTaskSection =
        document.getElementById("supply_task_section");

    const supplyTaskInput =
        document.getElementById("supply_task");

    // 只有值班显示时间
    if (dutySection) {
        dutySection.style.display =
            selectedRole === "值班"
                ? "grid"
                : "none";
    }

    // 只有供台显示设供台 / 收供台
    if (supplyTaskSection) {
        supplyTaskSection.style.display =
            selectedRole === "供台"
                ? "block"
                : "none";
    }

    // 离开供台时清除旧选择
    if (selectedRole !== "供台") {
        if (supplyTaskInput) {
            supplyTaskInput.value = "";
        }

        document
            .querySelectorAll(".supply-task-btn")
            .forEach(function (button) {
                button.classList.remove("active");
            });
    }

    if (selectedRole === "值班") {
        updateTimeOptions();
    }
}

function selectSupplyTask(buttonElement) {
    document
        .querySelectorAll(".supply-task-btn")
        .forEach(function (button) {
            button.classList.remove("active");
        });

    buttonElement.classList.add("active");

    const task =
        buttonElement.dataset.supplyTask || "";

    const supplyTaskInput =
        document.getElementById("supply_task");

    if (supplyTaskInput) {
        supplyTaskInput.value = task;
    }
}

function openFoodOfferingSignup(event) {

    if (event) {
        event.preventDefault();
    }

    // 如果报名表目前是隐藏状态，先打开
    if (typeof showDailySignup === "function") {
        showDailySignup();
    }

    const foodButton = document.querySelector(
        '.choice-btn[data-role="素食结缘"]'
    );

    if (foodButton) {
        selectRole(foodButton);
    }

    const signupBox = document.getElementById(
        "daily_signup_box"
    );

    if (signupBox) {
        signupBox.scrollIntoView({
            behavior: "smooth",
            block: "start"
        });
    }
}


/* =========================================================
   值班时间
   ========================================================= */

function setDefaultDutyTimes() {
    const startSelect = document.getElementById("start_time");
    const endSelect = document.getElementById("end_time");

    if (startSelect) {
        startSelect.value = "10:00am";
    }

    if (endSelect) {
        endSelect.value = "6:00pm";
    }
}


function resetAllTimeOptions() {
    const startSelect = document.getElementById("start_time");
    const endSelect = document.getElementById("end_time");

    if (startSelect) {
        Array.from(startSelect.options).forEach(function (option) {
            option.hidden = false;
            option.disabled = false;
        });
    }

    if (endSelect) {
        Array.from(endSelect.options).forEach(function (option) {
            option.hidden = false;
            option.disabled = false;
        });
    }
}


function selectFirstAvailableOption(selectElement) {
    if (!selectElement) {
        return;
    }

    const currentOption =
        selectElement.options[selectElement.selectedIndex];

    if (currentOption && !currentOption.disabled) {
        return;
    }

    for (let i = 0; i < selectElement.options.length; i++) {
        if (!selectElement.options[i].disabled) {
            selectElement.selectedIndex = i;
            return;
        }
    }
}


async function updateTimeOptions() {
    const roleInput = document.getElementById("role_select");
    const selectedRole = roleInput ? roleInput.value : "值班";

    if (selectedRole !== "值班") {
        return;
    }

    const dateInput =
        document.querySelector("input[name='signup_date']");

    const selectedDate = dateInput ? dateInput.value : "";

    let dutyStartMinutes = 10 * 60; // 默认 10:00am

    if (selectedDate) {
        try {
            const response = await fetch(
                "/volunteer/day_info?date=" +
                encodeURIComponent(selectedDate)
            );

            const data = await response.json();

            if (data.duty_start) {
                const parsed =
                    parseTimeToMinutes(data.duty_start);

                if (parsed !== null) {
                    dutyStartMinutes = parsed;
                }
            }
        } catch (error) {
            console.warn(
                "读取当天值班时间失败，使用默认 10:00am。",
                error
            );
        }
    }

    applyTimeFilter(dutyStartMinutes);
}


function applyTimeFilter(dutyStartMinutes) {
    const startSelect = document.getElementById("start_time");
    const endSelect = document.getElementById("end_time");

    if (!startSelect || !endSelect) {
        return;
    }

    resetAllTimeOptions();

    const today = isSignupDateToday();
    const currentMinutes = getCurrentRoundedTimeMinutes();

    Array.from(startSelect.options).forEach(function (option) {
        const optionMinutes =
            parseTimeToMinutes(option.value);

        let shouldHide = false;

        if (
            optionMinutes !== null &&
            optionMinutes < dutyStartMinutes
        ) {
            shouldHide = true;
        }

        if (
            today &&
            optionMinutes !== null &&
            optionMinutes < currentMinutes
        ) {
            shouldHide = true;
        }

        option.hidden = shouldHide;
        option.disabled = shouldHide;
    });

    selectFirstAvailableOption(startSelect);
    updateEndTimeOptions();
}


function updateEndTimeOptions() {
    const startSelect = document.getElementById("start_time");
    const endSelect = document.getElementById("end_time");

    if (!startSelect || !endSelect) {
        return;
    }

    const startMinutes =
        parseTimeToMinutes(startSelect.value);

    Array.from(endSelect.options).forEach(function (option) {
        const endMinutes =
            parseTimeToMinutes(option.value);

        const shouldHide =
            startMinutes !== null &&
            endMinutes !== null &&
            endMinutes <= startMinutes;

        option.hidden = shouldHide;
        option.disabled = shouldHide;
    });

    selectFirstAvailableOption(endSelect);
}


/* =========================================================
   报名确认
   ========================================================= */

async function openSignupConfirm(event) {
    if (event) {
        event.preventDefault();
    }

    if (signupFormSubmitting) {
        return false;
    }

    const keywordInput = document.getElementById("keyword");
    const branchInput = document.getElementById("branch");
    const roleInput = document.getElementById("role_select");
    const dateInput =
        document.querySelector("input[name='signup_date']");

    const keyword = keywordInput
        ? keywordInput.value.trim()
        : "";

    const branch = branchInput
        ? branchInput.value
        : "CHE";

    const role = roleInput
        ? roleInput.value
        : "值班";

    const signupDate = dateInput
        ? dateInput.value
        : "";

    if (!keyword) {
        alert("请输入义工编号、姓名或电话。");
        keywordInput?.focus();
        return false;
    }

    if (!signupDate) {
        alert("请选择报名日期。");
        dateInput?.focus();
        return false;
    }

    let volunteerText =
        currentVolunteerText || keyword;

    try {
        const response = await fetch(
            "/volunteer/query_volunteer" +
            "?keyword=" + encodeURIComponent(keyword) +
            "&branch=" + encodeURIComponent(branch)
        );

        const data = await response.json();

        if (data.ok) {
            volunteerText =
                `${data.volunteer_id}　${data.name}`;

            currentVolunteerText = volunteerText;
        }
    } catch (error) {
        console.warn("查询义工资料失败：", error);
    }

    let confirmHtml = `
        <p>
            <b>义工：</b><br>
            ${volunteerText}
        </p>

        <p>
            <b>日期：</b><br>
            ${signupDate}
        </p>

        <p>
            <b>岗位：</b><br>
            ${role}
        </p>
    `;

    if (role === "供台") {
        const supplyTaskInput =
            document.getElementById("supply_task");

        const supplyTask =
            supplyTaskInput ? supplyTaskInput.value : "";

        if (!supplyTask) {
            alert("请选择设供台或收供台。");
            return false;
        }

        const supplyTaskLabel =
            supplyTask === "setup"
                ? "设供台"
                : "收供台";

        confirmHtml += `
            <p>
                <b>供台工作：</b><br>
                ${supplyTaskLabel}
            </p>
        `;
    }

    if (role === "值班") {
        const startSelect =
            document.getElementById("start_time");

        const endSelect =
            document.getElementById("end_time");

        const startTime =
            startSelect ? startSelect.value : "";

        const endTime =
            endSelect ? endSelect.value : "";

        if (
            !startTime ||
            !endTime ||
            parseTimeToMinutes(endTime) <=
            parseTimeToMinutes(startTime)
        ) {
            alert("请选择正确的开始时间和结束时间。");
            return false;
        }

        confirmHtml += `
            <p>
                <b>时间：</b><br>
                ${startTime} ～ ${endTime}
            </p>
        `;
    }

    const confirmContent =
        document.getElementById("confirm_content");

    const confirmModal =
        document.getElementById("confirm_modal");

    if (confirmContent) {
        confirmContent.innerHTML = confirmHtml;
    }

    if (confirmModal) {
        confirmModal.style.display = "flex";
    }

    return false;
}


function closeSignupConfirm() {
    const confirmModal =
        document.getElementById("confirm_modal");

    if (confirmModal) {
        confirmModal.style.display = "none";
    }
}


function submitSignupForm() {
    if (signupFormSubmitting) {
        return;
    }

    const form =
        document.getElementById("daily_signup_form");

    if (!form) {
        alert("找不到报名表单，请刷新页面后再试。");
        return;
    }

    signupFormSubmitting = true;
    form.submit();
}


/* =========================================================
   义工查询
   ========================================================= */

function lookupVolunteer() {
    clearTimeout(lookupTimer);

    lookupTimer = setTimeout(async function () {
        const keywordInput =
            document.getElementById("keyword");

        const branchInput =
            document.getElementById("branch");

        const resultBox =
            document.getElementById(
                "volunteer_lookup_result"
            );

        if (!keywordInput || !resultBox) {
            return;
        }

        const keyword = keywordInput.value.trim();
        const branch = branchInput
            ? branchInput.value
            : "CHE";

        currentVolunteerText = "";

        if (!keyword) {
            resultBox.innerHTML = "";
            return;
        }

        resultBox.innerHTML = "🔎 查询中...";

        try {
            const response = await fetch(
                "/volunteer/query_volunteer" +
                "?keyword=" + encodeURIComponent(keyword) +
                "&branch=" + encodeURIComponent(branch)
            );

            const data = await response.json();

            if (data.ok) {
                currentVolunteerText =
                    `${data.volunteer_id}　${data.name}`;

                resultBox.innerHTML = `
                    <span style="
                        color:green;
                        font-weight:bold;
                    ">
                        ✅ ${currentVolunteerText}
                    </span>
                `;
            } else {
                resultBox.innerHTML = `
                    <span style="
                        color:#c0392b;
                        font-weight:bold;
                    ">
                        ⚠️ ${data.message}
                    </span>
                `;
            }
        } catch (error) {
            resultBox.innerHTML = `
                <span style="
                    color:#c0392b;
                    font-weight:bold;
                ">
                    ⚠️ 查询失败
                </span>
            `;
        }
    }, 400);
}


/* =========================================================
   分会按钮
   ========================================================= */

function toggleBranch() {
    const button =
        document.getElementById("branch_btn");

    const branchInput =
        document.getElementById("branch");

    if (!button || !branchInput) {
        return;
    }

    if (branchInput.value === "CHE") {
        branchInput.value = "STW";
        button.innerText = "STW";
        button.style.background = "#dc3545";
    } else {
        branchInput.value = "CHE";
        button.innerText = "CHE";
        button.style.background = "#28a745";
    }

    currentVolunteerText = "";
    lookupVolunteer();
}


/* =========================================================
   展开每日报名
   ========================================================= */

function showDailySignup() {
    const signupBox =
        document.getElementById("daily_signup_box");

    if (!signupBox) {
        return;
    }

    signupBox.style.display = "block";

    setTimeout(function () {
        signupBox.scrollIntoView({
            behavior: "smooth",
            block: "start"
        });
    }, 100);
}


/* =========================================================
   页面初始化
   ========================================================= */

document.addEventListener("DOMContentLoaded", function () {
    const roleInput =
        document.getElementById("role_select");

    const dateInput =
        document.querySelector(
            "input[name='signup_date']"
        );

    const startSelect =
        document.getElementById("start_time");

    const keywordInput =
        document.getElementById("keyword");

    const selectedRole =
        roleInput?.value || "值班";

    setDefaultDutyTimes();
    updateRoleUI(selectedRole);

    if (dateInput) {
        dateInput.addEventListener(
            "change",
            function () {
                setDefaultDutyTimes();

                if (
                    document.getElementById(
                        "role_select"
                    )?.value === "值班"
                ) {
                    updateTimeOptions();
                }
            }
        );
    }

    if (startSelect) {
        startSelect.addEventListener(
            "change",
            updateEndTimeOptions
        );
    }

    if (keywordInput) {
        keywordInput.addEventListener(
            "input",
            lookupVolunteer
        );
    }
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

