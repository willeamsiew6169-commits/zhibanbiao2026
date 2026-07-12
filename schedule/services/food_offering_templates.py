# schedule/services/food_offering_templates.py

FOOD_OFFERING_SIGNUP_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>素食结缘报名</title>
<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">
<style>
.food-date-box{
    text-align:center;
    padding:18px;
    border-radius:18px;
    background:#fff8e7;
    margin-top:18px;
}
.food-date-main{
    font-size:24px;
    font-weight:800;
}
.food-date-sub{
    margin-top:8px;
    font-size:20px;
    font-weight:700;
    color:#7c3aed;
}
</style>
</head>
<body>

<div class="page">

<div class="card">

    <h1 class="page-title">
        🥗 素食结缘报名
    </h1>

    <div class="page-subtitle">
        请填写结缘食物名称及份量
    </div>

    <div class="food-date-box">

        <div class="food-date-main">
            📅 {{ display_date }}　
            {{ weekday_text }}
        </div>

        {% if lunar_text %}
        <div class="food-date-sub">
            🌙 {{ lunar_text }}
        </div>
        {% endif %}

        {% if festival_name %}
        <div class="food-date-sub">
            🪷 {{ festival_name }}
        </div>
        {% endif %}

    </div>

    <div class="alert alert-warning"
         style="margin-top:18px;">

        📌 请写明食物名称及份量，例如：<br>
        炒饭／浅盘、炸木薯／100个、
        六味糖水／一煲。<br><br>

        ⏰ 报名截止：
        <b>{{ deadline_text }}</b><br>

        🩷 所有结缘食物请于当天
        <b>10:30am 前</b>送到。

    </div>

    <div class="btn-row">

        <a
            class="btn-tool btn-purple"
            href="{{ url_for(
                'schedule.food_offering_status',
                date=offering_date
            ) }}">

            📋 查看结缘名单

        </a>

    </div>

</div>


<div class="card">

<form
    method="post"
    action="{{ url_for(
        'schedule.food_offering_signup'
    ) }}"
    onsubmit="return validateFoodOfferingForm();">

    <input
        type="hidden"
        name="offering_date"
        value="{{ offering_date }}">

    <div class="form-group">

        <label class="form-label">
            义工编号／姓名／电话
        </label>

        <div class="branch-search-row">

            <button
                type="button"
                id="food_branch_btn"
                class="branch-toggle-btn"
                onclick="toggleFoodBranch()">
                CHE
            </button>

            <input
                type="hidden"
                id="food_branch"
                name="branch"
                value="CHE">

            <input
                type="text"
                class="form-input"
                name="keyword"
                required
                placeholder="例如：108／张三／电话号码">

        </div>

    </div>


    <div class="form-group">

        <label class="form-label">
            结缘种类
        </label>

        <input
            type="hidden"
            name="category"
            id="food_category"
            value="">

        <div class="choice-grid">

        {% for key, item in categories.items() %}

            {% set summary =
                category_summary[key] %}

            <button
                type="button"
                class="choice-btn food-category-btn"
                data-category="{{ key }}"
                onclick="selectFoodCategory(this)"
                {% if summary.is_full %}
                disabled
                {% endif %}>

                <span class="choice-icon">
                    {{ item.icon }}
                </span>

                <span class="choice-label">
                    {{ item.label }}
                </span>

                <span class="choice-help">

                    {% if summary.is_full %}
                        已满
                    {% else %}
                        剩 {{ summary.remaining }} 位
                    {% endif %}

                </span>

            </button>

        {% endfor %}

        </div>

    </div>


    <div class="form-group">

        <label class="form-label">
            食物名称
        </label>

        <input
            type="text"
            class="form-input"
            name="food_name"
            id="food_name"
            required
            maxlength="100"
            placeholder="例如：炒饭、咖喱菜、六味糖水">

    </div>


    <div class="form-group">

        <label class="form-label">
            份量
        </label>

        <input
            type="text"
            class="form-input"
            name="quantity_text"
            id="quantity_text"
            required
            maxlength="100"
            placeholder="例如：浅盘、深盘、80个、一煲">

    </div>


    <button
        type="submit"
        class="btn-tool btn-success btn-full">

        🙏 提交素食结缘

    </button>

</form>


<div class="btn-row"
     style="margin-top:18px;">

    <a
        class="btn-tool btn-secondary"
        href="/volunteer">

        ← 返回义工首页

    </a>

</div>

</div>

</div>


<script>

function selectFoodCategory(button){

    if(button.disabled){
        return;
    }

    document
        .querySelectorAll(
            ".food-category-btn"
        )
        .forEach(function(btn){

            btn.classList.remove("active");
            btn.classList.remove("selected");

        });

    button.classList.add("active");
    button.classList.add("selected");

    document
        .getElementById("food_category")
        .value = button.dataset.category;
}


function validateFoodOfferingForm(){

    const category =
        document
            .getElementById(
                "food_category"
            )
            .value;

    const foodName =
        document
            .getElementById(
                "food_name"
            )
            .value
            .trim();

    const quantity =
        document
            .getElementById(
                "quantity_text"
            )
            .value
            .trim();

    if(!category){
        alert("请选择结缘种类。");
        return false;
    }

    if(!foodName){
        alert("请填写食物名称。");
        return false;
    }

    if(!quantity){
        alert("请填写食物份量。");
        return false;
    }

    return confirm(
        "确认提交素食结缘报名吗？\\n\\n" +
        "食物：" + foodName + "\\n" +
        "份量：" + quantity
    );
}


function toggleFoodBranch(){

    const input =
        document.getElementById(
            "food_branch"
        );

    const button =
        document.getElementById(
            "food_branch_btn"
        );

    if(input.value === "CHE"){

        input.value = "STW";
        button.textContent = "STW";

    }else{

        input.value = "CHE";
        button.textContent = "CHE";
    }
}

</script>

</body>
</html>
"""


FOOD_OFFERING_STATUS_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>素食结缘名单</title>
<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">
<style>
.offering-person{
    padding:15px 0;
    border-bottom:1px solid #e5e7eb;
}
.offering-name{
    font-size:22px;
    font-weight:800;
}
.offering-food{
    margin-top:5px;
    font-size:19px;
    color:#555;
}
.category-count{
    font-size:18px;
    font-weight:700;
    color:#666;
}
</style>
</head>
<body>

<div class="page">

<div class="card">

    <h1 class="page-title">
        🥗 素食结缘名单
    </h1>

    <div class="page-subtitle">
        📅 {{ display_date }}　
        {{ weekday_text }}
    </div>

    {% if lunar_text %}
    <div style="
        text-align:center;
        margin-top:8px;
        font-size:20px;
        font-weight:700;
        color:#7c3aed;
    ">
        🌙 {{ lunar_text }}
    </div>
    {% endif %}

    {% if festival_name %}
    <div style="
        text-align:center;
        margin-top:6px;
        font-size:20px;
        font-weight:700;
    ">
        🪷 {{ festival_name }}
    </div>
    {% endif %}

    {% if success %}
    <div class="alert alert-success"
         style="margin-top:18px;">

        🎉 素食结缘报名成功，
        感恩您的发心。

    </div>
    {% endif %}

    {% if duplicate_food %}
    <div class="alert alert-warning"
         style="margin-top:18px;">

        ⚠️ 提醒：已有义工登记相同食物
        「{{ duplicate_food }}」。

    </div>
    {% endif %}

    <div class="btn-row">

        {% if signup_open %}

        <a
            class="btn-tool btn-success"
            href="{{ url_for(
                'schedule.food_offering_page'
            ) }}">

            ➕ 我要结缘

        </a>

        {% endif %}


        {% if is_admin %}

        <a
            class="btn-tool btn-purple"
            href="{{ url_for(
                'schedule.food_offering_whatsapp',
                date=offering_date
            ) }}">

            📋 WhatsApp 通知

        </a>

        {% endif %}


        <a
            class="btn-tool btn-secondary"
            href="/volunteer">

            ← 返回义工首页

        </a>

    </div>

</div>


{% for key, item in categories.items() %}

    {% set records =
        grouped_records[key] %}

    {% set summary =
        category_summary[key] %}

    <div class="card">

        <h2 class="section-title">
            {{ item.icon }}
            {{ item.label }}
        </h2>

        <div class="category-count">

            已报名
            {{ summary.count }}／
            {{ summary.limit }}

            {% if summary.is_full %}
                · 🔴 已满
            {% else %}
                · 尚余
                {{ summary.remaining }}
                位
            {% endif %}

        </div>

        {% if records %}

            {% for row in records %}

            <div class="offering-person">

                <div class="offering-name">
                    {{ loop.index }}.
                    {{ row.volunteer_name }}
                </div>

                <div class="offering-food">
                    {{ row.food_name }}
                    （{{ row.quantity_text }}）
                </div>

                {% if is_admin %}

                <form
                    method="post"
                    action="{{ url_for(
                        'schedule.food_offering_cancel',
                        signup_id=row.id
                    ) }}"
                    style="margin-top:10px;"
                    onsubmit="
                        return confirm(
                            '确认取消这笔素食结缘报名吗？'
                        );
                    ">

                    <input
                        type="hidden"
                        name="date"
                        value="{{ offering_date }}">

                    <button
                        type="submit"
                        class="btn-tool btn-danger mini-btn">

                        取消

                    </button>

                </form>

                {% endif %}

            </div>

            {% endfor %}

        {% else %}

            <div class="empty-state">
                暂时没有报名
            </div>

        {% endif %}

    </div>

{% endfor %}

</div>

</body>
</html>
"""


FOOD_OFFERING_WHATSAPP_HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>素食结缘 WhatsApp</title>
<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">
</head>
<body>

<div class="page">

<div class="card">

    <h1 class="page-title">
        📋 素食结缘 WhatsApp
    </h1>

    <div class="page-subtitle">
        可直接复制后发送到 WhatsApp 群组
    </div>

    <textarea
        id="whatsapp_text"
        class="form-input"
        style="
            min-height:560px;
            line-height:1.7;
            white-space:pre-wrap;
        "
        readonly>{{ whatsapp_text }}</textarea>

    <div class="btn-row">

        <button
            type="button"
            class="btn-tool btn-success"
            onclick="copyWhatsAppText()">

            📋 复制 WhatsApp

        </button>

        <a
            class="btn-tool btn-secondary"
            href="{{ url_for(
                'schedule.food_offering_status',
                date=offering_date
            ) }}">

            ← 返回名单

        </a>

    </div>

</div>

</div>


<script>

async function copyWhatsAppText(){

    const text =
        document
            .getElementById(
                "whatsapp_text"
            )
            .value;

    try{

        await navigator.clipboard
            .writeText(text);

        alert(
            "✅ 已复制 WhatsApp 内容"
        );

    }catch(error){

        const textarea =
            document.getElementById(
                "whatsapp_text"
            );

        textarea.select();

        document.execCommand("copy");

        alert(
            "✅ 已复制 WhatsApp 内容"
        );
    }
}

</script>

</body>
</html>
"""
