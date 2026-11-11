"""从 schedule_web.py 原样迁移的页面 HTML 常量（不修改内容）。"""

_MANAGEMENT_MENU_HTML = """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{ title }}</title>
<style>
body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#f2f5f9;margin:0;padding:24px;color:#172b4d}
main{max-width:780px;margin:35px auto;background:white;padding:30px;border-radius:18px;box-shadow:0 5px 22px #172b4d12}
h1{font-size:30px;margin:0 0 26px}.links{display:grid;gap:16px}
.links a{display:block;text-align:center;padding:19px 14px;border-radius:13px;color:white;text-decoration:none;font-size:23px;font-weight:bold;background:#237d76}
.links a:nth-child(2){background:#596879}.links a:nth-child(3){background:#146e86}
.back{display:inline-block;margin-top:25px;color:#275de4;font-size:19px;text-decoration:none}
</style></head><body><main><h1>{{ title }}</h1><div class="links">
{% for item in links %}<a href="{{ item.url }}">{{ item.label }}</a>{% endfor %}
</div><a class="back" href="/schedule#management-functions">⬅ 返回负责人首页</a>
</main></body></html>"""


LOGIN_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>负责人排班系统</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.login-wrap{
    max-width:620px;
    margin:60px auto;
    padding:20px;
}

.login-card{
    text-align:center;
}

.login-icon{
    font-size:64px;
    margin-bottom:10px;
}

.login-title{
    font-size:52px;
    font-weight:bold;
    color:#1e293b;
}

.login-subtitle{
    margin:15px 0 35px;
    color:#666;
    font-size:22px;
    line-height:1.8;
}

.login-input{
    margin-top:10px;
}

.login-input input{
    text-align:center;
    letter-spacing:6px;
    font-size:34px;
}

@media (max-width:700px){

    .login-title{
        font-size:40px;
    }

    .login-subtitle{
        font-size:20px;
    }

}
</style>

<script>
window.addEventListener("DOMContentLoaded",function(){

    document.querySelector("input[name='pin']").focus();

});
</script>

</head>

<body>

<div class="login-wrap">

    <div class="card login-card">

        <div class="login-icon">
            📅
        </div>

        <div class="login-title">
            负责人排班系统
        </div>

        <div class="login-subtitle">
            请输入负责人 PIN<br>
            进入排班管理系统。
        </div>

        <form method="post">

            <div class="form-group">

                <label class="form-label">
                    负责人 PIN
                </label>

                <div class="login-input">

                    <input
                        class="form-input"
                        type="password"
                        name="pin"
                        placeholder="请输入负责人 PIN"
                        inputmode="numeric"
                        autocomplete="new-password"
                        autocorrect="off"
                        autocapitalize="off"
                        spellcheck="false"
                        readonly
                        onfocus="this.removeAttribute('readonly');">

                </div>

            </div>

            <div class="btn-row">

                <button
                    class="btn-tool btn-green btn-full"
                    type="submit">

                    🔓 进入负责人系统

                </button>

            </div>

        </form>

    </div>

</div>

</body>
</html>
"""


_FIXED_ROLES_HTML = """
<!doctype html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>固定岗位人员管理</title>
<style>
*{box-sizing:border-box}body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#f3f6fb;margin:0;padding:20px;color:#243047}
.wrap{max-width:960px;margin:auto}.card{background:white;padding:24px;margin:16px 0;border-radius:16px;box-shadow:0 3px 16px #22334a12}
h1{font-size:27px}h2{font-size:21px}input,select,button{font-size:16px;padding:11px;border-radius:9px;border:1px solid #ccd3df}
button{cursor:pointer}.primary{background:#187c71;color:white;border:0;padding:12px 20px}.danger{background:#ba3e43;color:white;border:0;padding:8px 14px}
.row{display:flex;flex-wrap:wrap;gap:12px;align-items:center}table{width:100%;border-collapse:collapse}td,th{padding:12px;border-bottom:1px solid #e5e9ef;text-align:left}
.notice{padding:12px;background:#e6f4ed;border-radius:8px}.error{background:#ffe7e7}a{color:#187c71}
.tabs,.branches{display:flex;gap:9px;flex-wrap:wrap}.tab,.branch{background:#edf1f7;color:#35435a;border:1px solid #dce3ed;font-weight:700}
.tab.active{background:#2663e9;color:white;border-color:#2663e9}.branch.active{background:#187c71;color:white;border-color:#187c71}
.role-panel{display:none}.role-panel.active{display:block}.muted{color:#65728a;font-size:14px}.search-box{position:relative;flex:1;min-width:220px}
.search-box input{width:100%}.results{position:absolute;z-index:10;top:100%;left:0;right:0;background:white;box-shadow:0 8px 20px #0002;border-radius:9px;max-height:260px;overflow:auto}
.result{display:block;width:100%;text-align:left;border:0;background:white;border-bottom:1px solid #edf0f4;border-radius:0}.result:hover{background:#edf6f4}
.selected{color:#137767;font-weight:700;min-height:24px}.weekday{padding:10px 0 2px;border-bottom:1px solid #e5e9ef}
@media(max-width:600px){body{padding:12px}.card{padding:16px}table{font-size:14px}.tab{flex:1}}
</style></head><body><div class="wrap">
<a href="/schedule/management/fixed">← 返回佛台及固定岗位管理</a>
<h1>👥 固定岗位人员管理</h1>
<p class="muted">管理供花、供果、佛学班及整理佛台每星期固定轮值；此页面不会直接修改已发布排班或签到记录。</p>
{% with messages = get_flashed_messages(with_categories=true) %}{% for category,msg in messages %}
<div class="notice {{ 'error' if category == 'error' else '' }}">{{ msg }}</div>
{% endfor %}{% endwith %}
<div class="card"><h2>📂 选择岗位组别</h2><div class="tabs">
<button type="button" class="tab active" data-role="供花">🌸 供花（{{ members['供花']|length }}）</button>
<button type="button" class="tab" data-role="供果">🍎 供果（{{ members['供果']|length }}）</button>
<button type="button" class="tab" data-role="佛学班">📖 佛学班（{{ members['佛学班']|length }}）</button>
<button type="button" class="tab" data-role="佛台">🪷 整理佛台</button>
</div></div>
<div class="card"><h2>➕ 加入固定成员</h2>
<form method="post" id="add-form"><input type="hidden" name="action" id="add-action" value="add">
<input type="hidden" name="role" id="add-role" value="供花"><input type="hidden" name="volunteer_id" id="selected-id">
<div class="row" style="margin-bottom:14px"><span>编号类别：</span><div class="branches">
<button type="button" class="branch active" data-branch="CHE">CHE</button>
<button type="button" class="branch" data-branch="STW">STW</button>
<button type="button" class="branch" data-branch="8XX">8xx</button></div></div>
<div class="row"><label class="search-box">编号或姓名
<input type="text" id="member-search" placeholder="输入编号（如208）或姓名（如蕭）" autocomplete="off" required>
<div id="search-results" class="results" hidden></div></label>
<label id="weekday-field" hidden>星期 <select name="weekday" id="weekday-select">{% for i in range(7) %}<option value="{{ i }}">星期{{ '一二三四五六日'[i] }}</option>{% endfor %}</select></label>
<button class="primary" type="submit" id="add-submit">加入供花组</button></div>
<p class="selected" id="selected-person"></p><p class="muted">输入姓名第一个字即可查询；选择结果后才可加入。正式编号由 Supabase 在册义工资料确认。</p>
</form></div>
{% for role in roles %}
<div class="card role-panel {{ 'active' if role == '供花' else '' }}" data-panel="{{ role }}">
<h2>{{ '🌸' if role == '供花' else ('🍎' if role == '供果' else '📖') }} {{ role }}（{{ members[role]|length }} 人）</h2>
{% if members[role] %}<table><thead><tr><th>编号</th><th>姓名</th><th>操作</th></tr></thead><tbody>
{% for member in members[role] %}<tr><td>{{ member.volunteer_id }}</td><td>{{ member.name }}</td>
<td><form method="post" onsubmit="return confirm('确定从{{ role }}移除 {{ member.name }}？');">
<input type="hidden" name="action" value="remove"><input type="hidden" name="role" value="{{ role }}">
<input type="hidden" name="volunteer_id" value="{{ member.volunteer_id }}"><button class="danger">移除</button></form></td></tr>{% endfor %}
</tbody></table>{% else %}<p>尚无固定成员。</p>{% endif %}</div>
{% endfor %}
<div class="card role-panel" data-panel="佛台"><h2>🪷 整理佛台 · 每星期固定轮值</h2>
<p class="muted">名单保存在 Supabase；实际自动排班仍按原有 Excel 读取。</p>
{% for i in range(7) %}
<h3 class="weekday">星期{{ '一二三四五六日'[i] }}（{{ buddha_members[i]|length }} 人）</h3>
{% if buddha_members[i] %}<table><thead><tr><th>顺序</th><th>编号</th><th>姓名</th><th>操作</th></tr></thead><tbody>
{% for member in buddha_members[i] %}<tr><td>{{ member.position }}</td><td>{{ member.volunteer_id }}</td><td>{{ member.name }}</td><td>
<form method="post" onsubmit="return confirm('确定从星期{{ '一二三四五六日'[i] }}移除 {{ member.name }}？');">
<input type="hidden" name="action" value="buddha_remove"><input type="hidden" name="weekday" value="{{ i }}">
<input type="hidden" name="volunteer_id" value="{{ member.volunteer_id }}"><button class="danger" type="submit">移除</button>
</form></td></tr>{% endfor %}</tbody></table>{% else %}<p>尚无固定人员。</p>{% endif %}
{% endfor %}</div>
</div>
<script>
(function(){
const tabs=document.querySelectorAll('.tab'), branches=document.querySelectorAll('.branch');
const input=document.getElementById('member-search'), results=document.getElementById('search-results');
const selected=document.getElementById('selected-id'), label=document.getElementById('selected-person');
const action=document.getElementById('add-action'), roleInput=document.getElementById('add-role');
const weekday=document.getElementById('weekday-field'), submit=document.getElementById('add-submit');
let branch='CHE', role='供花', timer=null, seq=0;
function clear(){seq++;selected.value='';label.textContent='';results.hidden=true;results.replaceChildren();}
function setRole(next){role=next;tabs.forEach(b=>b.classList.toggle('active',b.dataset.role===role));
document.querySelectorAll('.role-panel').forEach(p=>p.classList.toggle('active',p.dataset.panel===role));
action.value=role==='佛台'?'buddha_add':'add';roleInput.value=role==='佛台'?'':role;
weekday.hidden=role!=='佛台';submit.textContent=role==='佛台'?'加入佛台固定轮值':'加入'+role+'组';clear();}
tabs.forEach(b=>b.addEventListener('click',()=>setRole(b.dataset.role)));
branches.forEach(b=>b.addEventListener('click',()=>{branch=b.dataset.branch;branches.forEach(x=>x.classList.toggle('active',x===b));clear();if(input.value.trim())lookup();}));
async function lookup(){const q=input.value.trim();clear();const current=++seq;if(!q)return;
try{const response=await fetch('/schedule/fixed-roles/search?branch='+encodeURIComponent(branch)+'&q='+encodeURIComponent(q),{credentials:'same-origin'});
if(!response.ok)throw Error();const data=await response.json();if(current!==seq)return;
if(!data.results.length){const item=document.createElement('div');item.className='result';item.textContent='没有找到匹配的在册义工';results.append(item);}
for(const person of data.results){const item=document.createElement('button');item.type='button';item.className='result';item.textContent=person.id+'　'+person.name;
item.addEventListener('click',()=>{selected.value=person.id;label.textContent='已选择：'+person.name+'（'+person.id+'）';input.value=person.name;results.hidden=true;});results.append(item);}results.hidden=false;
}catch(e){if(current===seq){label.textContent='查询暂时不可用，请稍后重试';}}
}
input.addEventListener('input',()=>{clearTimeout(timer);clear();seq++;timer=setTimeout(lookup,220);});
document.getElementById('add-form').addEventListener('submit',e=>{if(!selected.value){e.preventDefault();alert('请先从搜索结果中选择正确的义工。');input.focus();}});
document.addEventListener('click',e=>{if(!e.target.closest('.search-box'))results.hidden=true;});
})();
</script></body></html>
"""

SCHEDULE_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="manifest"
      href="/schedule-manifest.json?v=5">

<link rel="icon"
      type="image/png"
      sizes="192x192"
      href="/static/schedule_icon_192.png?v=5">

<link rel="apple-touch-icon"
      href="/static/schedule_icon_512.png?v=5">

<meta name="theme-color"
      content="#43a047">

<title>负责人排班系统</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<script>
document.addEventListener("DOMContentLoaded", function () {

    function syncDates() {
        const mainDate = document.getElementById("main_date");

        const rawDate = document.getElementById("raw_single_date");
        const genDate = document.getElementById("generate_day_date");
        const fillDate = document.getElementById("fill_pending_date");
        const viewAssignedDate = document.getElementById("view_assigned_date");
        const overrideDate = document.getElementById("override_date");
        const prebookDate = document.getElementById("prebook_single_date");
        const publishDate = document.getElementById("publish_date");
        const noticeDate = document.getElementById("notice_date");

        if (!mainDate) return;

        if (rawDate) rawDate.value = mainDate.value;
        if (genDate) genDate.value = mainDate.value;
        if (fillDate) fillDate.value = mainDate.value;
        if (viewAssignedDate) viewAssignedDate.value = mainDate.value;
        if (overrideDate) overrideDate.value = mainDate.value;
        if (prebookDate) prebookDate.value = mainDate.value;
        if (publishDate) publishDate.value = mainDate.value;
        if (noticeDate) noticeDate.value = mainDate.value;
    }

    syncDates();

    const mainDate = document.getElementById("main_date");
    if (mainDate) {
        mainDate.addEventListener("change", syncDates);
    }
});
</script>

<script>
function openMasterDay(day) {
    const valid = Array.from(document.querySelectorAll('#master_day_select option')).some(o => o.value === day);
    if (!valid) { alert('请选择列表中的大日子'); return; }
    window.location.href = '/schedule?mode=special&override_date=' + encodeURIComponent(day) + '&panel=master-table';
}
document.addEventListener('DOMContentLoaded', function () {
    const select = document.getElementById('master_day_select');
    const calendar = document.getElementById('master_calendar');
    if (select) select.addEventListener('change', function () { openMasterDay(this.value); });
    if (calendar) calendar.addEventListener('change', function () {
        if (select && Array.from(select.options).some(o => o.value === this.value)) openMasterDay(this.value);
        else { alert('请选择未来90天内的大日子'); calendar.value = select ? select.value : ''; }
    });
    const removeCalendar = document.getElementById('remove_day_calendar');
    if (removeCalendar) removeCalendar.addEventListener('change', function () {
        const params = new URLSearchParams(window.location.search);
        params.set('mode', 'special');
        params.set('panel', 'master-table');
        params.set('remove_date', this.value);
        window.location.href = '/schedule?' + params.toString();
    });
    const panel = new URLSearchParams(window.location.search).get('panel');
    if (panel === 'master-table' && select && select.value !== new URLSearchParams(window.location.search).get('override_date')) { openMasterDay(select.value); return; }
    if (panel === 'master-table' || panel === 'extra-buddha' || panel === 'buddha-leave') showSpecial(panel);
});
function showSpecial(id) {
    if (id === 'master-table') {
        const select = document.getElementById('master_day_select');
        const shown = new URLSearchParams(window.location.search).get('override_date');
        if (select && select.value !== shown) { openMasterDay(select.value); return; }
    }
    const boxes = document.querySelectorAll(".special-box");

    boxes.forEach(function(box) {
        box.style.display = "none";
    });

    const target = document.getElementById(id);

    if (target) {
        target.style.display = "block";
        setTimeout(function () {
            target.scrollIntoView({ behavior: "smooth", block: "start" });
        }, 80);
    }
}
</script>

<style>
.login-wrap{
    max-width:620px;
    margin:60px auto;
    padding:20px;
}

.login-card{
    text-align:center;
}

.login-icon{
    font-size:64px;
    margin-bottom:10px;
}

.login-title{
    font-size:52px;
    font-weight:bold;
    color:#1e293b;
}

.login-subtitle{
    margin:15px 0 35px;
    color:#666;
    font-size:22px;
    line-height:1.8;
}

.login-input{
    margin-top:10px;
}

.login-input input{
    text-align:center;
    letter-spacing:6px;
    font-size:34px;
}

html {
    scroll-behavior: smooth;
}

@media (max-width:700px){
    .login-title{
        font-size:40px;
    }

    .login-subtitle{
        font-size:20px;
    }
}
/* 每天必用入口：醒目金黄色，避免与其他管理按钮混淆 */
.daily-entry-highlight { display:block; text-align:center; background:linear-gradient(135deg,#ffbf24,#f17800) !important; color:#172235 !important; border:3px solid #e18a00 !important; font-weight:900 !important; font-size:clamp(21px,3vw,30px) !important; padding:20px 14px !important; border-radius:15px !important; box-shadow:0 5px 0 #b76a00,0 9px 18px rgba(173,99,0,.23) !important; text-decoration:none; }
.daily-entry-highlight:hover { filter:brightness(1.06); transform:translateY(-2px); }
/* 滚动时固定在右下角，手机和电脑均可使用 */
.floating-home { position:fixed; right:20px; bottom:22px; z-index:9999; background:#245de6; color:white !important; padding:11px 16px; border-radius:999px; font-weight:800; font-size:15px; box-shadow:0 4px 15px rgba(0,0,0,.28); text-decoration:none; border:2px solid white; }
.floating-home:hover { background:#1247b9; }
@media(max-width:600px) { .floating-home { right:12px; bottom:12px; font-size:14px; padding:10px 13px; } }
</style>

</head>

<body>
{% if mode in ("day", "signup", "special", "controls") %}<a class="floating-home" href="{{ '/schedule/management/fixed?date=' ~ override_date if mode == 'special' else ('/schedule/management/settings?date=' ~ override_date if mode == 'controls' else '/schedule') }}" aria-label="返回上一级">⬅ 返回</a>{% endif %}
<div class="box">

<div class="top-bar">
    <h1>📅 负责人排班系统</h1>
    {% if mode in ("day", "signup", "special", "controls") %}
    <a class="btn-tool btn-blue" style="width:auto" href="{{ '/schedule/management/fixed?date=' ~ override_date if mode == 'special' else ('/schedule/management/settings?date=' ~ override_date if mode == 'controls' else '/schedule') }}">⬅ 返回上一级</a>
    {% endif %}

    <a class="btn-tool btn-gray" style="width:auto;min-width:160px;" href="/schedule/logout">
        🚪 退出登录
    </a>
</div>

{% if mode not in ("day", "signup", "special", "controls") %}
<div class="card">
 <h2 class="section-title">📊 今日报名概况 · {{ override_date }}</h2>
 <div class="action-grid">
  <div class="date-card">📋 已报名：<b>{{ day_summary.total }}</b></div>
  <div class="date-card">🟠 待安排：<b>{{ day_summary.pending }}</b></div>
  <div class="date-card">🟢 已安排：<b>{{ day_summary.assigned }}</b></div>
  <div class="date-card">⚪ 已取消：<b>{{ day_summary.cancelled }}</b></div>
 </div>
 <a class="btn-tool daily-entry-highlight" href="/schedule?mode=day&override_date={{ override_date }}">⭐ 进入今日安排 → <span style="font-size:.7em">（每日必用）</span></a>
</div>

<div class="card" id="management-functions">
    <h2 class="section-title">🌸 管理功能</h2>
    <div class="main-menu">
      <a class="btn-tool btn-orange" style="background:#ef7118;border-color:#ef7118;color:#fff" href="/schedule/notice_center?date={{ override_date }}&amp;return_section=management">📢 通知中心</a>
      <a class="btn-tool btn-purple" style="background:#8039e8;border-color:#8039e8;color:#fff" href="/schedule/signups">📋 报名记录</a>
      <a class="btn-tool btn-gray" style="background:#d64a84;border-color:#d64a84;color:#fff" href="/schedule/management/fixed?date={{ override_date }}">🪷 佛台及固定岗位管理</a>
      <a class="btn-tool btn-gray" style="background:#596879;border-color:#596879;color:#fff" href="/schedule/management/settings?date={{ override_date }}">⚙️ 系统设置</a>
      <a class="btn-tool btn-green" style="background:#b47a17;border-color:#b47a17;color:#fff" href="/reports">📚 资料中心</a>
      <a class="btn-tool btn-blue" style="background:#275de4;border-color:#275de4;color:#fff" href="/volunteer">👥 打开义工报名页面</a>
    </div>
</div>
{% endif %}

{% if mode in ("day", "signup", "special", "controls") %}
{% if mode == "day" %}

<div class="card today-panel" id="today-schedule">
    <h2 class="section-title">📅 今日概况</h2>
    <div class="dashboard-note">负责人打开后，先确认日期、发布状态、报名人数和提醒。</div>

    <div class="date-card">
        <b>🗓 日期资料</b><br>
        阳历：{{ special_day_info.solar_date }}<br>
        农历：{{ special_day_info.lunar_text }}<br>

        {% if special_day_info.is_special %}
            <span class="text-red" style="font-weight:bold;">
                🔴 {{ special_day_info.special_names | join("、") }}
            </span><br>
            模板：{{ special_day_info.template_text }}<br>

            {% if special_day_info.setup_shifu %}
                🛕 需要设师父供台<br>
            {% endif %}

            {% if special_day_info.remove_next_day %}
                📦 明日需要收供台<br>
            {% endif %}
        {% else %}
            🟢 平时日
        {% endif %}

        {% if remove_info.need_remove_today_after_12 %}
            <br>⚠️ 今日中午后需要收昨日供台
        {% endif %}
    </div>

    <div class="date-status-row compact-date-row">

        <div class="date-picker-card compact-date-picker">

            <div class="status-title">
                📅 选择工作日期
            </div>

            <input
                type="date"
                id="main_date"
                class="date-input"
                value="{{ override_date }}"
                onchange="goToScheduleDate(this.value)"
            >

        </div>

    </div>

    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-title">👥 已报名</div>
            <div class="kpi-value">{{ day_summary.total }}</div>
        </div>

        <div class="kpi-card">
            <div class="kpi-title">🟡 待安排</div>
            <div class="kpi-value">{{ day_summary.pending }}</div>
        </div>

        <div class="kpi-card">
            <div class="kpi-title">🟢 已安排</div>
            <div class="kpi-value">{{ day_summary.assigned }}</div>
        </div>

        <div class="kpi-card">
            <div class="kpi-title">⚪ 已取消</div>
            <div class="kpi-value">{{ day_summary.cancelled }}</div>
        </div>
    </div>

    <div class="info-box" style="margin-top:14px;">
        <b>👥 当日义工报名详情（{{ override_date }}）</b>
        <div style="font-size:13px;color:#64748b;margin:5px 0 10px;">显示所有报名记录，包括待安排、已安排及已取消。</div>
        {% if signup_details %}
        <div style="overflow-x:auto;">
          <table style="width:100%;border-collapse:collapse;font-size:14px;min-width:590px;">
            <thead><tr style="background:#edf3f9;text-align:left;">
              <th style="padding:9px;border-bottom:1px solid #dce5ee;">义工姓名</th>
              <th style="padding:9px;border-bottom:1px solid #dce5ee;">报名岗位</th>
              <th style="padding:9px;border-bottom:1px solid #dce5ee;">报名时段</th>
              <th style="padding:9px;border-bottom:1px solid #dce5ee;">安排地点</th>
              <th style="padding:9px;border-bottom:1px solid #dce5ee;">状态</th>
            </tr></thead>
            <tbody>
            {% for person in signup_details %}
              <tr style="border-bottom:1px solid #edf1f5;{% if person.status == 'cancelled' %}opacity:.55;{% endif %}">
                <td style="padding:9px;">{{ person.name or person.volunteer_id or '—' }}</td>
                <td style="padding:9px;">{{ person.role or '—' }}</td>
                <td style="padding:9px;white-space:nowrap;">{{ person.start_time|string if person.start_time else '—' }} ～ {{ person.end_time|string if person.end_time else '—' }}</td>
                <td style="padding:9px;">
                  {% if person.assignment_places %}
                    <div style="display:flex;flex-wrap:wrap;align-items:center;gap:6px;">
                      {% for place in person.assignment_places %}
                        <span style="display:inline-block;white-space:nowrap;padding:5px 9px;border-radius:7px;font-size:13px;line-height:1.4;{% if '观音堂' in place %}background:#ffedd5;color:#9a3412;{% elif '活动中心' in place %}background:#fef9c3;color:#854d0e;{% else %}background:#e8f3ff;color:#1e40af;{% endif %}">{{ place }}</span>
                      {% endfor %}
                    </div>
                  {% else %}{{ person.assigned_place or '—' }}{% endif %}
                </td>
                <td style="padding:9px;white-space:nowrap;">
                  {% if person.status == 'assigned' %}<span style="color:#16803c;font-weight:700;">🟢 已安排</span>
                  {% elif person.status == 'cancelled' %}<span style="color:#64748b;">⚪ 已取消</span>
                  {% elif person.status == 'pending' %}<span style="color:#a66a00;font-weight:700;">🟡 待安排</span>
                  {% else %}{{ person.status }}{% endif %}
                </td>
              </tr>
            {% endfor %}
            </tbody>
          </table>
        </div>
        {% else %}<div style="color:#64748b;">这一天暂时没有报名记录。</div>{% endif %}
    </div>

    <div class="info-box">
        <b>📢 缺义工提醒</b><br>

        {% if shortage_summary %}
            {% for line in shortage_summary %}
                {{ line }}<br>
            {% endfor %}
        {% else %}
            🟢 目前没有缺义工岗位
        {% endif %}
    </div>
</div>

<div class="card">
    <h2 class="section-title">🔔 今日提醒</h2>
    <div class="reminder-stack">
<div class="alert alert-error">
    <h3 style="margin-top:0;">🔴 未安排提醒</h3>
    今日：<b>{{ pending_counts.today }}</b> 人　
    明日：<b>{{ pending_counts.tomorrow }}</b> 人　
    本月：<b>{{ pending_counts.month }}</b> 人
</div>



        {% if republish_info and republish_info.need_republish %}
        <div class="alert alert-warning">
            🟠 正式值班表已有更新<br>
            最后更新时间：{{ republish_info.last_schedule_change }}<br>
            请到通知中心复制最新正式值班表 WhatsApp。
        </div>
        {% endif %}
<div class="alert alert-warning">
    <h2>🌸 供台报名提醒</h2>

    {% for item in supply_alerts %}
        <div style="margin-bottom:14px;">
            <b>
                {{ item.date.strftime("%m/%d") }}
                {{ item.festival_label }}
            </b><br>

            {% if item.scheduled_names %}
                <span style="color:#176b39;">🟢 已安排设供台：{{ "、".join(item.scheduled_names) }}</span><br>
            {% endif %}
            {% if item.names %}
                义工自行报名：{{ "、".join(item.names) }}
            {% elif not item.scheduled_names %}
                <span style="color:#c0392b;">尚未安排设供台人员，也没有义工报名</span>
            {% endif %}
        </div>
    {% endfor %}
</div>


    </div>
</div>

<div class="card" id="today-operations">
    <h2 class="section-title">📝 今日操作｜负责人工作流程</h2>
    <div class="dashboard-note" style="margin-bottom:14px;line-height:1.8;">
        请按① → ② → ③ → ④操作。以下按钮针对目前选择的排班日期；发送 WhatsApp 前请先核对日期及名单。
    </div>

    <div class="action-grid">
        <div>
            <div style="font-weight:700;margin-bottom:7px;">① 晚上7:00｜发出报名通知</div>
            <a href="/schedule/notice_center?date={{ override_date }}&amp;return_section=today" class="btn-tool btn-orange">
                📢 报名通知／通知中心
                <span class="helper-text">进入通知中心，复制明日的报名通知，发送到 WhatsApp 群组；也可查看缺义工通知。</span>
            </a>
        </div>

        <div>
            <div style="font-weight:700;margin-bottom:7px;">② 晚上10:00｜安排报名义工</div>
            {% if not is_published %}
                {% if not is_today_or_past %}
                <form method="post" action="/schedule/generate_day">
                    <input type="hidden" name="date" id="generate_day_date" value="{{ override_date }}">
                    <button type="submit" class="btn-tool btn-blue">
                        🔄 重新自动排班
                        <span class="helper-text">报名截止后点击；系统依报名资料安排岗位、地点及时间。发布前可再次执行。</span>
                    </button>
                </form>
                {% else %}
                <form method="post" action="/schedule/fill_pending">
                    <input type="hidden" name="date" id="fill_pending_date" value="{{ override_date }}">
                    <button type="submit" class="btn-tool btn-blue">
                        ➕ 补入待安排
                        <span class="helper-text">今天或过去日期：把尚未安排的报名义工补入空缺，不重新生成整张排班表。</span>
                    </button>
                </form>
                {% endif %}
            {% else %}
                <div class="btn-tool btn-gray" style="cursor:default;box-sizing:border-box;">
                    ✅ 已发布，自动补排生效
                    <span class="helper-text">正式发布后，新报名由系统自动补排，不必再按「重新自动排班」。</span>
                </div>
            {% endif %}
        </div>

        <div>
            <div style="font-weight:700;margin-bottom:7px;">③ 排班完成后｜检查及调整</div>
            <form method="post" action="/schedule/view_assigned">
                <input type="hidden" name="date" id="view_assigned_date" value="{{ override_date }}">
                <button type="submit" class="btn-tool btn-purple">
                    👀 查看排班
                    <span class="helper-text">检查义工姓名、值班时间、岗位和地点；如有需要，进入后手动调整。</span>
                </button>
            </form>
        </div>

        <div>
            <div style="font-weight:700;margin-bottom:7px;">④ 检查无误后｜正式发布</div>
            <form method="post" action="/schedule/publish">
                <input type="hidden" name="date" id="publish_date" value="{{ override_date }}">
                <button type="submit" class="btn-tool {% if is_published %}btn-gray{% else %}btn-green{% endif %}">
                    {% if is_published %}
                        ✅ 已发布
                        <span class="helper-text">值班表已生效。若有新报名，系统会自动补排；可继续查看最新安排。</span>
                    {% else %}
                        🟢 发布正式值班表
                        <span class="helper-text">确认名单正确后点击发布，再到通知中心复制正式值班表发送 WhatsApp。</span>
                    {% endif %}
                </button>
            </form>
        </div>
    </div>

    <div class="dashboard-note" style="margin-top:14px;line-height:1.8;">
        💡 <b>发布后有新义工报名：</b>系统会自动补入空缺岗位。负责人可按「查看排班」核对最新名单；不需要重新自动排班或重新发布。<br>
        📌 <b>提醒：</b>进入通知中心或查看排班，不代表 WhatsApp 已发送或名单已确认；请负责人自行核对。
    </div>

    {% if is_published %}
    <div style="margin-top:12px;">
        <a href="/schedule/attendance_status" class="btn-tool" style="background:#078b9b;border-color:#078b9b;color:#fff;">
            📋 签到情况
            <span class="helper-text">发布后可查看义工签到纪录及当天签到情况。</span>
        </a>
    </div>
    {% endif %}
</div>


{% endif %}
{% if mode == "signup" %}
<div class="card">
    <h2 class="section-title">📋 报名录入</h2>
    <div class="dashboard-note">负责人只需要在这里加入报名：可以人工输入，也可以贴 WhatsApp 文字解析。</div>

    <div class="signup-unified">

        <div class="signup-block">
            <h3>➕ 负责人代报名</h3>

            <form method="post" action="/schedule/prebook_add">
                <input type="hidden" name="mode" value="day">
                <input type="hidden" name="single_date" id="prebook_single_date" value="{{ override_date }}">

                <div class="form-row-large">
                    <label>义工编号 / 姓名</label>

                    <div class="branch-search-row">
                        <button
                            type="button"
                            id="admin_branch_btn"
                            onclick="toggleAdminBranch()"
                            class="branch-toggle-btn">
                            CHE
                        </button>

                        <input type="hidden" id="admin_branch" name="branch" value="CHE">

                        <input
                            name="keyword"
                            id="admin_keyword"
                            placeholder="例如：108 / 姓名">
                    </div>

                    <div id="volunteer_lookup_result" class="lookup-text"></div>
                </div>

                {% for role in roles %}
                <label class="role-btn">

                    <input
                        type="checkbox"
                        name="roles"
                        value="{{ role }}">

                    {% if role == "值班" %}
                        👷 值班
                    {% elif role == "卫生" %}
                        🧹 卫生
                    {% elif role == "供台" %}
                        🛕 供台
                    {% else %}
                        {{ role }}
                    {% endif %}

                </label>
                {% endfor %}

                <div class="form-row-large">
                    <p>时间（只给值班用）</p>

                    <div class="time-grid">
                        <div>
                            <label>开始</label>
                            <select name="start_time" style="width:100%;">
                            {% for t in times %}
                                <option value="{{ t }}">{{ t }}</option>
                            {% endfor %}
                            </select>
                        </div>

                        <div>
                            <label>结束</label>
                            <select name="end_time" style="width:100%;">
                            {% for t in times %}
                                <option value="{{ t }}">{{ t }}</option>
                            {% endfor %}
                            </select>
                        </div>
                    </div>
                </div>

                <button type="submit" class="btn-tool btn-blue">➕ 加入报名</button>
            </form>
        </div>

        <div class="signup-divider"></div>

        <div class="signup-block">
            <h3>📥 WhatsApp 报名解析</h3>

            <form method="post" action="/schedule/parse_raw">
                <input type="hidden" name="single_date" id="raw_single_date" value="{{ override_date }}">

                <textarea
                    name="raw_signup"
                    rows="8"
                    placeholder="输入姓名 / 时间，例如：&#10;张三 10:00am-2:00pm&#10;李四 8:00am-12:00pm"></textarea>

                <br><br>

                <button type="submit" class="btn-tool btn-blue">📥 解析加入报名</button>
            </form>
        </div>

    </div>
</div>



{% endif %}
{% if mode == "special" %}
    <div class="card">
        <h2 class="section-title">⚙️ 特殊设置</h2>

        <div class="quick-actions">
            <button type="button" class="btn-tool btn-gray" onclick="showSpecial('master-table')">
                🛕 师父供台
                <span class="helper-text">点击后自动跳到设置区</span>
            </button>

            <button type="button" class="btn-tool btn-gray" onclick="showSpecial('extra-buddha')">
                🌸 初一补人
            </button>

            <button type="button" class="btn-tool btn-gray" onclick="showSpecial('buddha-leave')">
                🔁 佛台请假
            </button>
        </div>

        <!-- 1. 师父供台 -->
        <div id="master-table" class="special-box" style="display:none;">
            <h3>🛕 师父供台设置</h3>

            {% if supply_signups %}
            <div style="
                background:#fff8e1;
                border:2px solid #f3d36b;
                border-radius:14px;
                padding:15px;
                margin:15px 0;
                font-size:22px;
            ">
                <b>🌸 已报名供台义工：</b><br>
                {% for p in supply_signups %}
                    {{ p.name }}{% if not loop.last %}、{% endif %}
                {% endfor %}
            </div>
            {% endif %}

            <div class="info-box" style="margin:12px 0;">
                <b>📅 提前选择未来大日子（90天内）</b>
                <p style="font-size:16px;">选好日期后会自动载入该日已保存的供台人员。</p>
                {% if master_dates %}
                <select id="master_day_select" aria-label="选择大日子" style="width:100%;font-size:19px;">
                    {% for item in master_dates %}
                    <option value="{{ item.date }}" {% if item.date == selected_master_date %}selected{% endif %}>{{ item.label }}</option>
                    {% endfor %}
                </select>
                <label for="master_calendar">或打开日历选择：</label>
                <input type="date" id="master_calendar" value="{{ selected_master_date }}"
                       min="{{ master_dates[0].date }}" max="{{ master_dates[-1].date }}"
                       style="width:100%;font-size:19px;padding:9px;">
                <small>日历只接受上方列表中的大日子；普通日期不会保存。</small>
                {% else %}
                <p>未来90天内没有找到大日子，请检查农历规则。</p>
                {% endif %}
            </div>

            <form method="post" action="/schedule/save_day_flags">
                <input type="hidden" name="form_type" value="master_table">
                <input type="hidden" name="date" value="{{ selected_master_date }}">

                <div class="info-box">
                    系统会根据初一、十五及佛诞日自动判断是否需要设供台。<br>
                    若前一天应收供台但未安排人员，系统会继续提醒收供台。
                </div>

                <h4>🪷 设师父供台人员</h4>
                <p style="color:#666;font-size:20px;">通常 1 至 2 位佛台组人员即可。</p>

                <select name="setup_person_1">

                    <option value="">请选择</option>

                    {% for p in supply_signups %}
                        <option value="{{ p.name }}"
                            {% if setup_saved and setup_saved[0] == p.name %}selected{% endif %}>
                            {{ p.name }}
                        </option>
                    {% endfor %}

                    {% for n in buddha_names %}
                        <option value="{{ n }}"
                            {% if setup_saved and setup_saved[0] == n %}selected{% endif %}>
                            {{ n }}
                        </option>
                    {% endfor %}

                </select>

                <select name="setup_person_2">
                    <option value="">请选择</option>

                    {% for p in supply_signups %}
                        <option value="{{ p.name }}"
                            {% if setup_saved|length > 1 and setup_saved[1] == p.name %}selected{% endif %}>
                            {{ p.name }}
                        </option>
                    {% endfor %}

                    {% for n in buddha_names %}
                        <option value="{{ n }}" {% if setup_saved|length > 1 and setup_saved[1] == n %}selected{% endif %}>
                            {{ n }}
                        </option>
                    {% endfor %}
                </select>

                <hr>

                <button type="submit" class="btn-tool btn-blue">💾 保存设供台人员</button>
            </form>

            <form method="post" action="/schedule/save_day_flags" style="margin-top:24px;">
                <input type="hidden" name="form_type" value="remove_table">
                <input type="hidden" name="setup_date" value="{{ selected_master_date }}">
                <h4>📦 收师父供台人员</h4>
                <label for="remove_day_calendar"><b>📅 选择收供台日期（未来90天内，普通日也可以）</b></label>
                <input id="remove_day_calendar" type="date" name="date" value="{{ selected_remove_date }}"
                       min="{{ remove_min_date }}" max="{{ remove_max_date }}" required
                       style="width:100%;font-size:19px;padding:9px;margin:10px 0;">
                <p style="font-size:20px;font-weight:bold;color:#b45309;">📅 当前收供台安排日期：{{ selected_remove_date }}</p>
                <p style="font-size:15px;color:#666;">默认日期与值班一致（18:00后切换明天）；手动更换日期会载入该日已保存的收供台人员。</p>
                <p style="color:#666;font-size:20px;">
                    通常 1 至 2 位佛台组人员。若没有安排人员，系统视为继续供奉。
                </p>

                <select name="remove_person_1">
                    <option value="">请选择</option>
                    {% for n in buddha_names %}
                    <option value="{{ n }}" {% if remove_saved and remove_saved[0] == n %}selected{% endif %}>
                        {{ n }}
                    </option>
                    {% endfor %}
                </select>

                <select name="remove_person_2">
                    <option value="">请选择</option>
                    {% for n in buddha_names %}
                    <option value="{{ n }}" {% if remove_saved|length > 1 and remove_saved[1] == n %}selected{% endif %}>
                        {{ n }}
                    </option>
                    {% endfor %}
                </select>

                <p>普通义工帮忙收供台：</p>
                <input
                    name="remove_extra_person"
                    value="{{ remove_saved[2] if remove_saved|length > 2 else '' }}"
                    placeholder="例如：张三，可空"
                    style="width:100%;font-size:22px;"
                >

                <br><br>

                <button type="submit" class="btn-tool btn-blue">💾 保存收供台人员</button>
            </form>
        </div>

        <!-- 2. 初一补人 -->
        <div id="extra-buddha" class="special-box" style="display:none;">
            <h3>🌸 初一整理佛台补人</h3>

            {% if extra_buddha_dates %}
            <form method="get" action="/schedule">
                <input type="hidden" name="mode" value="special">
                <input type="hidden" name="override_date" value="{{ override_date }}">
                <input type="hidden" name="panel" value="extra-buddha">
                <label for="extra_buddha_date">选择初一日期：</label>
                <select id="extra_buddha_date" name="extra_buddha_date" onchange="this.form.submit()" style="width:100%;font-size:22px;">
                    {% for d in extra_buddha_dates %}
                    <option value="{{ d }}" {% if d == selected_extra_buddha_date %}selected{% endif %}>{{ d }}（初一）</option>
                    {% endfor %}
                </select>
            </form>
            <div class="info-box">
                <b>🛕 当天负责整理佛台义工</b><br><br>

                {% if extra_buddha_fixed_names %}
                    <form method="post" action="/schedule/override">
                        <input type="hidden" name="date" value="{{ selected_extra_buddha_date }}">
                        <input type="hidden" name="return_panel" value="extra-buddha">
                        <input type="hidden" name="return_override_date" value="{{ override_date }}">
                        {% for n in extra_buddha_fixed_names %}
                        <div style="margin:10px 0;display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                            <span>🟢 {{ n }}</span>
                            <input type="hidden" name="original_name" value="{{ n }}">
                            <select name="replacement_name" style="font-size:20px;max-width:100%;">
                                <option value="">不变，保留原本</option>
                                <option value="__REMOVE__">请假，不找替补</option>
                                {% for candidate in buddha_names %}
                                {% if candidate != n %}<option value="{{ candidate }}">换人：{{ candidate }}</option>{% endif %}
                                {% endfor %}
                            </select>
                        </div>
                        {% endfor %}
                        <button type="submit" class="btn-tool btn-blue">💾 保存请假 / 换人</button>
                    </form>
                    <form method="post" action="/schedule/override/reset" style="margin-top:10px;">
                        <input type="hidden" name="date" value="{{ selected_extra_buddha_date }}">
                        <input type="hidden" name="return_panel" value="extra-buddha">
                        <input type="hidden" name="return_override_date" value="{{ override_date }}">
                        <button type="submit" class="btn-tool btn-red">↩ 恢复当天固定佛台名单</button>
                    </form>
                {% else %}
                    ⚪ 该日没有固定整理佛台义工
                    <form method="post" action="/schedule/override/reset" style="margin-top:10px;">
                        <input type="hidden" name="date" value="{{ selected_extra_buddha_date }}">
                        <input type="hidden" name="return_panel" value="extra-buddha">
                        <input type="hidden" name="return_override_date" value="{{ override_date }}">
                        <button type="submit" class="btn-tool btn-red">↩ 恢复当天固定佛台名单</button>
                    </form>
                {% endif %}
            </div>

            <form method="post" action="/schedule/save_day_flags">
                <input type="hidden" name="form_type" value="extra_buddha">
                <input type="hidden" name="date" value="{{ selected_extra_buddha_date }}">

                <p style="font-size:20px;color:#666;">
                    仅初一需要。可增加第 3 位整理佛台义工。
                </p>

                <p>额外第 3 位整理佛台义工：</p>

                <select name="extra_buddha_person" style="width:100%;font-size:22px;">
                    <option value="">不需要增加</option>
                    {% for n in buddha_names %}
                    <option value="{{ n }}" {% if extra_buddha_saved == n %}selected{% endif %}>
                        {{ n }}
                    </option>
                    {% endfor %}
                </select>

                <br><br>

                <button type="submit" class="btn-tool btn-blue">💾 保存初一补人</button>
            </form>
            {% else %}
            <p>未来90天内没有可选择的初一日期。</p>
            {% endif %}
        </div>

        <!-- 3. 佛台请假 -->
        <div id="buddha-leave" class="special-box" style="display:none;">
            <h3>🔁 佛台请假 / 换人</h3>

            <form method="get" action="/schedule" style="margin-bottom:18px;">
                <input type="hidden" name="mode" value="special">
                <input type="hidden" name="override_date" value="{{ override_date }}">
                <input type="hidden" name="panel" value="buddha-leave">
                <label for="buddha_leave_date">选择请假 / 换人日期：</label>
                <input type="date" id="buddha_leave_date" name="buddha_leave_date"
                       value="{{ selected_buddha_leave_date }}"
                       min="{{ buddha_leave_dates[0] }}" max="{{ buddha_leave_dates[-1] }}"
                       onchange="this.form.submit()"
                       style="width:100%;font-size:22px;box-sizing:border-box;">
            </form>

            <form method="post" action="/schedule/override">
                <input type="hidden" name="date" value="{{ selected_buddha_leave_date }}">
                <input type="hidden" name="return_panel" value="buddha-leave">
                <input type="hidden" name="return_override_date" value="{{ override_date }}">

                {% if buddha_leave_names %}
                    {% for old_name in buddha_leave_names %}
                        <div style="font-size:22px; margin:12px 0;">
                            {{ old_name }}
                            <input type="hidden" name="original_name" value="{{ old_name }}">

                            换成：
                            <select name="replacement_name">
                                <option value="">不换，保留原本</option>
                                <option value="__REMOVE__">请假，不找替补</option>

                                {% for n in buddha_names %}
                                <option value="{{ n }}">{{ n }}</option>
                                {% endfor %}
                            </select>
                        </div>
                    {% endfor %}

                    <button type="submit" class="btn-tool btn-blue">💾 保存佛台请假 / 换人</button>
                {% else %}
                    <p style="font-size:22px;">这一天没有固定佛台。</p>
                {% endif %}
            </form>
            <form method="post" action="/schedule/override/reset" style="margin-top:14px;">
                <input type="hidden" name="date" value="{{ selected_buddha_leave_date }}">
                <input type="hidden" name="return_panel" value="buddha-leave">
                <input type="hidden" name="return_override_date" value="{{ override_date }}">
                <button type="submit" class="btn-tool btn-red">↩ 恢复当天固定佛台名单</button>
            </form>
        </div>
    </div>

{% endif %}
{% if mode == "controls" %}
<div class="card">

    <h2 class="section-title">⚙️ 报名开放控制</h2>

    <table class="setting-table">

        <!-- 多日报名 -->

        <tr>
            <td style="padding:12px 0;">
                <b>📅 下个月多日报名</b><br>

                {% if multi_day_signup_open %}
                    <span style="color:green;font-weight:bold;">
                        🟢 已开放
                    </span>
                {% else %}
                    <span style="color:#999;font-weight:bold;">
                        ⚪ 已关闭
                    </span>
                {% endif %}
            </td>

            <td style="text-align:right;white-space:nowrap;">

                <form method="post"
                      action="/schedule/toggle_setting"
                      class="inline-form">

                    <input type="hidden"
                           name="key"
                           value="multi_day_signup_open">

                    <input type="hidden"
                           name="force"
                           value="true">

                    <button class="btn-tool btn-green mini-btn">
                        开启
                    </button>

                </form>

                <form method="post"
                      action="/schedule/toggle_setting"
                      class="inline-form">

                    <input type="hidden"
                           name="key"
                           value="multi_day_signup_open">

                    <input type="hidden"
                           name="force"
                           value="false">

                    <button class="btn-tool btn-red mini-btn">
                        关闭
                    </button>

                </form>

            </td>
            
        </tr>

        <tr>
            <td colspan="2">
                <hr>
            </td>
        </tr>

        <!-- 膳食 -->

        <tr>

            <td style="padding:12px 0;">

                <b>🍱 膳食组报名</b><br>

                {% if meal_signup_open %}
                    <span style="color:green;font-weight:bold;">
                        🟢 已开放
                    </span>
                {% else %}
                    <span style="color:#999;font-weight:bold;">
                        ⚪ 已关闭
                    </span>
                {% endif %}

            </td>

            <td style="text-align:right;white-space:nowrap;">

                <form method="post"
                      action="/schedule/toggle_setting"
                      class="inline-form">

                    <input type="hidden"
                           name="key"
                           value="meal_signup_open">

                    <input type="hidden"
                           name="force"
                           value="true">

                    <button class="btn-tool btn-green mini-btn">
                        开启
                    </button>

                </form>

                <form method="post"
                      action="/schedule/toggle_setting"
                      class="inline-form">

                    <input type="hidden"
                           name="key"
                           value="meal_signup_open">

                    <input type="hidden"
                           name="force"
                           value="false">

                    <button class="btn-tool btn-red mini-btn">
                        关闭
                    </button>

                </form>

            </td>

        </tr>

        <tr>
            <td colspan="2">
                <hr>
            </td>
        </tr>

        <!-- 素食结缘 -->

        <tr>

            <td style="padding:12px 0;">

                <b>🥗 素食结缘报名</b><br>

                {% if food_offering_open %}
                    <span style="color:green;font-weight:bold;">
                        🟢 已开放
                    </span>
                {% else %}
                    <span style="color:#999;font-weight:bold;">
                        ⚪ 已关闭
                    </span>
                {% endif %}

            </td>

            <td style="text-align:right;white-space:nowrap;">

                <form method="post"
                    action="/schedule/toggle_setting"
                    class="inline-form">

                    <input type="hidden"
                        name="key"
                        value="food_offering_open">

                    <input type="hidden"
                        name="force"
                        value="true">

                    <button class="btn-tool btn-green mini-btn">
                        开启
                    </button>

                </form>

                <form method="post"
                    action="/schedule/toggle_setting"
                    class="inline-form">

                    <input type="hidden"
                        name="key"
                        value="food_offering_open">

                    <input type="hidden"
                        name="force"
                        value="false">

                    <button class="btn-tool btn-red mini-btn">
                        关闭
                    </button>

                </form>

            </td>

        </tr>

    </table>

</div>



{% endif %}

{% endif %}

</div>
<script>
function toggleAdminBranch(){
    const btn = document.getElementById("admin_branch_btn");
    const branch = document.getElementById("admin_branch");

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

window.addEventListener("load", function () {
    if (window.location.hash === "#today-schedule") {
        setTimeout(function () {
            var target = document.getElementById("today-schedule");
            if (target) {
                target.scrollIntoView({
                    behavior: "smooth",
                    block: "start"
                });
            }
        }, 300);
    }
});

</script>
<script>
function goToScheduleDate(dateValue) {
    if (!dateValue) return;

    window.location.href =
        "/schedule?mode=day&override_date=" + encodeURIComponent(dateValue);
}
</script>
<script>
// 每20秒轻量检查；报名/取消/自动补排变化后自动更新负责人数据。
// 正在输入或修改表单时暂停刷新，避免丢失未保存内容。
(function () {
    const dateValue = {{ override_date|tojson }};
    const apiUrl = '/schedule/api/admin-signup-version?date=' + encodeURIComponent(dateValue);
    let previousVersion = null;
    let busy = false;
    let editing = false;
    document.addEventListener('input', function (event) {
        if (event.target.closest('input,textarea,select,[contenteditable]')) editing = true;
    }, true);
    async function checkSignupUpdates() {
        if (busy || document.hidden || editing) return;
        busy = true;
        try {
            const response = await fetch(apiUrl, {cache:'no-store', credentials:'same-origin'});
            if (!response.ok) return;
            const data = await response.json();
            if (previousVersion !== null && data.version !== previousVersion) {
                const active = document.activeElement;
                if (active && active.matches('input,textarea,select,[contenteditable]')) return;
                window.location.reload();
                return;
            }
            previousVersion = data.version;
        } catch (error) {
            // 网络暂时断线时保留当前页面，下次继续检查。
        } finally { busy = false; }
    }
    checkSignupUpdates();
    setInterval(checkSignupUpdates, 20000);
})();
</script>
</body>
</html>
"""


DAY_OUTPUT_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>正式值班表</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>

.page-wrap{
    max-width:1100px;
    margin:auto;
    padding:25px;
}

.output-box{
    min-height:650px;
}

.output-box textarea{
    width:100%;
    height:650px;

    font-size:26px;
    line-height:1.7;

    padding:18px;

    border-radius:16px;
    border:2px solid #d6d6d6;

    box-sizing:border-box;

    resize:vertical;
}

</style>

<script>
function copySchedule(){

    const ta = document.getElementById("output");

    if(!ta){
        alert("❌ 找不到值班表内容");
        return;
    }

    ta.focus();
    ta.select();
    ta.setSelectionRange(0, 999999);

    const ok = document.execCommand("copy");

    if(ok){
        alert("✅ 已复制，可以贴去 WhatsApp 群");
    }else{
        alert("❌ 复制失败，请手动全选复制");
    }
}
</script>

</head>

<body>

<div class="page-wrap">

<div class="card">

    <h1 class="page-title">
        📋 正式值班表
    </h1>

    <p class="page-subtitle">
        可直接复制到 WhatsApp 发布。
    </p>

    <div class="btn-row">

        <a
            class="btn-tool btn-gray"
            href="/schedule?mode=day">

            ⬅ 返回负责人首页

        </a>

        <button
            type="button"
            class="btn-tool btn-green"
            onclick="copySchedule()">

            📋 一键复制

        </button>

    </div>

    <div class="output-box">

        <textarea
            id="output"
            readonly>{{ output }}</textarea>

    </div>

</div>

</div>

</body>
</html>
"""


MONTHLY_PREBOOK_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>月预报名表</title>

<link rel="stylesheet"
      href="{{ url_for('static', filename='css/toolbox.css') }}">

<style>
.page-wrap{
    max-width:1100px;
    margin:auto;
    padding:25px;
}

.output-box textarea{
    width:100%;
    height:650px;
    font-size:26px;
    line-height:1.7;
    padding:18px;
    border-radius:16px;
    border:2px solid #d6d6d6;
    box-sizing:border-box;
    resize:vertical;
}
</style>

<script>
function copyOutput(btn){
    const text = document.getElementById("output").value;
    navigator.clipboard.writeText(text);

    btn.innerText = "✅ 已复制";

    setTimeout(function(){
        btn.innerText = "📋 一键复制";
    }, 2000);
}
</script>

</head>

<body>

<div class="page-wrap">

<div class="card">

    <h1 class="page-title">📢 月预报名表</h1>

    <p class="page-subtitle">
        可直接复制到 WhatsApp 群发布。
    </p>

    <div class="btn-row">

        <a class="btn-tool btn-gray"
           href="/schedule?mode=prebook">
            ⬅ 返回月预报名
        </a>

        <button class="btn-tool btn-green"
                onclick="copyOutput(this)">
            📋 一键复制
        </button>

    </div>

    <div class="output-box">
        <textarea id="output" readonly>{{ output }}</textarea>
    </div>

</div>

</div>

</body>
</html>
"""


