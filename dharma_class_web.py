# dharma_class_web.py


from io import BytesIO
from db import get_conn
from datetime import date
from flask import send_file
from openpyxl import Workbook, load_workbook
from psycopg2.extras import RealDictCursor
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from flask import Blueprint, render_template_string, request, redirect, url_for, flash, send_file


dharma_class_bp = Blueprint(
    "dharma_class",
    __name__,
    url_prefix="/class"
)


STATUS_LABELS = {
    "present": "出席",
    "absent": "缺席",
    "leave": "请假",
    "late": "迟到",
}

def beautify_simple_excel(ws):

    header_fill = PatternFill("solid", fgColor="5B9BD5")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = center

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0

        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            value = str(value) if value is not None else ""
            max_len = max(max_len, len(value))

        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    for row_idx in range(1, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 24

@dharma_class_bp.route("/")
def class_home():

    today_str = date.today().isoformat()

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    g.id,
                    g.name,
                    count(s.id) as total_students,
                    sum(case when a.status in ('present', 'late') then 1 else 0 end) as attended_count,
                    count(a.id) as marked_count
                from dharma_class_groups g
                left join dharma_students s
                    on s.group_id = g.id
                   and s.status = 'active'
                left join dharma_attendance a
                    on a.student_id = s.id
                   and a.class_date = %s
                where g.is_active = true
                group by g.id, g.name, g.sort_order
                order by g.sort_order, g.id
            """, (today_str,))
            group_stats = cur.fetchall()

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>蕉赖佛学班系统</title>
<link rel="stylesheet" href="/static/css/toolbox.css">
</head>

<body>
<div class="page">

    <div class="card">
        <h1 class="page-title">📘 蕉赖佛学班系统</h1>
        <p class="page-subtitle">今日：{{ today_str }}</p>

        <div class="summary-grid">
            {% for g in group_stats %}
            <div class="summary-box">
                <div style="font-size:20px;font-weight:700;">{{ g.name }}</div>
                <div style="font-size:28px;font-weight:800;margin-top:6px;">
                    {{ g.attended_count or 0 }} / {{ g.total_students or 0 }}
                </div>
                <div style="color:#777;margin-top:4px;">
                    已点名 {{ g.marked_count or 0 }} 位
                </div>
            </div>
            {% endfor %}
        </div>

        <div class="btn-row">
            <a class="btn-tool btn-primary" href="/class/attendance">
                📝 今日点名
            </a>
        </div>

        <div class="btn-row">
            <a class="btn-tool btn-success" href="/class/students">
                👧 学生名单
            </a>

            <a class="btn-tool btn-warning" href="/class/reports">
                📊 出席统计
            </a>
        </div>

        <div class="btn-row">
            <a class="btn-tool btn-purple" href="/class/records">
                📅 点名记录
            </a>
        </div>
                                  
        <div class="btn-row">
            <a class="btn-tool btn-warning" href="/class/promote">
                🎓 年度升班
            </a>
        </div>
                                  
    </div>

</div>
</body>
</html>
""",
        today_str=today_str,
        group_stats=group_stats
    )


@dharma_class_bp.route("/attendance", methods=["GET", "POST"])
def class_attendance():

    selected_date = request.values.get("class_date") or date.today().isoformat()
    selected_group_id = request.values.get("group_id")

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select id, name
                from dharma_class_groups
                where is_active = true
                order by sort_order, id
            """)
            groups = cur.fetchall()

            if not selected_group_id and groups:
                selected_group_id = str(groups[0]["id"])

            if request.method == "POST":

                student_ids = request.form.getlist("student_id")
                marked_by = request.form.get("marked_by", "").strip() or "老师"

                for sid in student_ids:
                    status = request.form.get(f"status_{sid}", "present")
                    remark = request.form.get(f"remark_{sid}", "").strip()

                    cur.execute("""
                        insert into dharma_attendance
                            (branch, class_date, student_id, group_id, status, remark, marked_by)
                        values
                            ('CHE', %s, %s, %s, %s, %s, %s)
                        on conflict (class_date, student_id)
                        do update set
                            status = excluded.status,
                            remark = excluded.remark,
                            marked_by = excluded.marked_by,
                            marked_at = now()
                    """, (
                        selected_date,
                        sid,
                        selected_group_id,
                        status,
                        remark,
                        marked_by
                    ))

                conn.commit()
                flash("点名已保存。", "good")
                return redirect(url_for(
                    "dharma_class.class_attendance",
                    class_date=selected_date,
                    group_id=selected_group_id
                ))

            cur.execute("""
                select
                    s.id,
                    s.student_no,
                    s.name,
                    s.english_name,
                    coalesce(a.status, 'present') as attendance_status,
                    coalesce(a.remark, '') as attendance_remark
                from dharma_students s
                left join dharma_attendance a
                    on a.student_id = s.id
                   and a.class_date = %s
                where s.status = 'active'
                  and s.group_id = %s
                order by s.student_no nulls last, s.name
            """, (selected_date, selected_group_id))

            students = cur.fetchall()

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>佛学班点名</title>
<link rel="stylesheet" href="/static/css/toolbox.css">

<style>
.attendance-row{
    border:1px solid #e5e7eb;
    border-radius:16px;
    padding:16px;
    margin-bottom:14px;
    background:#fff;
}
.student-name{
    font-size:22px;
    font-weight:700;
    margin-bottom:10px;
}
.status-grid{
    display:grid;
    grid-template-columns:repeat(4, 1fr);
    gap:8px;
    margin-bottom:10px;
}
.status-option{
    border:1px solid #ddd;
    border-radius:14px;
    padding:10px;
    text-align:center;
    font-size:18px;
    background:#f8f9fa;
}
.status-option input{
    transform:scale(1.3);
    margin-right:6px;
}
.quick-actions{
    display:flex;
    gap:10px;
    margin-bottom:18px;
}
.quick-actions button{
    flex:1;
}
@media(max-width:600px){
    .status-grid{
        grid-template-columns:repeat(2, 1fr);
    }
}
</style>
</head>

<body>
<div class="page">

    <div class="card">
        <h1 class="page-title">📝 佛学班点名</h1>
        <p class="page-subtitle">选择日期和组别后，为学生记录出席。</p>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, msg in messages %}
                    <div class="alert">{{ msg }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <form method="get">
            <div class="form-group">
                <label class="form-label">日期</label>
                <input class="form-input" type="date" name="class_date" value="{{ selected_date }}">
            </div>

            <div class="form-group">
                <label class="form-label">组别</label>
                <select class="form-input" name="group_id">
                    {% for g in groups %}
                        <option value="{{ g.id }}"
                            {% if selected_group_id|string == g.id|string %}selected{% endif %}>
                            {{ g.name }}
                        </option>
                    {% endfor %}
                </select>
            </div>

            <div class="btn-row">
                <button class="btn-tool btn-primary" type="submit">
                    🔍 查看学生
                </button>
            </div>
        </form>
    </div>

    <div class="card">
        <h2 class="section-title">学生出席</h2>

        {% if students %}
        <form method="post">
            <input type="hidden" name="class_date" value="{{ selected_date }}">
            <input type="hidden" name="group_id" value="{{ selected_group_id }}">

            <div class="form-group">
                <label class="form-label">点名老师</label>
                <input class="form-input" name="marked_by" placeholder="例如：陈老师">
            </div>
                                  
            <div class="quick-actions">
                <button class="btn-tool btn-success" type="button" onclick="markAllPresent()">
                    ✅ 全部出席
                </button>
            </div>

            {% for s in students %}
                <div class="attendance-row">
                    <input type="hidden" name="student_id" value="{{ s.id }}">

                    <div class="student-name">
                        {{ s.name }}
                        {% if s.english_name %}
                            <span style="color:#777;font-weight:400;">{{ s.english_name }}</span>
                        {% endif %}
                    </div>

                    <div class="status-grid">
                        {% for key, label in status_labels.items() %}
                            <label class="status-option">
                                <input type="radio"
                                       name="status_{{ s.id }}"
                                       value="{{ key }}"
                                       {% if s.attendance_status == key %}checked{% endif %}>
                                {{ label }}
                            </label>
                        {% endfor %}
                    </div>

                    <input class="form-input"
                           name="remark_{{ s.id }}"
                           value="{{ s.attendance_remark }}"
                           placeholder="备注，可不填">
                </div>
            {% endfor %}

            <div class="btn-row">
                <button class="btn-tool btn-success" type="submit">
                    ✅ 保存点名
                </button>
            </div>
        </form>

        {% else %}
            <div class="empty-state">
                这个组别还没有学生。下一步我们会做「学生名单」页面。
            </div>
        {% endif %}

        <div class="btn-row">
            <a class="btn-tool btn-secondary" href="/class">
                ⬅ 返回佛学班首页
            </a>
        </div>
    </div>
                                  
<script>
function markAllPresent(){
    document.querySelectorAll('input[type="radio"][value="present"]').forEach(function(r){
        r.checked = true;
    });
}
</script>

</div>
</body>
</html>
""",
        groups=groups,
        students=students,
        selected_date=selected_date,
        selected_group_id=selected_group_id,
        status_labels=STATUS_LABELS
    )

@dharma_class_bp.route("/students")
def class_students():

    group_id = request.args.get("group_id")
    q = request.args.get("q", "").strip()

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select id, name
                from dharma_class_groups
                where is_active = true
                order by sort_order, id
            """)
            groups = cur.fetchall()

            params = []
            where_sql = "where s.status = 'active'"
            params = []

            if group_id:
                where_sql += " and s.group_id = %s"
                params.append(group_id)

            if q:
                like = f"%{q}%"
                where_sql += """
                    and (
                        s.name ilike %s
                        or s.english_name ilike %s
                        or s.parent_phone ilike %s
                        or s.student_no ilike %s
                    )
                """
                params.extend([like, like, like, like])

            if group_id:
                where_sql += " and s.group_id = %s"
                params.append(group_id)

            cur.execute(f"""
                select
                    s.id,
                    s.student_no,
                    s.name,
                    s.english_name,
                    s.parent_phone,
                    s.status,
                    s.remark,
                    g.name as group_name
                from dharma_students s
                left join dharma_class_groups g on g.id = s.group_id
                {where_sql}
                order by g.sort_order, s.student_no nulls last, s.name
            """, params)

            students = cur.fetchall()

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>学生名单</title>
<link rel="stylesheet" href="/static/css/toolbox.css">
</head>

<body>
<div class="page">

    <div class="card">
        <h1 class="page-title">👧 学生名单</h1>
        <p class="page-subtitle">管理佛学班学生资料。</p>

        <form method="get">

            <div class="form-group">
                <label class="form-label">搜索学生</label>
                <input class="form-input"
                    name="q"
                    value="{{ q or '' }}"
                    placeholder="姓名 / 英文名 / 电话 / 编号">
            </div>

            <div class="form-group">
                <label class="form-label">组别筛选</label>
                <select class="form-input" name="group_id">
                    <option value="">全部组别</option>
                    {% for g in groups %}
                        <option value="{{ g.id }}"
                            {% if group_id|string == g.id|string %}selected{% endif %}>
                            {{ g.name }}
                        </option>
                    {% endfor %}
                </select>
            </div>

            <div class="btn-row">
                <button class="btn-tool btn-primary" type="submit">
                    🔍 搜索
                </button>
            </div>

        </form>

        <div class="btn-row">
            <a class="btn-tool btn-success" href="/class/students/add">
                ➕ 新增学生
            </a>
        </div>
                                  
        <div class="btn-row">
            <a class="btn-tool btn-purple" href="/class/students/import">
                📥 导入学生 Excel
            </a>

            <a class="btn-tool btn-secondary" href="/class/students/template">
                📄 下载模板
            </a>
        </div>

        <div class="btn-row">
            <a class="btn-tool btn-warning" href="/class/students/export">
                📤 导出学生名单
            </a>
        </div>
                                  
        <div class="btn-row">
            <a class="btn-tool btn-purple" href="/class/students/import">
                📥 导入学生 Excel
            </a>
        </div>
                                  
    </div>

    <div class="card">
        <h2 class="section-title">学生列表</h2>

        {% if students %}
            <div class="table-responsive">
                <table class="record-table">
                    <thead>
                        <tr>
                            <th>编号</th>
                            <th>姓名</th>
                            <th>英文名</th>
                            <th>组别</th>
                            <th>家长电话</th>
                            <th>备注</th>
                            <th>操作</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for s in students %}
                        <tr>
                            <td>{{ s.student_no or "-" }}</td>
                            <td>{{ s.name }}</td>
                            <td>{{ s.english_name or "-" }}</td>
                            <td>{{ s.group_name or "-" }}</td>
                            <td>{{ s.parent_phone or "-" }}</td>
                            <td>{{ s.remark or "-" }}</td>
                            <td>
                                <a class="btn-tool btn-warning"
                                style="font-size:16px; min-height:40px; padding:8px 12px;"
                                href="/class/students/edit/{{ s.id }}">
                                    ✏ 编辑
                                </a>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        {% else %}
            <div class="empty-state">
                还没有学生资料。
            </div>
        {% endif %}

        <div class="btn-row">
            <a class="btn-tool btn-secondary" href="/class">
                ⬅ 返回佛学班首页
            </a>
        </div>
    </div>

</div>
</body>
</html>
""",
        groups=groups,
        students=students,
        group_id=group_id,
        q=q
    )

@dharma_class_bp.route("/students/add", methods=["GET", "POST"])
def class_students_add():

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select id, name
                from dharma_class_groups
                where is_active = true
                order by sort_order, id
            """)
            groups = cur.fetchall()

            if request.method == "POST":
                student_no = request.form.get("student_no", "").strip()
                name = request.form.get("name", "").strip()
                english_name = request.form.get("english_name", "").strip()
                parent_name = request.form.get("parent_name", "").strip()
                parent_phone = request.form.get("parent_phone", "").strip()
                group_id = request.form.get("group_id")
                remark = request.form.get("remark", "").strip()

                if not name:
                    flash("学生姓名必须填写。", "bad")
                else:
                    cur.execute("""
                        insert into dharma_students
                            (branch, student_no, name, english_name,parent_name, parent_phone, group_id, status, remark)
                        values
                            ('CHE', %s, %s, %s, %s, %s, %s, 'active', %s)
                    """, (
                        student_no or None,
                        name,
                        english_name or None,
                        parent_name or None,
                        parent_phone or None,
                        group_id,
                        remark or None
                    ))

                    conn.commit()
                    flash("学生已新增。", "good")
                    return redirect(url_for("dharma_class.class_students"))

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>新增学生</title>
<link rel="stylesheet" href="/static/css/toolbox.css">
</head>

<body>
<div class="page">

    <div class="card">
        <h1 class="page-title">➕ 新增学生</h1>
        <p class="page-subtitle">加入佛学班学生资料。</p>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, msg in messages %}
                    <div class="alert">{{ msg }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <form method="post">

            <div class="form-group">
                <label class="form-label">学生编号</label>
                <input class="form-input" name="student_no" placeholder="可不填，例如：S001">
            </div>

            <div class="form-group">
                <label class="form-label">中文姓名</label>
                <input class="form-input" name="name" required>
            </div>

            <div class="form-group">
                <label class="form-label">英文名</label>
                <input class="form-input" name="english_name">
            </div>
                                  
            <div class="form-group">
                <label class="form-label">家长姓名</label>
                <input class="form-input" name="parent_name">
            </div>

            <div class="form-group">
                <label class="form-label">家长电话</label>
                <input class="form-input" name="parent_phone">
            </div>

            <div class="form-group">
                <label class="form-label">组别</label>
                <select class="form-input" name="group_id" required>
                    {% for g in groups %}
                        <option value="{{ g.id }}">{{ g.name }}</option>
                    {% endfor %}
                </select>
            </div>

            <div class="form-group">
                <label class="form-label">备注</label>
                <textarea class="form-input" name="remark" rows="3"></textarea>
            </div>

            <div class="btn-row">
                <button class="btn-tool btn-success" type="submit">
                    ✅ 保存学生
                </button>
            </div>

        </form>

        <div class="btn-row">
            <a class="btn-tool btn-secondary" href="/class/students">
                ⬅ 返回学生名单
            </a>
        </div>
    </div>

</div>
</body>
</html>
""", groups=groups)

@dharma_class_bp.route("/students/edit/<int:student_id>", methods=["GET", "POST"])
def class_students_edit(student_id):

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select id, name
                from dharma_class_groups
                where is_active = true
                order by sort_order, id
            """)
            groups = cur.fetchall()

            cur.execute("""
                select *
                from dharma_students
                where id = %s
            """, (student_id,))
            student = cur.fetchone()

            if not student:
                flash("找不到这个学生。", "bad")
                return redirect(url_for("dharma_class.class_students"))

            if request.method == "POST":
                action = request.form.get("action", "save")

                if action == "inactive":
                    cur.execute("""
                        update dharma_students
                        set status = 'inactive'
                        where id = %s
                    """, (student_id,))
                    conn.commit()
                    flash("学生已停用。", "good")
                    return redirect(url_for("dharma_class.class_students"))

                student_no = request.form.get("student_no", "").strip()
                name = request.form.get("name", "").strip()
                english_name = request.form.get("english_name", "").strip()
                parent_name = request.form.get("parent_name", "").strip()
                parent_phone = request.form.get("parent_phone", "").strip()
                group_id = request.form.get("group_id")
                remark = request.form.get("remark", "").strip()

                if not name:
                    flash("学生姓名必须填写。", "bad")
                else:
                    cur.execute("""
                        update dharma_students
                        set
                            student_no = %s,
                            name = %s,
                            english_name = %s,
                            parent_name = %s,
                            parent_phone = %s,
                            group_id = %s,
                            remark = %s
                        where id = %s
                    """, (
                        student_no or None,
                        name,
                        english_name or None,
                        parent_name or None,
                        parent_phone or None,
                        group_id,
                        remark or None,
                        student_id
                    ))

                    conn.commit()
                    flash("学生资料已更新。", "good")
                    return redirect(url_for("dharma_class.class_students"))

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>编辑学生</title>
<link rel="stylesheet" href="/static/css/toolbox.css">
</head>

<body>
<div class="page">

    <div class="card">
        <h1 class="page-title">✏ 编辑学生</h1>
        <p class="page-subtitle">修改学生资料或停用学生。</p>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, msg in messages %}
                    <div class="alert">{{ msg }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <form method="post">

            <div class="form-group">
                <label class="form-label">学生编号</label>
                <input class="form-input" name="student_no" value="{{ student.student_no or '' }}">
            </div>

            <div class="form-group">
                <label class="form-label">中文姓名</label>
                <input class="form-input" name="name" value="{{ student.name }}" required>
            </div>

            <div class="form-group">
                <label class="form-label">英文名</label>
                <input class="form-input" name="english_name" value="{{ student.english_name or '' }}">
            </div>
                                  
            <div class="form-group">
                <label class="form-label">家长姓名</label>
                <input class="form-input" name="parent_name" value="{{ student.parent_name or '' }}">
            </div>

            <div class="form-group">
                <label class="form-label">家长电话</label>
                <input class="form-input" name="parent_phone" value="{{ student.parent_phone or '' }}">
            </div>

            <div class="form-group">
                <label class="form-label">组别</label>
                <select class="form-input" name="group_id" required>
                    {% for g in groups %}
                        <option value="{{ g.id }}"
                            {% if student.group_id|string == g.id|string %}selected{% endif %}>
                            {{ g.name }}
                        </option>
                    {% endfor %}
                </select>
            </div>

            <div class="form-group">
                <label class="form-label">备注</label>
                <textarea class="form-input" name="remark" rows="3">{{ student.remark or '' }}</textarea>
            </div>

            <div class="btn-row">
                <button class="btn-tool btn-success" type="submit" name="action" value="save">
                    ✅ 保存修改
                </button>
            </div>

            <div class="btn-row">
                <button class="btn-tool btn-danger" type="submit" name="action" value="inactive"
                        onclick="return confirm('确定要停用这个学生吗？停用后点名不会再显示。');">
                    🗑 停用学生
                </button>
            </div>

        </form>

        <div class="btn-row">
            <a class="btn-tool btn-secondary" href="/class/students">
                ⬅ 返回学生名单
            </a>
        </div>
    </div>

</div>
</body>
</html>
""",
        student=student,
        groups=groups
    )

@dharma_class_bp.route("/reports")
def class_reports():

    today = date.today()
    ym = request.args.get("ym") or today.strftime("%Y-%m")

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select id, name
                from dharma_class_groups
                where is_active = true
                order by sort_order, id
            """)
            groups = cur.fetchall()

            cur.execute("""
                select
                    g.name as group_name,
                    s.id as student_id,
                    s.student_no,
                    s.name,
                    s.english_name,
                    count(a.id) as total_records,
                    sum(case when a.status = 'present' then 1 else 0 end) as present_count,
                    sum(case when a.status = 'late' then 1 else 0 end) as late_count,
                    sum(case when a.status = 'leave' then 1 else 0 end) as leave_count,
                    sum(case when a.status = 'absent' then 1 else 0 end) as absent_count
                from dharma_students s
                left join dharma_class_groups g on g.id = s.group_id
                left join dharma_attendance a
                    on a.student_id = s.id
                   and to_char(a.class_date, 'YYYY-MM') = %s
                where s.status = 'active'
                group by g.name, g.sort_order, s.id
                order by g.sort_order, s.student_no nulls last, s.name
            """, (ym,))

            rows = cur.fetchall()

    for r in rows:
        total = r["total_records"] or 0
        present = (r["present_count"] or 0) + (r["late_count"] or 0)

        if total > 0:
            r["rate"] = round(present * 100 / total, 1)
        else:
            r["rate"] = None

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>出席统计</title>
<link rel="stylesheet" href="/static/css/toolbox.css">
</head>

<body>
<div class="page">

    <div class="card">
        <h1 class="page-title">📊 出席统计</h1>
        <p class="page-subtitle">查看学生每月出席率。</p>

        <form method="get">
            <div class="form-group">
                <label class="form-label">月份</label>
                <input class="form-input" type="month" name="ym" value="{{ ym }}">
            </div>

            <div class="btn-row">
                <button class="btn-tool btn-primary" type="submit">
                    🔍 查看统计
                </button>
            </div>
                                  
            <div class="btn-row">
                <a class="btn-tool btn-success"
                href="/class/reports/monthly/export?ym={{ ym }}">
                    📥 下载月报 Excel
                </a>
            </div>
                                  
            <div class="btn-row">
                <a class="btn-tool btn-warning"
                href="/class/reports/yearly/export?year={{ ym[:4] }}">
                    📥 下载年报 Excel
                </a>
            </div>
                                  
        </form>
    </div>

    <div class="card">
        <h2 class="section-title">{{ ym }} 出席率</h2>

        {% if rows %}
            <div class="table-responsive">
                <table class="record-table">
                    <thead>
                        <tr>
                            <th>组别</th>
                            <th>编号</th>
                            <th>姓名</th>
                            <th>出席</th>
                            <th>迟到</th>
                            <th>请假</th>
                            <th>缺席</th>
                            <th>出席率</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for r in rows %}
                        <tr>
                            <td>{{ r.group_name or "-" }}</td>
                            <td>{{ r.student_no or "-" }}</td>
                            <td>
                                {{ r.name }}
                                {% if r.english_name %}
                                    <span style="color:#777;">{{ r.english_name }}</span>
                                {% endif %}
                            </td>
                            <td>{{ r.present_count or 0 }}</td>
                            <td>{{ r.late_count or 0 }}</td>
                            <td>{{ r.leave_count or 0 }}</td>
                            <td>{{ r.absent_count or 0 }}</td>
                            <td>
                                {% if r.rate is not none %}
                                    {{ r.rate }}%
                                {% else %}
                                    -
                                {% endif %}
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        {% else %}
            <div class="empty-state">
                暂时没有学生资料。
            </div>
        {% endif %}

        <div class="btn-row">
            <a class="btn-tool btn-secondary" href="/class">
                ⬅ 返回佛学班首页
            </a>
        </div>
    </div>

</div>
</body>
</html>
""",
        ym=ym,
        rows=rows,
        groups=groups
    )

@dharma_class_bp.route("/records")
def class_records():

    selected_date = request.args.get("class_date") or date.today().isoformat()
    group_id = request.args.get("group_id")

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select id, name
                from dharma_class_groups
                where is_active = true
                order by sort_order, id
            """)
            groups = cur.fetchall()

            params = [selected_date]
            where_group = ""

            if group_id:
                where_group = "and s.group_id = %s"
                params.append(group_id)

            cur.execute(f"""
                select
                    s.student_no,
                    s.name,
                    s.english_name,
                    g.name as group_name,
                    a.status,
                    a.remark,
                    a.marked_by,
                    a.marked_at
                from dharma_attendance a
                join dharma_students s on s.id = a.student_id
                left join dharma_class_groups g on g.id = s.group_id
                where a.class_date = %s
                {where_group}
                order by g.sort_order, s.student_no nulls last, s.name
            """, params)

            records = cur.fetchall()

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>点名记录</title>
<link rel="stylesheet" href="/static/css/toolbox.css">
</head>

<body>
<div class="page">

    <div class="card">
        <h1 class="page-title">📅 点名记录</h1>
        <p class="page-subtitle">查看某一天佛学班点名情况。</p>

        <form method="get">
            <div class="form-group">
                <label class="form-label">日期</label>
                <input class="form-input" type="date" name="class_date" value="{{ selected_date }}">
            </div>

            <div class="form-group">
                <label class="form-label">组别</label>
                <select class="form-input" name="group_id">
                    <option value="">全部组别</option>
                    {% for g in groups %}
                        <option value="{{ g.id }}"
                            {% if group_id|string == g.id|string %}selected{% endif %}>
                            {{ g.name }}
                        </option>
                    {% endfor %}
                </select>
            </div>

            <div class="btn-row">
                <button class="btn-tool btn-primary" type="submit">
                    🔍 查看记录
                </button>
            </div>
        </form>
    </div>

    <div class="card">
        <h2 class="section-title">{{ selected_date }} 点名记录</h2>

        {% if records %}
            <div class="table-responsive">
                <table class="record-table">
                    <thead>
                        <tr>
                            <th>组别</th>
                            <th>编号</th>
                            <th>姓名</th>
                            <th>状态</th>
                            <th>备注</th>
                            <th>老师</th>
                            <th>保存时间</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for r in records %}
                        <tr>
                            <td>{{ r.group_name or "-" }}</td>
                            <td>{{ r.student_no or "-" }}</td>
                            <td>
                                {{ r.name }}
                                {% if r.english_name %}
                                    <span style="color:#777;">{{ r.english_name }}</span>
                                {% endif %}
                            </td>
                            <td>
                                {% if r.status == "present" %}
                                    ✅ 出席
                                {% elif r.status == "late" %}
                                    🟡 迟到
                                {% elif r.status == "leave" %}
                                    🟠 请假
                                {% elif r.status == "absent" %}
                                    ❌ 缺席
                                {% else %}
                                    {{ r.status }}
                                {% endif %}
                            </td>
                            <td>{{ r.remark or "-" }}</td>
                            <td>{{ r.marked_by or "-" }}</td>
                            <td>{{ r.marked_at or "-" }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        {% else %}
            <div class="empty-state">
                这一天还没有点名记录。
            </div>
        {% endif %}

        <div class="btn-row">
            <a class="btn-tool btn-secondary" href="/class">
                ⬅ 返回佛学班首页
            </a>
        </div>
    </div>

</div>
</body>
</html>
""",
        selected_date=selected_date,
        group_id=group_id,
        groups=groups,
        records=records
    )

@dharma_class_bp.route("/reports/monthly/export")
def class_export_monthly_report():

    ym = request.args.get("ym") or date.today().strftime("%Y-%m")

    status_symbol = {
        "present": "✓",
        "late": "迟",
        "leave": "请",
        "absent": "✗",
    }

    status_fill = {
        "✓": PatternFill("solid", fgColor="C6EFCE"),
        "迟": PatternFill("solid", fgColor="FCE4D6"),
        "请": PatternFill("solid", fgColor="FFF2CC"),
        "✗": PatternFill("solid", fgColor="F4CCCC"),
    }

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select id, name, sort_order
                from dharma_class_groups
                where is_active = true
                order by sort_order, id
            """)
            groups = cur.fetchall()

            cur.execute("""
                select distinct class_date
                from dharma_attendance
                where to_char(class_date, 'YYYY-MM') = %s
                order by class_date
            """, (ym,))
            class_dates = [r["class_date"] for r in cur.fetchall()]

            cur.execute("""
                select
                    s.id,
                    s.student_no,
                    s.name,
                    s.english_name,
                    s.group_id,
                    g.name as group_name,
                    g.sort_order
                from dharma_students s
                join dharma_class_groups g on g.id = s.group_id
                where s.status = 'active'
                  and g.is_active = true
                order by g.sort_order, s.student_no nulls last, s.name
            """)
            students = cur.fetchall()

            cur.execute("""
                select student_id, class_date, status
                from dharma_attendance
                where to_char(class_date, 'YYYY-MM') = %s
            """, (ym,))
            att_rows = cur.fetchall()

    attendance_map = {}
    for r in att_rows:
        attendance_map[(r["student_id"], r["class_date"])] = r["status"]

    wb = Workbook()

    # ===== 样式 =====
    title_fill = PatternFill("solid", fgColor="1F4E78")
    header_fill = PatternFill("solid", fgColor="5B9BD5")
    white_font = Font(color="FFFFFF", bold=True)
    title_font = Font(size=18, color="FFFFFF", bold=True)
    bold_font = Font(bold=True)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    # ===== 总览 Sheet =====
    ws = wb.active
    ws.title = "总览"

    ws.merge_cells("A1:H1")
    ws["A1"] = f"蕉赖佛学班出席月报  {ym}"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 34

    ws["A3"] = "本月课程日期"
    ws["A3"].font = bold_font

    if class_dates:
        date_text = "、".join([f"{d.month}/{d.day}" for d in class_dates])
    else:
        date_text = "本月还没有点名记录"

    ws["B3"] = date_text

    ws.append([])
    ws.append(["组别", "总学生人数", "课程次数", "平均出席率", "备注"])

    for cell in ws[5]:
        cell.fill = header_fill
        cell.font = white_font
        cell.alignment = center
        cell.border = border

    for g in groups:
        group_students = [s for s in students if s["group_id"] == g["id"]]
        total_student = len(group_students)

        rates = []

        for s in group_students:
            total = 0
            attended = 0

            for d in class_dates:
                st = attendance_map.get((s["id"], d))
                if st:
                    total += 1
                    if st in ("present", "late"):
                        attended += 1

            if total > 0:
                rates.append(attended / total)

        avg_rate = sum(rates) / len(rates) if rates else None

        ws.append([
            g["name"],
            total_student,
            len(class_dates),
            avg_rate,
            ""
        ])

    for row in ws.iter_rows(min_row=6, max_row=ws.max_row):
        for cell in row:
            cell.border = border
            cell.alignment = center

        rate_cell = row[3]
        if rate_cell.value is not None:
            rate_cell.number_format = "0%"
            if rate_cell.value >= 0.9:
                rate_cell.fill = PatternFill("solid", fgColor="C6EFCE")
            elif rate_cell.value >= 0.75:
                rate_cell.fill = PatternFill("solid", fgColor="FFEB9C")
            else:
                rate_cell.fill = PatternFill("solid", fgColor="F4CCCC")

    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:E{ws.max_row}"

    # ===== 各组 Sheet =====
    for g in groups:
        ws_g = wb.create_sheet(g["name"][:31])

        headers = ["学生编号", "学生姓名"]

        for d in class_dates:
            headers.append(f"{d.month}/{d.day}")

        headers += ["出席", "迟到", "请假", "缺席", "出席率"]

        end_col = max(8, len(headers))

        ws_g.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=end_col
        )

        ws_g["A1"] = f"{g['name']} 出席月报  {ym}"
        ws_g["A1"].font = title_font
        ws_g["A1"].fill = title_fill
        ws_g["A1"].alignment = center
        ws_g.row_dimensions[1].height = 34

        ws_g.append([])
        ws_g.append(headers)

        for cell in ws_g[3]:
            cell.fill = header_fill
            cell.font = white_font
            cell.alignment = center
            cell.border = border

        group_students = [s for s in students if s["group_id"] == g["id"]]

        if not group_students:
            ws_g.append(["", "目前没有学生资料"])

        for s in group_students:
            present_count = 0
            late_count = 0
            leave_count = 0
            absent_count = 0
            total_count = 0

            row = [
                s["student_no"] or "",
                s["name"] or "",
            ]

            for d in class_dates:
                st = attendance_map.get((s["id"], d))

                if st:
                    total_count += 1

                if st == "present":
                    present_count += 1
                elif st == "late":
                    late_count += 1
                elif st == "leave":
                    leave_count += 1
                elif st == "absent":
                    absent_count += 1

                row.append(status_symbol.get(st, ""))

            attended = present_count + late_count
            rate = attended / total_count if total_count > 0 else None

            row += [
                present_count,
                late_count,
                leave_count,
                absent_count,
                rate
            ]

            ws_g.append(row)

        for row in ws_g.iter_rows(min_row=4, max_row=ws_g.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = center

                if cell.value in status_fill:
                    cell.fill = status_fill[cell.value]
                    cell.font = Font(bold=True)

            rate_cell = row[-1]
            if rate_cell.value is not None and isinstance(rate_cell.value, (int, float)):
                rate_cell.number_format = "0%"

                if rate_cell.value >= 0.9:
                    rate_cell.fill = PatternFill("solid", fgColor="C6EFCE")
                elif rate_cell.value >= 0.75:
                    rate_cell.fill = PatternFill("solid", fgColor="FFEB9C")
                else:
                    rate_cell.fill = PatternFill("solid", fgColor="F4CCCC")

        ws_g.freeze_panes = "C4"
        ws_g.auto_filter.ref = f"A3:{get_column_letter(ws_g.max_column)}{ws_g.max_row}"

        ws_g.column_dimensions["A"].width = 14
        ws_g.column_dimensions["B"].width = 22

    # ===== 全部 Sheet 自动美化 =====
    for sheet in wb.worksheets:

        for col_idx in range(1, sheet.max_column + 1):
            col_letter = get_column_letter(col_idx)

            max_len = 0
            for row_idx in range(1, sheet.max_row + 1):
                cell = sheet.cell(row=row_idx, column=col_idx)
                value = str(cell.value) if cell.value is not None else ""
                max_len = max(max_len, len(value))

            width = max_len + 4

            if col_idx == 1:
                width = max(width, 14)
            elif col_idx == 2:
                width = max(width, 22)
            else:
                width = max(width, 10)

            sheet.column_dimensions[col_letter].width = width

        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(
                    horizontal=cell.alignment.horizontal or "center",
                    vertical="center"
                )

        for i in range(1, sheet.max_row + 1):
            sheet.row_dimensions[i].height = 24

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"蕉赖佛学班出席月报_{ym}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@dharma_class_bp.route("/students/import", methods=["GET", "POST"])
def class_students_import():

    result = None

    if request.method == "POST":
        file = request.files.get("excel_file")

        if not file or file.filename == "":
            flash("请选择 Excel 文件。", "bad")
            return redirect(url_for("dharma_class.class_students_import"))

        wb = load_workbook(file, data_only=True)
        ws = wb.active

        inserted = 0
        duplicate = 0
        blank = 0
        invalid = 0

        inserted_names = []
        duplicate_names = []
        error_messages = []

        with get_conn() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:

                cur.execute("""
                    select id, name
                    from dharma_class_groups
                    where is_active = true
                """)
                groups = cur.fetchall()

                group_map = {
                    g["name"].strip(): g["id"]
                    for g in groups
                }

                for row_no, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):

                    student_no = str(row[0]).strip() if len(row) > 0 and row[0] else ""
                    name = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                    english_name = str(row[2]).strip() if len(row) > 2 and row[2] else ""
                    parent_name = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                    parent_phone = str(row[4]).strip() if len(row) > 4 and row[4] else ""
                    group_name = str(row[5]).strip() if len(row) > 5 and row[5] else ""
                    remark = str(row[6]).strip() if len(row) > 6 and row[6] else ""

                    if not name:
                        blank += 1
                        continue

                    group_id = group_map.get(group_name)

                    if not group_id:
                        invalid += 1
                        error_messages.append(
                            f"第 {row_no} 行：组别「{group_name}」不存在"
                        )
                        continue

                    # 重复检查：优先学生编号，其次 中文姓名 + 家长电话
                    if student_no:
                        cur.execute("""
                            select id
                            from dharma_students
                            where student_no = %s
                            limit 1
                        """, (student_no,))
                    else:
                        cur.execute("""
                            select id
                            from dharma_students
                            where name = %s
                              and coalesce(parent_phone, '') = %s
                            limit 1
                        """, (name, parent_phone or ""))

                    exists = cur.fetchone()

                    if exists:
                        duplicate += 1
                        duplicate_names.append(name)
                        continue

                    cur.execute("""
                        insert into dharma_students
                            (branch, student_no, name, english_name, parent_name,
                             parent_phone, group_id, status, remark)
                        values
                            ('CHE', %s, %s, %s, %s, %s, %s, 'active', %s)
                    """, (
                        student_no or None,
                        name,
                        english_name or None,
                        parent_name or None,
                        parent_phone or None,
                        group_id,
                        remark or None
                    ))

                    inserted += 1
                    inserted_names.append(name)

                conn.commit()

        result = {
            "inserted": inserted,
            "duplicate": duplicate,
            "blank": blank,
            "invalid": invalid,
            "inserted_names": inserted_names,
            "duplicate_names": duplicate_names,
            "error_messages": error_messages,
        }

        flash("导入完成。", "good")

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>导入学生 Excel</title>
<link rel="stylesheet" href="/static/css/toolbox.css">
</head>

<body>
<div class="page">

    <div class="card">
        <h1 class="page-title">📥 导入学生 Excel</h1>
        <p class="page-subtitle">上传学生名单，系统会自动新增并检查重复。</p>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, msg in messages %}
                    <div class="alert">{{ msg }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <div class="empty-state" style="text-align:left;">
            Excel 第一行标题请使用：<br><br>
            学生编号｜中文姓名｜英文名｜家长姓名｜家长电话｜组别｜备注<br><br>
            组别目前只接受：低年组、少年组、高年组
        </div>

        <div class="btn-row">
            <a class="btn-tool btn-secondary" href="/class/students/template">
                📄 下载 Excel 模板
            </a>
        </div>

        <form method="post" enctype="multipart/form-data">

            <div class="form-group">
                <label class="form-label">选择 Excel 文件</label>
                <input class="form-input" type="file" name="excel_file" accept=".xlsx" required>
            </div>

            <div class="btn-row">
                <button class="btn-tool btn-success" type="submit">
                    ✅ 上传并导入
                </button>
            </div>

        </form>
    </div>

    {% if result %}
    <div class="card">
        <h2 class="section-title">📋 导入结果</h2>

        <div class="summary-grid">
            <div class="summary-box">
                <div>新增</div>
                <div style="font-size:30px;font-weight:800;">{{ result.inserted }}</div>
            </div>

            <div class="summary-box">
                <div>重复</div>
                <div style="font-size:30px;font-weight:800;">{{ result.duplicate }}</div>
            </div>

            <div class="summary-box">
                <div>空白</div>
                <div style="font-size:30px;font-weight:800;">{{ result.blank }}</div>
            </div>

            <div class="summary-box">
                <div>错误</div>
                <div style="font-size:30px;font-weight:800;">{{ result.invalid }}</div>
            </div>
        </div>

        {% if result.inserted_names %}
            <h3 class="section-title">✅ 新增学生</h3>
            <div class="empty-state" style="text-align:left;">
                {% for n in result.inserted_names[:30] %}
                    ✓ {{ n }}<br>
                {% endfor %}
            </div>
        {% endif %}

        {% if result.duplicate_names %}
            <h3 class="section-title">⚠️ 重复已跳过</h3>
            <div class="empty-state" style="text-align:left;">
                {% for n in result.duplicate_names[:30] %}
                    • {{ n }}<br>
                {% endfor %}
            </div>
        {% endif %}

        {% if result.error_messages %}
            <h3 class="section-title">❌ 错误资料</h3>
            <div class="empty-state" style="text-align:left;color:#b91c1c;">
                {% for e in result.error_messages[:30] %}
                    {{ e }}<br>
                {% endfor %}
            </div>
        {% endif %}
    </div>
    {% endif %}

    <div class="card">
        <div class="btn-row">
            <a class="btn-tool btn-secondary" href="/class/students">
                ⬅ 返回学生名单
            </a>
        </div>
    </div>

</div>
</body>
</html>
""", result=result)

@dharma_class_bp.route("/students/template")
def class_students_template():

    wb = Workbook()
    ws = wb.active
    ws.title = "学生导入模板"

    headers = [
        "学生编号",
        "中文姓名",
        "英文名",
        "家长姓名",
        "家长电话",
        "组别",
        "备注"
    ]

    ws.append(headers)

    ws.append([
        "001",
        "张小明",
        "Zhang Xiao Ming",
        "张爸爸",
        "0123456789",
        "低年组",
        ""
    ])

    ws.append([
        "002",
        "李小美",
        "Lee Mei Mei",
        "李妈妈",
        "01122334455",
        "少年组",
        ""
    ])

    beautify_simple_excel(ws)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="佛学班学生导入模板.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@dharma_class_bp.route("/students/export")
def class_students_export():

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                select
                    s.student_no,
                    s.name,
                    s.english_name,
                    s.parent_name,
                    s.parent_phone,
                    g.name as group_name,
                    s.status,
                    s.remark
                from dharma_students s
                left join dharma_class_groups g on g.id = s.group_id
                order by g.sort_order, s.student_no nulls last, s.name
            """)
            rows = cur.fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "学生名单"

    ws.append([
        "学生编号",
        "中文姓名",
        "英文名",
        "家长姓名",
        "家长电话",
        "组别",
        "状态",
        "备注"
    ])

    for r in rows:
        ws.append([
            r["student_no"] or "",
            r["name"] or "",
            r["english_name"] or "",
            r["parent_name"] or "",
            r["parent_phone"] or "",
            r["group_name"] or "",
            r["status"] or "",
            r["remark"] or ""
        ])

    beautify_simple_excel(ws)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="佛学班学生名单.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@dharma_class_bp.route("/reports/yearly/export")
def class_export_yearly_report():

    year = request.args.get("year") or date.today().strftime("%Y")

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select id, name, sort_order
                from dharma_class_groups
                where is_active = true
                order by sort_order, id
            """)
            groups = cur.fetchall()

            cur.execute("""
                select
                    s.id,
                    s.student_no,
                    s.name,
                    s.english_name,
                    s.group_id,
                    g.name as group_name,
                    g.sort_order
                from dharma_students s
                join dharma_class_groups g on g.id = s.group_id
                where s.status = 'active'
                  and g.is_active = true
                order by g.sort_order, s.student_no nulls last, s.name
            """)
            students = cur.fetchall()

            cur.execute("""
                select
                    student_id,
                    to_char(class_date, 'MM') as month_no,
                    status
                from dharma_attendance
                where to_char(class_date, 'YYYY') = %s
            """, (year,))
            att_rows = cur.fetchall()

    # student_id -> month -> counters
    stat = {}

    for r in att_rows:
        sid = r["student_id"]
        m = int(r["month_no"])
        status = r["status"]

        stat.setdefault(sid, {})
        stat[sid].setdefault(m, {
            "total": 0,
            "attended": 0,
            "present": 0,
            "late": 0,
            "leave": 0,
            "absent": 0,
        })

        stat[sid][m]["total"] += 1

        if status == "present":
            stat[sid][m]["present"] += 1
            stat[sid][m]["attended"] += 1
        elif status == "late":
            stat[sid][m]["late"] += 1
            stat[sid][m]["attended"] += 1
        elif status == "leave":
            stat[sid][m]["leave"] += 1
        elif status == "absent":
            stat[sid][m]["absent"] += 1

    wb = Workbook()

    title_fill = PatternFill("solid", fgColor="1F4E78")
    header_fill = PatternFill("solid", fgColor="5B9BD5")
    white_font = Font(color="FFFFFF", bold=True)
    title_font = Font(size=18, color="FFFFFF", bold=True)
    bold_font = Font(bold=True)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    center = Alignment(horizontal="center", vertical="center")

    # ===== 总览 Sheet =====
    ws = wb.active
    ws.title = "总览"

    ws.merge_cells("A1:F1")
    ws["A1"] = f"蕉赖佛学班出席年报 {year}"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 34

    ws.append([])
    ws.append(["组别", "学生人数", "全年记录", "平均出席率", "备注"])

    for cell in ws[3]:
        cell.fill = header_fill
        cell.font = white_font
        cell.alignment = center
        cell.border = border

    for g in groups:
        group_students = [s for s in students if s["group_id"] == g["id"]]
        rates = []
        total_records = 0

        for s in group_students:
            total = 0
            attended = 0

            for m in range(1, 13):
                data = stat.get(s["id"], {}).get(m)
                if data:
                    total += data["total"]
                    attended += data["attended"]

            total_records += total

            if total > 0:
                rates.append(attended / total)

        avg_rate = sum(rates) / len(rates) if rates else None

        ws.append([
            g["name"],
            len(group_students),
            total_records,
            avg_rate,
            ""
        ])

    for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
        for cell in row:
            cell.border = border
            cell.alignment = center

        rate_cell = row[3]
        if rate_cell.value is not None:
            rate_cell.number_format = "0%"

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:E{ws.max_row}"

    # ===== 各组 Sheet =====
    for g in groups:

        ws_g = wb.create_sheet(g["name"][:31])

        headers = ["学生编号", "学生姓名"]
        headers += [f"{m}月" for m in range(1, 13)]
        headers += ["全年出席率", "出席", "迟到", "请假", "缺席"]

        end_col = len(headers)

        ws_g.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=end_col
        )

        ws_g["A1"] = f"{g['name']} 出席年报 {year}"
        ws_g["A1"].font = title_font
        ws_g["A1"].fill = title_fill
        ws_g["A1"].alignment = center
        ws_g.row_dimensions[1].height = 34

        ws_g.append([])
        ws_g.append(headers)

        for cell in ws_g[3]:
            cell.fill = header_fill
            cell.font = white_font
            cell.alignment = center
            cell.border = border

        group_students = [s for s in students if s["group_id"] == g["id"]]

        if not group_students:
            ws_g.append(["", "目前没有学生资料"])

        for s in group_students:

            row = [
                s["student_no"] or "",
                s["name"] or "",
            ]

            yearly_total = 0
            yearly_attended = 0
            yearly_present = 0
            yearly_late = 0
            yearly_leave = 0
            yearly_absent = 0

            for m in range(1, 13):
                data = stat.get(s["id"], {}).get(m)

                if data and data["total"] > 0:
                    month_rate = data["attended"] / data["total"]
                    row.append(month_rate)

                    yearly_total += data["total"]
                    yearly_attended += data["attended"]
                    yearly_present += data["present"]
                    yearly_late += data["late"]
                    yearly_leave += data["leave"]
                    yearly_absent += data["absent"]
                else:
                    row.append(None)

            yearly_rate = yearly_attended / yearly_total if yearly_total > 0 else None

            row += [
                yearly_rate,
                yearly_present,
                yearly_late,
                yearly_leave,
                yearly_absent
            ]

            ws_g.append(row)

        # 美化内容
        for row in ws_g.iter_rows(min_row=4, max_row=ws_g.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = center

            # 月份 1-12 + 全年出席率
            for cell in row[2:15]:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "0%"

                    if cell.value >= 0.9:
                        cell.fill = PatternFill("solid", fgColor="C6EFCE")
                    elif cell.value >= 0.75:
                        cell.fill = PatternFill("solid", fgColor="FFEB9C")
                    else:
                        cell.fill = PatternFill("solid", fgColor="F4CCCC")

        ws_g.freeze_panes = "C4"
        ws_g.auto_filter.ref = f"A3:{get_column_letter(ws_g.max_column)}{ws_g.max_row}"

    # ===== 自动列宽 =====
    for sheet in wb.worksheets:

        for col_idx in range(1, sheet.max_column + 1):
            col_letter = get_column_letter(col_idx)

            max_len = 0
            for row_idx in range(1, sheet.max_row + 1):
                value = sheet.cell(row=row_idx, column=col_idx).value
                value = str(value) if value is not None else ""
                max_len = max(max_len, len(value))

            width = max(max_len + 4, 12)

            if col_idx == 1:
                width = max(width, 14)
            elif col_idx == 2:
                width = max(width, 22)

            sheet.column_dimensions[col_letter].width = width

        for row_idx in range(1, sheet.max_row + 1):
            sheet.row_dimensions[row_idx].height = 24

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"蕉赖佛学班出席年报_{year}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@dharma_class_bp.route("/promote", methods=["GET", "POST"])
def class_promote():

    target_year = request.values.get("target_year") or str(date.today().year + 1)
    action = request.form.get("action")

    with get_conn() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:

            cur.execute("""
                select id, name
                from dharma_class_groups
                where is_active = true
            """)
            groups = cur.fetchall()

            group_id_by_name = {g["name"]: g["id"] for g in groups}

            low_id = group_id_by_name.get("低年组")
            youth_id = group_id_by_name.get("少年组")
            high_id = group_id_by_name.get("高年组")

            promote_rules = []
            if low_id and youth_id:
                promote_rules.append((low_id, youth_id, "低年组", "少年组"))
            if youth_id and high_id:
                promote_rules.append((youth_id, high_id, "少年组", "高年组"))

            cur.execute("""
                select
                    s.id,
                    s.student_no,
                    s.name,
                    s.english_name,
                    s.group_id,
                    g.name as group_name
                from dharma_students s
                join dharma_class_groups g on g.id = s.group_id
                where s.status = 'active'
                order by g.sort_order, s.student_no nulls last, s.name
            """)
            students = cur.fetchall()

            preview_rows = []

            for s in students:
                for from_id, to_id, from_name, to_name in promote_rules:
                    if s["group_id"] == from_id:
                        preview_rows.append({
                            "id": s["id"],
                            "student_no": s["student_no"],
                            "name": s["name"],
                            "english_name": s["english_name"],
                            "from_group": from_name,
                            "to_group": to_name,
                            "to_group_id": to_id,
                        })

            promoted_count = 0

            if request.method == "POST" and action == "confirm":

                for r in preview_rows:
                    cur.execute("""
                        update dharma_students
                        set group_id = %s
                        where id = %s
                    """, (r["to_group_id"], r["id"]))

                    promoted_count += 1

                conn.commit()
                flash(f"{target_year} 年度升班完成，共更新 {promoted_count} 位学生。", "good")

                return redirect(url_for("dharma_class.class_students"))

    return render_template_string("""
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>年度升班</title>
<link rel="stylesheet" href="/static/css/toolbox.css">
</head>

<body>
<div class="page">

    <div class="card">
        <h1 class="page-title">🎓 年度升班</h1>
        <p class="page-subtitle">
            预览后确认升班。低年组升少年组，少年组升高年组，高年组保持不变。
        </p>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, msg in messages %}
                    <div class="alert">{{ msg }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <form method="get">
            <div class="form-group">
                <label class="form-label">目标学年</label>
                <input class="form-input" name="target_year" value="{{ target_year }}">
            </div>

            <div class="btn-row">
                <button class="btn-tool btn-primary" type="submit">
                    🔍 预览升班
                </button>
            </div>
        </form>
    </div>

    <div class="card">
        <h2 class="section-title">升班预览</h2>

        {% if preview_rows %}
            <div class="table-responsive">
                <table class="record-table">
                    <thead>
                        <tr>
                            <th>编号</th>
                            <th>姓名</th>
                            <th>英文名</th>
                            <th>原组别</th>
                            <th>新组别</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for r in preview_rows %}
                        <tr>
                            <td>{{ r.student_no or "-" }}</td>
                            <td>{{ r.name }}</td>
                            <td>{{ r.english_name or "-" }}</td>
                            <td>{{ r.from_group }}</td>
                            <td>{{ r.to_group }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>

            <form method="post">
                <input type="hidden" name="target_year" value="{{ target_year }}">
                <input type="hidden" name="action" value="confirm">

                <div class="btn-row">
                    <button class="btn-tool btn-danger"
                            type="submit"
                            onclick="return confirm('确定要执行年度升班吗？执行后学生组别会被更新。');">
                        ✅ 确认执行升班
                    </button>
                </div>
            </form>

        {% else %}
            <div class="empty-state">
                目前没有需要升班的学生。高年组学生会保持不变。
            </div>
        {% endif %}

        <div class="btn-row">
            <a class="btn-tool btn-secondary" href="/class">
                ⬅ 返回佛学班首页
            </a>
        </div>
    </div>

</div>
</body>
</html>
""",
        target_year=target_year,
        preview_rows=preview_rows
    )

