"""Finance audit trail helpers and read-only audit pages."""
from __future__ import annotations

from typing import Any

from flask import Blueprint, flash, redirect, render_template_string, request, session, url_for

from db import db_query
from finance_common import (
    get_current_finance_branch,
    get_current_finance_user,
    get_request_ip,
    json_safe_value,
)

finance_audit_bp = Blueprint("finance_audit", __name__, url_prefix="/finance/audit")


def write_finance_audit(
    *,
    module: str,
    action: str,
    record_id: Any = None,
    old_value: Any = None,
    new_value: Any = None,
    reason: str = "",
    branch_code: str | None = None,
    actor: str | None = None,
    safe: bool = True,
) -> bool:
    """Write one audit record. safe=True prevents audit failure breaking finance work."""
    try:
        db_query(
            """
            insert into finance_audit_logs
            (
                branch_code,
                module,
                action,
                record_id,
                actor,
                ip_address,
                reason,
                old_value,
                new_value,
                created_at
            )
            values
            (
                %s, %s, %s, %s, %s,
                %s, %s, %s::jsonb, %s::jsonb, now()
            )
            """,
            (
                branch_code or get_current_finance_branch(),
                str(module or "finance"),
                str(action or "unknown"),
                str(record_id) if record_id is not None else None,
                actor or get_current_finance_user(),
                get_request_ip(),
                str(reason or ""),
                __import__("json").dumps(json_safe_value(old_value), ensure_ascii=False)
                if old_value is not None else None,
                __import__("json").dumps(json_safe_value(new_value), ensure_ascii=False)
                if new_value is not None else None,
            ),
        )
        return True
    except Exception:
        if safe:
            return False
        raise


def _require_finance_login():
    return bool(session.get("finance_login"))


@finance_audit_bp.route("/")
def audit_list():
    if not _require_finance_login():
        return redirect(url_for("finance.finance_login"))

    q = request.args.get("q", "").strip()
    module = request.args.get("module", "").strip()
    action = request.args.get("action", "").strip()

    clauses = ["branch_code = %s"]
    params: list[Any] = [get_current_finance_branch()]

    if q:
        clauses.append("(actor ilike %s or record_id ilike %s or reason ilike %s)")
        like = f"%{q}%"
        params.extend([like, like, like])
    if module:
        clauses.append("module = %s")
        params.append(module)
    if action:
        clauses.append("action = %s")
        params.append(action)

    rows = db_query(
        f"""
        select id, branch_code, module, action, record_id,
               actor, ip_address, reason, created_at
        from finance_audit_logs
        where {' and '.join(clauses)}
        order by id desc
        limit 500
        """,
        tuple(params),
        fetchall=True,
    ) or []

    return render_template_string("""
    <!doctype html><html lang="zh"><head>
    <meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <title>财政审计记录</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}">
    </head><body><div class="page" style="max-width:1300px">
      <div class="card"><h1 class="page-title">🧾 财政审计记录</h1>
      <form method="get"><div class="form-group"><label class="form-label">关键字</label>
      <input class="form-input" name="q" value="{{ q }}" placeholder="操作人、记录编号、原因"></div>
      <div class="btn-row"><button class="btn-tool btn-primary">查询</button>
      <a class="btn-tool btn-secondary" href="{{ url_for('finance_audit.audit_list') }}">清除</a></div></form></div>
      <div class="card"><div class="table-responsive"><table class="record-table">
      <thead><tr><th>时间</th><th>模块</th><th>动作</th><th>记录</th><th>操作人</th><th>IP</th><th>原因</th><th></th></tr></thead>
      <tbody>{% for r in rows %}<tr>
      <td>{{ r.created_at }}</td><td>{{ r.module }}</td><td>{{ r.action }}</td>
      <td>{{ r.record_id or '-' }}</td><td>{{ r.actor }}</td><td>{{ r.ip_address or '-' }}</td>
      <td>{{ r.reason or '-' }}</td><td><a class="btn-tool btn-secondary" href="{{ url_for('finance_audit.audit_detail', audit_id=r.id) }}">详情</a></td>
      </tr>{% else %}<tr><td colspan="8">暂无审计记录</td></tr>{% endfor %}</tbody></table></div></div>
      <a class="btn-tool btn-secondary" href="{{ url_for('finance.finance_home') }}">← 返回财政首页</a>
    </div></body></html>
    """, rows=rows, q=q)


@finance_audit_bp.route("/<int:audit_id>")
def audit_detail(audit_id: int):
    if not _require_finance_login():
        return redirect(url_for("finance.finance_login"))

    row = db_query(
        """
        select * from finance_audit_logs
        where id = %s and branch_code = %s
        limit 1
        """,
        (audit_id, get_current_finance_branch()),
        fetchone=True,
    )
    if not row:
        flash("找不到审计记录。", "danger")
        return redirect(url_for("finance_audit.audit_list"))

    return render_template_string("""
    <!doctype html><html lang="zh"><head><meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1"><title>审计详情</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='css/toolbox.css') }}"></head>
    <body><div class="page" style="max-width:1000px"><div class="card">
    <h1 class="page-title">🧾 审计详情 #{{ row.id }}</h1>
    <p><b>时间：</b>{{ row.created_at }}</p><p><b>模块：</b>{{ row.module }}</p>
    <p><b>动作：</b>{{ row.action }}</p><p><b>记录：</b>{{ row.record_id or '-' }}</p>
    <p><b>操作人：</b>{{ row.actor }}</p><p><b>IP：</b>{{ row.ip_address or '-' }}</p>
    <p><b>原因：</b>{{ row.reason or '-' }}</p>
    <h3>旧值</h3><pre style="white-space:pre-wrap">{{ row.old_value | tojson(indent=2) }}</pre>
    <h3>新值</h3><pre style="white-space:pre-wrap">{{ row.new_value | tojson(indent=2) }}</pre>
    </div><a class="btn-tool btn-secondary" href="{{ url_for('finance_audit.audit_list') }}">← 返回</a>
    </div></body></html>
    """, row=row)
