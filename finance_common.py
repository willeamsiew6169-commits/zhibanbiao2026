"""Finance V6 shared helpers.

This module must stay independent from finance blueprints to avoid circular imports.
"""
from __future__ import annotations

import json
import os
from datetime import date, datetime
from decimal import Decimal
from typing import Any
from zoneinfo import ZoneInfo

from flask import has_request_context, request, session

from db import db_query

MY_TZ = ZoneInfo("Asia/Kuala_Lumpur")
DEFAULT_FINANCE_BRANCH = os.getenv("DEFAULT_FINANCE_BRANCH", "CHE").strip().upper() or "CHE"

FINANCE_STATUS_CONFIRMED = "confirmed"
FINANCE_STATUS_CANCELLED = "cancelled"
MEMBER_PAYMENT_ACTIVE = "active"
MEMBER_PAYMENT_CANCELLED = "cancelled"


def malaysia_now() -> datetime:
    return datetime.now(MY_TZ)


def malaysia_today() -> date:
    return malaysia_now().date()


def money(value: Any) -> float:
    try:
        if value in (None, ""):
            return 0.0
        return float(Decimal(str(value).replace(",", "").strip()))
    except (ValueError, TypeError, ArithmeticError):
        return 0.0


def normalize_finance_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value or "").strip()
    if not text:
        return malaysia_today()

    try:
        if len(text) == 7:
            return datetime.strptime(text + "-01", "%Y-%m-%d").date()
        return datetime.strptime(text[:10], "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"无法识别财政日期：{text}") from exc


def get_finance_ym(value: Any) -> str:
    return normalize_finance_date(value).strftime("%Y-%m")


def get_current_finance_branch() -> str:
    if has_request_context():
        return str(session.get("finance_branch") or DEFAULT_FINANCE_BRANCH).strip().upper()
    return DEFAULT_FINANCE_BRANCH


def get_current_finance_user() -> str:
    if has_request_context():
        return str(
            session.get("finance_user")
            or session.get("admin_name")
            or session.get("member_admin_name")
            or "Finance User"
        ).strip()
    return "System"


def get_request_ip() -> str:
    if not has_request_context():
        return ""
    forwarded = request.headers.get("X-Forwarded-For", "").split(",")[0].strip()
    return forwarded or request.remote_addr or ""


def json_safe_value(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(k): json_safe_value(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [json_safe_value(v) for v in value]
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return value


def json_dumps_safe(value: Any) -> str:
    return json.dumps(json_safe_value(value), ensure_ascii=False, sort_keys=True)


def get_month_close_record(
    record_date: Any,
    fund_account: str = "观音堂日常户口",
):
    """
    读取指定月份、指定户口已经完成的月结记录。

    record_date 支持：
    - "2026-07"
    - "2026-07-15"
    - date / datetime

    默认读取：
    - 观音堂日常户口
    """

    ym = get_finance_ym(record_date)

    return db_query(
        """
        select *
        from finance_month_close
        where ym = %s
          and fund_account = %s
          and status = 'closed'
        order by id desc
        limit 1
        """,
        (
            ym,
            fund_account,
        ),
        fetchone=True,
    )


def is_finance_month_closed(record_date: Any, fund_account: str | None = None) -> bool:
    return bool(get_month_close_record(record_date, fund_account))


def get_month_lock_message(record_date: Any, fund_account: str | None = None) -> str:
    ym = get_finance_ym(record_date)
    account_text = f"（{fund_account}）" if fund_account else ""
    return (
        f"{ym} {account_text}已经完成月结，"
        "财政记录已锁定，不能新增、修改或作废。"
        "如确实需要调整，请先由管理员解除月结。"
    )


def require_finance_month_open(record_date: Any, fund_account: str | None = None) -> str | None:
    if is_finance_month_closed(record_date, fund_account):
        return get_month_lock_message(record_date, fund_account)
    return None
