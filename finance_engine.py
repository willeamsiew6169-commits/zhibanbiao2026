"""
Finance Web V7 shared business engine (V2).

This module contains reusable finance business logic for CHE, HQ and STW.
It imports the project's shared database helpers directly from db.py, so no
runtime configuration step is required.

App setup::

    from finance_engine import finance_engine_api_bp
    app.register_blueprint(finance_engine_api_bp)

Other modules may import the public helper functions directly, for example::

    from finance_engine import search_donors, get_member_snapshot
"""


from dataclasses import asdict, dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import re
from typing import Any, Iterable, Mapping, Optional

from flask import Blueprint, jsonify, request

from db import db_query


finance_engine_api_bp = Blueprint("finance_engine_api", __name__)


# =========================================================
# Database helpers
# =========================================================


def _fetchall(sql: str, params: tuple[Any, ...] = ()) -> list[dict[str, Any]]:
    rows = db_query(sql, params, fetchall=True) or []
    return [dict(row) for row in rows]


def _fetchone(sql: str, params: tuple[Any, ...] = ()) -> Optional[dict[str, Any]]:
    row = db_query(sql, params, fetchone=True)
    return dict(row) if row else None


def _execute(sql: str, params: tuple[Any, ...] = ()) -> Any:
    return db_query(sql, params)


# =========================================================
# Common normalization
# =========================================================


def normalize_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def normalize_phone(value: Any) -> str:
    """Keep digits only so 012-345 6789 and 0123456789 match."""
    return re.sub(r"\D", "", str(value or ""))


def normalize_branch(value: Any, default: str = "CHE") -> str:
    branch = normalize_text(value).upper()
    return branch if branch in {"CHE", "STW", "HQ"} else default


def normalize_member_id(value: Any, default_branch: str = "CHE") -> str:
    raw = normalize_text(value).upper().replace(" ", "")
    if not raw:
        return ""

    match = re.fullmatch(r"(CHE|STW)[-_]?(\d+)", raw)
    if match:
        return f"{match.group(1)}-{int(match.group(2))}"

    if raw.isdigit():
        return f"{normalize_branch(default_branch)}-{int(raw)}"

    return raw


def normalize_receipt_category(value: Any) -> str:
    """Map UI aliases to the five physical receipt books."""
    category = normalize_text(value)

    aliases = {
        "月费": "月费",
        "monthly fee": "月费",
        "财布施": "财布施",
        "特别布施": "财布施",
        "临时特别布施": "财布施",
        "纯檀香": "财布施",
        "纯檀香布施": "财布施",
        "观音堂纯檀香布施": "财布施",
        "观音村": "观音村",
        "膳食结缘": "膳食结缘",
        "初一十五": "膳食结缘",
        "初一十五膳食结缘": "膳食结缘",
    }

    return aliases.get(category.lower(), aliases.get(category, category))


def money(value: Any, default: Decimal = Decimal("0.00")) -> Decimal:
    try:
        result = Decimal(str(value or "0")).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError, TypeError):
        return default
    return result


def iso_date(value: Any) -> Optional[date]:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def _json_value(value: Any) -> Any:
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return value


def json_ready(row: Mapping[str, Any]) -> dict[str, Any]:
    return {key: _json_value(value) for key, value in row.items()}


# =========================================================
# Donor / person search
# =========================================================


@dataclass
class PersonSearchResult:
    source: str
    source_label: str
    person_key: str
    member_id: str = ""
    volunteer_id: str = ""
    name: str = ""
    english_name: str = ""
    phone: str = ""
    branch: str = ""
    status: str = ""
    last_category: str = ""
    last_amount: Decimal = Decimal("0.00")
    last_record_date: Optional[date] = None

    def to_dict(self) -> dict[str, Any]:
        return json_ready(asdict(self))


def _person_key(name: Any, phone: Any, member_id: Any = "") -> str:
    clean_phone = normalize_phone(phone)
    clean_member = normalize_text(member_id).upper()
    clean_name = normalize_text(name).casefold()
    return clean_phone or clean_member or clean_name


def search_donors(
    keyword: str,
    *,
    branch: Optional[str] = None,
    limit: int = 20,
) -> list[dict[str, Any]]:
    """Search members, volunteers and historical finance donors together."""
    keyword = normalize_text(keyword)
    if len(keyword) < 1:
        return []

    branch_filter = normalize_branch(branch) if branch else None
    member_id = normalize_member_id(keyword, branch_filter or "CHE")
    digits = normalize_phone(keyword)
    like = f"%{keyword}%"
    phone_like = f"%{digits}%" if digits else "%__NO_PHONE_MATCH__%"

    members = _fetchall(
        """
        select
            member_id,
            name,
            coalesce(english_name, '') as english_name,
            coalesce(phone, '') as phone,
            coalesce(branch, '') as branch,
            coalesce(status, '') as status
        from members
        where
            (%s is null or upper(coalesce(branch, '')) = %s)
            and (
                   upper(coalesce(member_id, '')) = upper(%s)
                or name ilike %s
                or coalesce(english_name, '') ilike %s
                or regexp_replace(coalesce(phone, ''), '[^0-9]', '', 'g') like %s
            )
        order by
            case when upper(coalesce(member_id, '')) = upper(%s) then 0 else 1 end,
            name
        limit %s
        """,
        (
            branch_filter,
            branch_filter,
            member_id,
            like,
            like,
            phone_like,
            member_id,
            limit,
        ),
    )

    volunteers = _fetchall(
        """
        select
            id as volunteer_id,
            name,
            coalesce(phone, '') as phone,
            coalesce(branch, '') as branch,
            coalesce(status, '') as status
        from volunteers
        where
            (%s is null or upper(coalesce(branch, '')) = %s)
            and (
                   upper(coalesce(id::text, '')) = upper(%s)
                or name ilike %s
                or regexp_replace(coalesce(phone, ''), '[^0-9]', '', 'g') like %s
            )
        order by
            case when upper(coalesce(id::text, '')) = upper(%s) then 0 else 1 end,
            name
        limit %s
        """,
        (
            branch_filter,
            branch_filter,
            member_id,
            like,
            phone_like,
            member_id,
            limit,
        ),
    )

    # finance_records 没有 branch 栏位，因此从 member_id / receipt_no 推断。
    # HQ 不是 finance_records 的实体分会代码，所以 HQ 搜索不限制 CHE/STW。
    history_branch_filter = branch_filter if branch_filter in {"CHE", "STW"} else None

    history = _fetchall(
        """
        select distinct on (
            coalesce(
                nullif(regexp_replace(coalesce(phone, ''), '[^0-9]', '', 'g'), ''),
                lower(name)
            )
        )
            coalesce(member_id, '') as member_id,
            name,
            coalesce(phone, '') as phone,
            case
                when upper(coalesce(member_id, '')) like 'STW%%'
                  or upper(coalesce(receipt_no, '')) like 'STW%%'
                    then 'STW'
                when upper(coalesce(member_id, '')) like 'CHE%%'
                  or upper(coalesce(receipt_no, '')) like 'CHE%%'
                    then 'CHE'
                else ''
            end as branch,
            category as last_category,
            amount as last_amount,
            record_date as last_record_date
        from finance_records
        where
            record_type = 'income'
            and coalesce(status, 'active') != 'cancelled'
            and category != '月费'
            and (
                %s is null
                or case
                    when upper(coalesce(member_id, '')) like 'STW%%'
                      or upper(coalesce(receipt_no, '')) like 'STW%%'
                        then 'STW'
                    when upper(coalesce(member_id, '')) like 'CHE%%'
                      or upper(coalesce(receipt_no, '')) like 'CHE%%'
                        then 'CHE'
                    else ''
                end = %s
            )
            and (
                   name ilike %s
                or upper(coalesce(member_id, '')) = upper(%s)
                or regexp_replace(coalesce(phone, ''), '[^0-9]', '', 'g') like %s
            )
        order by
            coalesce(
                nullif(regexp_replace(coalesce(phone, ''), '[^0-9]', '', 'g'), ''),
                lower(name)
            ),
            record_date desc,
            id desc
        limit %s
        """,
        (
            history_branch_filter,
            history_branch_filter,
            like,
            member_id,
            phone_like,
            limit,
        ),
    )

    merged: dict[str, PersonSearchResult] = {}

    for row in members:
        key = _person_key(row.get("name"), row.get("phone"), row.get("member_id"))
        merged[key] = PersonSearchResult(
            source="member",
            source_label="月费会员",
            person_key=key,
            member_id=normalize_text(row.get("member_id")),
            name=normalize_text(row.get("name")),
            english_name=normalize_text(row.get("english_name")),
            phone=normalize_text(row.get("phone")),
            branch=normalize_text(row.get("branch")),
            status=normalize_text(row.get("status")),
        )

    for row in volunteers:
        key = _person_key(row.get("name"), row.get("phone"), row.get("volunteer_id"))
        existing = merged.get(key)
        if existing:
            existing.volunteer_id = normalize_text(row.get("volunteer_id"))
            if not existing.phone:
                existing.phone = normalize_text(row.get("phone"))
            continue

        merged[key] = PersonSearchResult(
            source="volunteer",
            source_label="义工",
            person_key=key,
            volunteer_id=normalize_text(row.get("volunteer_id")),
            name=normalize_text(row.get("name")),
            phone=normalize_text(row.get("phone")),
            branch=normalize_text(row.get("branch")),
            status=normalize_text(row.get("status")),
        )

    for row in history:
        key = _person_key(row.get("name"), row.get("phone"), row.get("member_id"))
        existing = merged.get(key)
        if existing:
            existing.last_category = normalize_text(row.get("last_category"))
            existing.last_amount = money(row.get("last_amount"))
            existing.last_record_date = iso_date(row.get("last_record_date"))
            continue

        merged[key] = PersonSearchResult(
            source="history",
            source_label="历史布施人",
            person_key=key,
            member_id=normalize_text(row.get("member_id")),
            name=normalize_text(row.get("name")),
            phone=normalize_text(row.get("phone")),
            branch=normalize_text(row.get("branch")),
            last_category=normalize_text(row.get("last_category")),
            last_amount=money(row.get("last_amount")),
            last_record_date=iso_date(row.get("last_record_date")),
        )

    source_rank = {"member": 0, "volunteer": 1, "history": 2}
    results = sorted(
        merged.values(),
        key=lambda item: (
            source_rank.get(item.source, 9),
            item.name.casefold(),
        ),
    )
    return [item.to_dict() for item in results[:limit]]


# =========================================================
# Member fast-entry snapshot
# =========================================================


def _month_start(value: date) -> date:
    return value.replace(day=1)


def _month_diff(start: date, end: date) -> int:
    return (end.year - start.year) * 12 + end.month - start.month


def get_member_snapshot(
    member_keyword: str,
    *,
    branch: str = "CHE",
    as_of: Optional[date] = None,
) -> Optional[dict[str, Any]]:
    normalized_id = normalize_member_id(member_keyword, branch)
    keyword = normalize_text(member_keyword)
    like = f"%{keyword}%"

    member = _fetchone(
        """
        select
            member_id,
            name,
            coalesce(english_name, '') as english_name,
            coalesce(phone, '') as phone,
            coalesce(branch, '') as branch,
            coalesce(status, '') as status
        from members
        where
               upper(member_id) = upper(%s)
            or name ilike %s
            or coalesce(english_name, '') ilike %s
        order by
            case when upper(member_id) = upper(%s) then 0 else 1 end,
            member_id
        limit 1
        """,
        (normalized_id, like, like, normalized_id),
    )
    if not member:
        return None

    payment = _fetchone(
        """
        select
            max(end_month) as paid_until,
            max(payment_date) as last_payment_date,
            coalesce(sum(amount), 0) as total_paid,
            coalesce(sum(month_count), 0) as total_months
        from member_payments
        where member_id = %s
        """,
        (member["member_id"],),
    ) or {}

    today = as_of or date.today()
    current_month = _month_start(today)
    paid_until = iso_date(payment.get("paid_until"))

    if paid_until:
        next_due_month = _month_start(paid_until)
        if next_due_month < current_month:
            months_due = _month_diff(next_due_month, current_month)
        else:
            months_due = 0
    else:
        months_due = 1

    recommended_months = max(1, months_due)

    result = {
        **member,
        "paid_until": paid_until,
        "last_payment_date": iso_date(payment.get("last_payment_date")),
        "total_paid": money(payment.get("total_paid")),
        "total_months": int(payment.get("total_months") or 0),
        "months_due": months_due,
        "recommended_months": recommended_months,
        "recommended_amount": Decimal(recommended_months * 50).quantize(Decimal("0.00")),
    }
    return json_ready(result)


# =========================================================
# Donor history / amount suggestion
# =========================================================


def get_donor_history(
    *,
    phone: str = "",
    name: str = "",
    branch: Optional[str] = None,
    limit: int = 8,
) -> dict[str, Any]:
    clean_phone = normalize_phone(phone)
    clean_name = normalize_text(name)
    if not clean_phone and not clean_name:
        return {"records": [], "suggested_amount": 0.0, "last_category": ""}

    branch_filter = normalize_branch(branch) if branch else None
    history_branch_filter = branch_filter if branch_filter in {"CHE", "STW"} else None

    rows = _fetchall(
        """
        select
            id,
            record_date,
            receipt_no,
            member_id,
            name,
            coalesce(phone, '') as phone,
            category,
            amount,
            payment_method,
            case
                when upper(coalesce(member_id, '')) like 'STW%%'
                  or upper(coalesce(receipt_no, '')) like 'STW%%'
                    then 'STW'
                when upper(coalesce(member_id, '')) like 'CHE%%'
                  or upper(coalesce(receipt_no, '')) like 'CHE%%'
                    then 'CHE'
                else ''
            end as branch
        from finance_records
        where
            record_type = 'income'
            and category != '月费'
            and coalesce(status, 'active') != 'cancelled'
            and (
                %s is null
                or case
                    when upper(coalesce(member_id, '')) like 'STW%'
                      or upper(coalesce(receipt_no, '')) like 'STW%'
                        then 'STW'
                    when upper(coalesce(member_id, '')) like 'CHE%'
                      or upper(coalesce(receipt_no, '')) like 'CHE%'
                        then 'CHE'
                    else ''
                end = %s
            )
            and (
                (%s <> '' and regexp_replace(coalesce(phone, ''), '[^0-9]', '', 'g') = %s)
                or
                (%s <> '' and lower(name) = lower(%s))
            )
        order by record_date desc, id desc
        limit %s
        """,
        (
            history_branch_filter,
            history_branch_filter,
            clean_phone,
            clean_phone,
            clean_name,
            clean_name,
            limit,
        ),
    )

    amounts = [money(row.get("amount")) for row in rows]
    suggested_amount = amounts[0] if amounts else Decimal("0.00")
    last_category = normalize_text(rows[0].get("category")) if rows else ""

    return {
        "records": [json_ready(row) for row in rows],
        "suggested_amount": float(suggested_amount),
        "last_category": last_category,
    }


# =========================================================
# Receipt books
# =========================================================


def get_receipt_book(branch: str, category: str) -> Optional[dict[str, Any]]:
    branch = normalize_branch(branch)
    category = normalize_receipt_category(category)
    return _fetchone(
        """
        select *
        from finance_receipt_books
        where upper(branch) = %s
          and category = %s
          and coalesce(is_active, true) = true
        limit 1
        """,
        (branch, category),
    )


def get_next_receipt_no(branch: str, category: str) -> str:
    """Read the next number without consuming it."""
    book = get_receipt_book(branch, category)
    if not book:
        raise ValueError(f"找不到收条簿：{branch} / {normalize_receipt_category(category)}")

    for field in ("next_receipt_no", "current_receipt_no", "last_receipt_no"):
        value = normalize_text(book.get(field))
        if value:
            if field == "last_receipt_no":
                return increment_document_no(value)
            return value

    raise ValueError("收条簿没有可用的号码字段")


def increment_document_no(value: str) -> str:
    raw = normalize_text(value).upper()
    match = re.fullmatch(r"(.*?)(\d+)$", raw)
    if not match:
        raise ValueError(f"无法识别号码格式：{value}")

    prefix, number = match.groups()
    return f"{prefix}{int(number) + 1:0{len(number)}d}"


def update_receipt_book_number(branch: str, category: str, receipt_no: str) -> None:
    branch = normalize_branch(branch)
    category = normalize_receipt_category(category)
    next_no = increment_document_no(receipt_no)

    _execute(
        """
        update finance_receipt_books
        set
            last_receipt_no = %s,
            next_receipt_no = %s,
            updated_at = now()
        where upper(branch) = %s
          and category = %s
          and coalesce(is_active, true) = true
        """,
        (receipt_no, next_no, branch, category),
    )


# =========================================================
# Unified record validation
# =========================================================


@dataclass
class ValidationMessage:
    code: str
    level: str
    message: str

    def to_dict(self) -> dict[str, str]:
        return asdict(self)


def validate_finance_record(payload: Mapping[str, Any]) -> dict[str, Any]:
    """Run non-destructive checks before an income or expense is saved."""
    messages: list[ValidationMessage] = []

    record_type = normalize_text(payload.get("record_type") or "income").lower()
    branch = normalize_branch(payload.get("branch") or "CHE")
    category = normalize_text(payload.get("category"))
    receipt_no = normalize_text(payload.get("receipt_no")).upper()
    pv_no = normalize_text(
        payload.get("payment_voucher_no") or payload.get("pv_no")
    ).upper()
    bank_ref = normalize_text(payload.get("bank_ref"))
    member_id = normalize_text(payload.get("member_id")).upper()
    record_date = iso_date(payload.get("record_date"))
    amount = money(payload.get("amount"), Decimal("-1"))

    if record_type not in {"income", "expense"}:
        messages.append(ValidationMessage("record_type", "error", "记录类型不正确。"))

    if not category:
        messages.append(ValidationMessage("category", "error", "请选择类别。"))

    if amount <= 0:
        messages.append(ValidationMessage("amount", "error", "金额必须大于 RM0。"))

    if record_date is None:
        messages.append(ValidationMessage("record_date", "error", "日期格式不正确。"))
    elif record_date > date.today():
        messages.append(ValidationMessage("future_date", "warning", "日期晚于今天，请确认。"))

    if receipt_no:
        duplicate = _fetchone(
            """
            select id, record_date, name, amount
            from finance_records
            where upper(receipt_no) = %s
              and coalesce(status, 'active') != 'cancelled'
            limit 1
            """,
            (receipt_no,),
        )
        if duplicate:
            messages.append(
                ValidationMessage(
                    "duplicate_receipt",
                    "error",
                    f"收条 {receipt_no} 已经存在。",
                )
            )

    if pv_no:
        duplicate = _fetchone(
            """
            select id, record_date, name, amount
            from finance_records
            where upper(coalesce(payment_voucher_no, '')) = %s
              and coalesce(status, 'active') != 'cancelled'
            limit 1
            """,
            (pv_no,),
        )
        if duplicate:
            messages.append(
                ValidationMessage("duplicate_pv", "error", f"PV {pv_no} 已经存在。")
            )

    if bank_ref:
        duplicate = _fetchone(
            """
            select id, record_date, name, amount
            from finance_records
            where lower(coalesce(bank_ref, '')) = lower(%s)
              and coalesce(status, 'active') != 'cancelled'
            limit 1
            """,
            (bank_ref,),
        )
        if duplicate:
            messages.append(
                ValidationMessage(
                    "duplicate_bank_ref",
                    "warning",
                    f"银行参考编号 {bank_ref} 曾经使用，请确认。",
                )
            )

    if member_id and record_date and category == "月费":
        same_day = _fetchone(
            """
            select id, receipt_no, amount
            from finance_records
            where upper(coalesce(member_id, '')) = %s
              and category = '月费'
              and record_date = %s
              and coalesce(status, 'active') != 'cancelled'
            limit 1
            """,
            (member_id, record_date),
        )
        if same_day:
            messages.append(
                ValidationMessage(
                    "same_member_same_day",
                    "warning",
                    f"{member_id} 今天已经有月费记录，请确认是否重复。",
                )
            )

    if record_date and amount > 0:
        similar = _fetchone(
            """
            select id, receipt_no, name, category
            from finance_records
            where record_date = %s
              and amount = %s
              and record_type = %s
              and coalesce(status, 'active') != 'cancelled'
            limit 1
            """,
            (record_date, amount, record_type),
        )
        if similar:
            messages.append(
                ValidationMessage(
                    "same_date_amount",
                    "info",
                    f"同一天已有一笔 RM{amount:.2f} 的{('收入' if record_type == 'income' else '支出')}，请留意。",
                )
            )

    if receipt_no and category:
        latest = _fetchone(
            """
            select receipt_no, record_date
            from finance_records
            where upper(coalesce(branch, '')) = %s
              and category = %s
              and receipt_no is not null
              and receipt_no <> ''
              and coalesce(status, 'active') != 'cancelled'
            order by record_date desc, id desc
            limit 1
            """,
            (branch, category),
        )
        if latest and record_date:
            latest_date = iso_date(latest.get("record_date"))
            if latest_date and record_date < latest_date:
                messages.append(
                    ValidationMessage(
                        "older_than_latest_receipt",
                        "warning",
                        "这张收条日期早于该类别上一张已录入收条，请确认日期。",
                    )
                )

    errors = [message.to_dict() for message in messages if message.level == "error"]
    warnings = [message.to_dict() for message in messages if message.level == "warning"]
    info = [message.to_dict() for message in messages if message.level == "info"]

    return {
        "ok": not errors,
        "errors": errors,
        "warnings": warnings,
        "info": info,
        "all": [message.to_dict() for message in messages],
    }


# =========================================================
# Shared statistics
# =========================================================


def get_workspace_statistics(branch: str, ym: str) -> dict[str, Any]:
    """One SQL result for V7 dashboard cards and month-close summaries."""
    branch = normalize_branch(branch)
    if not re.fullmatch(r"\d{4}-\d{2}", normalize_text(ym)):
        raise ValueError("ym 必须是 YYYY-MM")

    row = _fetchone(
        """
        select
            coalesce(sum(amount) filter (
                where record_type = 'income' and category = '月费'
            ), 0) as monthly_fee_total,

            count(*) filter (
                where record_type = 'income' and category = '月费'
            ) as monthly_fee_count,

            coalesce(sum(amount) filter (
                where record_type = 'income' and category != '月费'
            ), 0) as donation_total,

            count(*) filter (
                where record_type = 'income' and category != '月费'
            ) as donation_count,

            coalesce(sum(amount) filter (
                where record_type = 'expense'
            ), 0) as expense_total,

            count(*) filter (
                where record_type = 'expense'
            ) as expense_count
        from finance_records
        where upper(coalesce(branch, '')) = %s
          and to_char(record_date, 'YYYY-MM') = %s
          and coalesce(status, 'active') != 'cancelled'
        """,
        (branch, ym),
    ) or {}

    pending = _fetchone(
        """
        select
            count(*) as pending_count,
            coalesce(sum(amount), 0) as pending_amount
        from bank_pending_records
        where upper(coalesce(branch, '')) = %s
          and coalesce(status, 'pending') = 'pending'
        """,
        (branch,),
    ) or {}

    return json_ready({**row, **pending, "branch": branch, "ym": ym})


# =========================================================
# API routes
# =========================================================


@finance_engine_api_bp.get("/finance/api/donor-search")
def finance_api_donor_search():
    keyword = normalize_text(request.args.get("q"))
    branch = normalize_text(request.args.get("branch")) or None

    if not keyword:
        return jsonify({"ok": True, "results": []})

    try:
        results = search_donors(keyword, branch=branch, limit=20)
        return jsonify({"ok": True, "results": results})
    except Exception as exc:  # route boundary: return a useful API error
        return jsonify({"ok": False, "message": str(exc), "results": []}), 500


@finance_engine_api_bp.get("/finance/api/member/<member_keyword>")
def finance_api_member(member_keyword: str):
    branch = normalize_text(request.args.get("branch")) or "CHE"

    try:
        member = get_member_snapshot(member_keyword, branch=branch)
        if not member:
            return jsonify({"ok": False, "message": "找不到会员。"}), 404
        return jsonify({"ok": True, "member": member})
    except Exception as exc:
        return jsonify({"ok": False, "message": str(exc)}), 500


@finance_engine_api_bp.get("/finance/api/donor-history")
def finance_api_donor_history():
    phone = normalize_text(request.args.get("phone"))
    name = normalize_text(request.args.get("name"))
    branch = normalize_text(request.args.get("branch")) or None

    if not phone and not name:
        return jsonify({"ok": False, "message": "请提供电话或姓名。"}), 400

    try:
        result = get_donor_history(phone=phone, name=name, branch=branch)
        return jsonify({"ok": True, **result})
    except Exception as exc:
        return jsonify({"ok": False, "message": str(exc)}), 500


@finance_engine_api_bp.post("/finance/api/validate-record")
def finance_api_validate_record():
    payload = request.get_json(silent=True) or request.form.to_dict()

    try:
        result = validate_finance_record(payload)
        return jsonify(result), (200 if result["ok"] else 422)
    except Exception as exc:
        return jsonify({"ok": False, "message": str(exc)}), 500


__all__ = [
    "configure_finance_engine",
    "finance_engine_api_bp",
    "search_donors",
    "get_member_snapshot",
    "get_donor_history",
    "get_receipt_book",
    "get_next_receipt_no",
    "update_receipt_book_number",
    "validate_finance_record",
    "get_workspace_statistics",
    "normalize_member_id",
    "normalize_phone",
    "normalize_receipt_category",
]
