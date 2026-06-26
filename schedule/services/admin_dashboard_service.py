# admin_dashboard_service.py

from datetime import datetime, date

from lunar_rules import (
    get_special_day_info,
    get_next_day_remove_info,
)

from schedule.services.assignment_service import (
    load_display_records,
    load_schedule_admin_dashboard_data,
)

from schedule.services.shortage_service import build_shortage_summary_for_admin
from schedule.services.supply_service import (
    load_supply_signups_for_date,
    load_upcoming_supply_signup_alerts,
)
from schedule.services.publish_service import is_schedule_published
from schedule.services.settings_service import get_schedule_settings

def load_admin_dashboard_data(mode, override_date):
    selected_date_obj = datetime.strptime(
        override_date,
        "%Y-%m-%d"
    ).date()

    special_day_info = get_special_day_info(selected_date_obj)

    template_text = {
        "normal": "平时值班模板",
        "lunar_1_15": "初一十五值班模板",
        "buddhist_festival": "佛诞大日子模板",
    }

    special_day_info["template_text"] = template_text.get(
        special_day_info["template_type"],
        special_day_info["template_type"]
    )

    remove_info = get_next_day_remove_info(selected_date_obj)
    shortage_summary = build_shortage_summary_for_admin(override_date)
    supply_signups = load_supply_signups_for_date(override_date)

    settings = get_schedule_settings()

    try:
        supply_alert_days = int(settings.get("supply_alert_days", 7))
    except:
        supply_alert_days = 7

    supply_alerts = load_upcoming_supply_signup_alerts(
        days_ahead=supply_alert_days
    )

    published = is_schedule_published(override_date)

    pending_counts, day_summary, day_flags = load_schedule_admin_dashboard_data(
        override_date
    )

    is_today_or_past = selected_date_obj < date.today()

    today = date.today()

    if today.month == 12:
        next_year = today.year + 1
        next_month = 1
    else:
        next_year = today.year
        next_month = today.month + 1

    display_records = load_display_records(
        mode=mode,
        target_date=override_date if mode == "day" else None,
        year=next_year if mode == "prebook" else None,
        month=next_month if mode == "prebook" else None,
    )

    return {
        "special_day_info": special_day_info,
        "remove_info": remove_info,
        "shortage_summary": shortage_summary,
        "supply_signups": supply_signups,
        "supply_alerts": supply_alerts,
        "is_published": published,
        "pending_counts": pending_counts,
        "day_summary": day_summary,
        "day_flags": day_flags,
        "is_today_or_past": is_today_or_past,
        "default_year": next_year,
        "default_month": next_month,
        "records": display_records,
    }