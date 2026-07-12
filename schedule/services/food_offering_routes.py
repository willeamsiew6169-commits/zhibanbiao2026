# schedule/services/food_offering_routes.py
# 观音堂系统专用版 Food Offering V2

from datetime import datetime

from flask import (
    request,
    redirect,
    url_for,
    render_template_string,
    session,
)

from lunar_rules import get_special_day_info

from schedule.builders.time_utils import (
    malaysia_today,
)

from schedule.services.food_offering_service import (
    FOOD_OFFERING_CATEGORIES,
    WEEKDAY_TEXT,
    find_next_food_offering_date,
    get_food_offering_special_text,
    format_food_offering_deadline,
    is_food_offering_deadline_open,
    load_food_offering_records,
    resolve_food_offering_volunteer,
    save_food_offering_signup,
    cancel_food_offering_signup,
    build_food_offering_whatsapp_text,
)

from schedule.services.food_offering_templates import (
    FOOD_OFFERING_SIGNUP_HTML,
    FOOD_OFFERING_STATUS_HTML,
    FOOD_OFFERING_WHATSAPP_HTML,
)


def register_food_offering_routes(
    schedule_bp,
    *,
    is_schedule_setting_on,
):

    @schedule_bp.route(
        "/volunteer/food-offering",
        methods=["GET"],
        endpoint="food_offering_page",
    )
    def food_offering_page():

        if not is_schedule_setting_on(
            "food_offering_open"
        ):

            return """
            <h1>🥗 素食结缘报名尚未开放</h1>
            <p>
                请等待负责人开放下一场
                素食结缘报名。
            </p>
            <a href="/volunteer">
                返回义工首页
            </a>
            """

        offering_date, special_info = (
            find_next_food_offering_date(
                start_date=malaysia_today(),
                days_ahead=40,
            )
        )

        if not offering_date:

            return """
            <h1>❌ 找不到下一场大日子</h1>
            <p>
                系统在未来40天内找不到
                初一、十五或佛诞日。
            </p>
            <a href="/volunteer">
                返回义工首页
            </a>
            """
        
        lunar_text, festival_name = (
            get_food_offering_special_text(
                special_info
            )
        )

        _, category_summary = (
            load_food_offering_records(
                offering_date
            )
        )

        return render_template_string(
            FOOD_OFFERING_SIGNUP_HTML,

            offering_date=(
                offering_date.strftime(
                    "%Y-%m-%d"
                )
            ),

            display_date=(
                offering_date.strftime(
                    "%d/%m/%Y"
                )
            ),

            weekday_text=(
                WEEKDAY_TEXT[
                    offering_date.weekday()
                ]
            ),

            lunar_text=lunar_text,
            festival_name=festival_name,

            deadline_text=(
                format_food_offering_deadline(
                    offering_date
                )
            ),

            categories=(
                FOOD_OFFERING_CATEGORIES
            ),

            category_summary=(
                category_summary
            ),
        )


    @schedule_bp.route(
        "/volunteer/food-offering/signup",
        methods=["POST"],
        endpoint="food_offering_signup",
    )
    def food_offering_signup():

        if not is_schedule_setting_on(
            "food_offering_open"
        ):

            return """
            <h1>❌ 素食结缘报名尚未开放</h1>
            <a href="/volunteer">
                返回义工首页
            </a>
            """

        keyword = request.form.get(
            "keyword",
            ""
        ).strip()

        branch = request.form.get(
            "branch",
            "CHE"
        ).strip().upper()

        offering_date_text = (
            request.form.get(
                "offering_date",
                ""
            ).strip()
        )

        category = request.form.get(
            "category",
            ""
        ).strip().lower()

        food_name = request.form.get(
            "food_name",
            ""
        ).strip()

        quantity_text = request.form.get(
            "quantity_text",
            ""
        ).strip()

        try:

            offering_date = (
                datetime.strptime(
                    offering_date_text,
                    "%Y-%m-%d",
                ).date()
            )

        except ValueError:

            return """
            <h1>❌ 日期格式错误</h1>
            <a href="/volunteer/food-offering">
                返回重新报名
            </a>
            """

        expected_date, _ = (
            find_next_food_offering_date(
                start_date=malaysia_today(),
                days_ahead=40,
            )
        )

        if (
            not expected_date
            or offering_date != expected_date
        ):

            return """
            <h1>❌ 素食结缘日期已经改变</h1>
            <p>
                请重新进入报名页面确认日期。
            </p>
            <a href="/volunteer/food-offering">
                返回重新报名
            </a>
            """
        
        volunteer, volunteer_error = (
            resolve_food_offering_volunteer(
                keyword,
                branch,
            )
        )

        if volunteer_error:

            return f"""
            <h1>❌ 报名失败</h1>
            <p>{volunteer_error}</p>
            <a href="/volunteer/food-offering">
                返回重新报名
            </a>
            """

        result = save_food_offering_signup(
            volunteer_id=volunteer["id"],
            volunteer_name=volunteer["name"],
            offering_date=offering_date,
            category=category,
            food_name=food_name,
            quantity_text=quantity_text,
        )

        if not result["ok"]:

            return f"""
            <h1>❌ 报名失败</h1>
            <p>{result["error"]}</p>
            <a href="/volunteer/food-offering">
                返回重新报名
            </a>
            """

        duplicate_food = ""

        if result.get("duplicate"):

            duplicate_food = (
                result["duplicate"]["food_name"]
            )

        return redirect(
            url_for(
                "schedule.food_offering_status",

                date=offering_date.strftime(
                    "%Y-%m-%d"
                ),

                success="1",

                duplicate_food=(
                    duplicate_food
                ),
            )
        )


    @schedule_bp.route(
        "/volunteer/food-offering/status",
        methods=["GET"],
        endpoint="food_offering_status",
    )
    def food_offering_status():

        date_text = request.args.get(
            "date",
            ""
        ).strip()

        offering_date = None

        if date_text:

            try:

                offering_date = (
                    datetime.strptime(
                        date_text,
                        "%Y-%m-%d",
                    ).date()
                )

            except ValueError:

                offering_date = None

        if not offering_date:

            offering_date, special_info = (
                find_next_food_offering_date(
                    start_date=malaysia_today(),
                    days_ahead=40,
                )
            )

        else:

            special_info = (
                get_special_day_info(
                    offering_date
                )
            )

        if not offering_date:

            return """
            <h1>❌ 找不到素食结缘日期</h1>
            <a href="/volunteer">
                返回义工首页
            </a>
            """

        lunar_text, festival_name = (
            get_food_offering_special_text(
                special_info
            )
        )

        (
            grouped_records,
            category_summary,
        ) = load_food_offering_records(
            offering_date
        )

        signup_open = is_schedule_setting_on(
            "food_offering_open"
        )

        return render_template_string(
            FOOD_OFFERING_STATUS_HTML,

            offering_date=(
                offering_date.strftime(
                    "%Y-%m-%d"
                )
            ),

            display_date=(
                offering_date.strftime(
                    "%d/%m/%Y"
                )
            ),

            weekday_text=(
                WEEKDAY_TEXT[
                    offering_date.weekday()
                ]
            ),

            lunar_text=lunar_text,
            festival_name=festival_name,

            categories=(
                FOOD_OFFERING_CATEGORIES
            ),

            grouped_records=(
                grouped_records
            ),

            category_summary=(
                category_summary
            ),

            success=(
                request.args.get("success")
                == "1"
            ),

            duplicate_food=(
                request.args.get(
                    "duplicate_food",
                    ""
                ).strip()
            ),

            signup_open=signup_open,

            is_admin=bool(
                session.get(
                    "schedule_login"
                )
            ),
        )


    @schedule_bp.route(
        "/volunteer/food-offering/cancel/"
        "<int:signup_id>",
        methods=["POST"],
        endpoint="food_offering_cancel",
    )
    def food_offering_cancel(
        signup_id,
    ):

        if not session.get(
            "schedule_login"
        ):

            return redirect(
                url_for(
                    "schedule.schedule_admin"
                )
            )

        date_text = request.form.get(
            "date",
            ""
        ).strip()

        cancel_food_offering_signup(
            signup_id
        )

        return redirect(
            url_for(
                "schedule.food_offering_status",
                date=date_text,
            )
        )


    @schedule_bp.route(
        "/volunteer/food-offering/whatsapp",
        methods=["GET"],
        endpoint="food_offering_whatsapp",
    )
    def food_offering_whatsapp():

        if not session.get(
            "schedule_login"
        ):

            return redirect(
                url_for(
                    "schedule.schedule_admin"
                )
            )

        date_text = request.args.get(
            "date",
            ""
        ).strip()

        try:

            offering_date = (
                datetime.strptime(
                    date_text,
                    "%Y-%m-%d",
                ).date()
            )

        except ValueError:

            return """
            <h1>❌ 日期错误</h1>
            <a href="/schedule/admin">
                返回负责人页面
            </a>
            """

        special_info = (
            get_special_day_info(
                offering_date
            )
        )

        (
            grouped_records,
            category_summary,
        ) = load_food_offering_records(
            offering_date
        )

        whatsapp_text = (
            build_food_offering_whatsapp_text(
                offering_date,
                special_info,
                grouped_records,
                category_summary,
            )
        )

        return render_template_string(
            FOOD_OFFERING_WHATSAPP_HTML,

            whatsapp_text=(
                whatsapp_text
            ),

            offering_date=(
                offering_date.strftime(
                    "%Y-%m-%d"
                )
            ),
        )
