Food Offering V2（观音堂系统专用版）
=====================================

文件位置
--------

把 3 个 Python 文件放到：

schedule/services/

即：

schedule/services/food_offering_service.py
schedule/services/food_offering_routes.py
schedule/services/food_offering_templates.py


一、先执行 SQL
-------------

在 Supabase SQL Editor 执行：

food_offering_supabase.sql


二、schedule_web.py 修改
-----------------------

1. settings_service import 必须包含：

from schedule.services.settings_service import (
    get_schedule_setting,
    get_schedule_settings,
    save_schedule_setting,
    set_schedule_setting,
    is_schedule_setting_on,
)

2. import 注册函数：

from schedule.services.food_offering_routes import (
    register_food_offering_routes
)

3. 在 schedule_records = [] 下面注册：

register_food_offering_routes(
    schedule_bp,
    is_schedule_setting_on=is_schedule_setting_on,
)

注意：
必须在 Flask app.register_blueprint(schedule_bp) 之前执行。


三、义工首页按钮
---------------

顶部 volunteer-actions 内使用：

{% if food_offering_open %}

<a
    class="btn-tool btn-green"
    href="{{ url_for(
        'schedule.food_offering_page'
    ) }}">

    <span class="action-main-text">
        🥗 素食结缘报名
    </span>

    <span class="action-sub-text">
        填写食物名称及份量
    </span>

</a>

{% endif %}


四、删除旧版
-----------

从旧的每日报名 choice-grid 删除：

data-role="素食结缘"

删除：

food_offering_task_section

删除旧的：

selectFoodOfferingTask()

删除旧的：

/volunteer/food_offering_status

新版名单 URL 为：

/volunteer/food-offering/status


五、正式 URL
-----------

GET
/volunteer/food-offering

POST
/volunteer/food-offering/signup

GET
/volunteer/food-offering/status

GET
/volunteer/food-offering/whatsapp


六、若启动时出现 endpoint 已存在
-------------------------------

全项目搜索：

register_food_offering_routes(

实际调用只能有一次。
