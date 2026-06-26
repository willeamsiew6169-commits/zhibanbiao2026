# flatten_builder.py

def flatten_arranged_for_db(arranged):
    rows = []

    for group_name, people in arranged.items():

        for p in people:

            signup_id = None
            volunteer_id = None

            if isinstance(p, dict):

                row = dict(p)

                name = row.get("name") or row.get("姓名")
                start_time = row.get("start_time") or row.get("开始时间")
                end_time = row.get("end_time") or row.get("结束时间")

                signup_id = row.get("signup_id")
                volunteer_id = row.get("volunteer_id")

            elif isinstance(p, (list, tuple)):

                name = p[0] if len(p) > 0 else None
                start_time = p[1] if len(p) > 1 else None
                end_time = p[2] if len(p) > 2 else None

                row = {
                    "name": name,
                    "start_time": start_time,
                    "end_time": end_time,
                }

            else:

                name = str(p).strip() if p else None
                start_time = None
                end_time = None

                row = {
                    "name": name
                }

            if not name:
                continue

            shift_label = None
            assigned_place = group_name
            role = row.get("role") or row.get("岗位")

            # ------------------------
            # 整理佛台
            # ------------------------
            if group_name == "整理佛台":

                role = "整理佛台"
                assigned_place = "整理佛台"

                start_time = None
                end_time = None

            # ------------------------
            # 卫生
            # ------------------------
            elif group_name in [
                "佛堂卫生",
                "二楼卫生",
                "楼梯卫生"
            ]:

                role = "卫生"
                assigned_place = group_name

                start_time = None
                end_time = None

            # ------------------------
            # 供台
            # ------------------------
            elif group_name == "设师父供台":

                role = "供台"
                assigned_place = "设师父供台"

                start_time = None
                end_time = None

            # ------------------------
            # 值班
            # ------------------------
            elif group_name in [

                "绿观音堂",
                "绿活动中心",

                "橙观音堂",
                "橙活动中心",

                "黄观音堂",
                "黄活动中心",

            ]:

                role = "值班"

                shift_label = group_name[0] + "班"

                if "观音堂" in group_name:
                    assigned_place = "观音堂"

                elif "活动中心" in group_name:
                    assigned_place = "活动中心"

            else:

                assigned_place = group_name

            rows.append({

                "signup_id": signup_id,
                "volunteer_id": volunteer_id,

                "name": name,
                "role": role,

                "shift_label": shift_label,
                "assigned_place": assigned_place,

                "start_time": start_time,
                "end_time": end_time,

            })

    return rows