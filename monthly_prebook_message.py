import pandas as pd
import calendar
from datetime import datetime

PREBOOK_FILE = "prebook_schedule.xlsx"
OUT_FILE = "monthly_prebook_output.txt"

TARGET_YEAR = 2026
TARGET_MONTH = 5

WEEKDAY_CN = {
    0: "星期一",
    1: "星期二",
    2: "星期三",
    3: "星期四",
    4: "星期五",
    5: "星期六",
    6: "星期日",
}

MONTH_CN = {
    1: "一", 2: "二", 3: "三", 4: "四",
    5: "五", 6: "六", 7: "七", 8: "八",
    9: "九", 10: "十", 11: "十一", 12: "十二"
}


def clean_time(x):
    if pd.isna(x) or str(x).strip() == "":
        return ""
    return str(x).strip()


def format_name_time(row):
    name = str(row["姓名"]).strip()
    start = clean_time(row.get("开始时间", ""))
    end = clean_time(row.get("结束时间", ""))

    if start and end:
        return f"{name} {start}~{end}"
    else:
        return name


def main():
    df = pd.read_excel(PREBOOK_FILE, sheet_name="预报名")
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期"])

    df = df[
        (df["日期"].dt.year == TARGET_YEAR) &
        (df["日期"].dt.month == TARGET_MONTH)
    ].copy()

    lines = []

    lines.append(f"{TARGET_YEAR}年{MONTH_CN[TARGET_MONTH]}月份值班表")
    lines.append("")
    lines.append("每日值班义工")
    lines.append("")
    lines.append("🟧 卫生 ：2~3位")
    lines.append("🟧 星期一至星期五")
    lines.append("全日值班：2~4位义工")
    lines.append("🟧 星期六和星期日")
    lines.append("全日值班：2~6位义工")
    lines.append("")

    days_in_month = calendar.monthrange(TARGET_YEAR, TARGET_MONTH)[1]

    for day in range(1, days_in_month + 1):
        date_obj = datetime(TARGET_YEAR, TARGET_MONTH, day)
        weekday = WEEKDAY_CN[date_obj.weekday()]

        day_df = df[df["日期"].dt.day == day].copy()
        job_text = day_df["岗位"].astype(str)

        # ===== 分类 =====
        hygiene_df = day_df[job_text.str.contains("卫生", na=False)]

        full_day_df = day_df[job_text.str.contains("全日", na=False)]

        normal_duty_df = day_df[
            job_text.str.contains("值班|观音堂|活动中心", na=False)
            & ~job_text.str.contains("全日", na=False)
        ]

        # ===== 卫生名单 =====
        hygiene_names = []
        for _, row in hygiene_df.iterrows():
            name = str(row["姓名"]).strip()
            if name and name not in hygiene_names:
                hygiene_names.append(name)

        # ===== 全日值班名单 =====
        full_day_names = []
        for _, row in full_day_df.iterrows():
            name = str(row["姓名"]).strip()
            if name and name not in full_day_names:
                full_day_names.append(name)

        # ===== 普通/分段值班 =====
        normal_lines = []
        for _, row in normal_duty_df.iterrows():
            text = format_name_time(row)
            if text and text not in normal_lines:
                normal_lines.append(text)

        # ===== 输出日期 =====
        lines.append(f"{day}/{TARGET_MONTH}/{TARGET_YEAR}    {weekday}")

        # ===== 输出卫生 =====
        if hygiene_names:
            lines.append("值日卫生：" + "、".join(hygiene_names))
        else:
            lines.append("值日卫生：")

        # ===== 输出值班 =====
        if full_day_names:
            lines.append("全日值班：" + "、".join(full_day_names))
        else:
            lines.append("全日值班：")

        # 分段值班永远写在下面
        if normal_lines:
            lines.extend(normal_lines)

        lines.append("")

    output = "\n".join(lines)

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.write(output)

    print(f"已输出：{OUT_FILE}")


if __name__ == "__main__":
    main()