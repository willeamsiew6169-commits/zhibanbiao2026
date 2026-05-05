import pandas as pd
import calendar
from datetime import datetime
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PREBOOK_FILE = os.path.join(BASE_DIR, "prebook_schedule.xlsx")
OUT_FILE = os.path.join(BASE_DIR, "monthly_prebook_output.txt")

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


def generate_monthly_prebook_message(target_year, target_month):
    target_year = int(target_year)
    target_month = int(target_month)

    if not os.path.exists(PREBOOK_FILE):
        raise FileNotFoundError("找不到 prebook_schedule.xlsx")

    df = pd.read_excel(PREBOOK_FILE, sheet_name="预报名")
    df.columns = df.columns.astype(str).str.strip()

    need_cols = ["日期", "姓名", "岗位", "开始时间", "结束时间"]
    missing = [c for c in need_cols if c not in df.columns]
    if missing:
        raise ValueError(f"prebook_schedule.xlsx 缺少栏位：{missing}")

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期"])

    df = df[
        (df["日期"].dt.year == target_year) &
        (df["日期"].dt.month == target_month)
    ].copy()

    lines = []

    lines.append(f"{target_year}年{MONTH_CN[target_month]}月份值班表")
    lines.append("")
    lines.append("每日值班义工")
    lines.append("")
    lines.append("🟧 卫生 ：2~3位")
    lines.append("🟧 星期一至星期五")
    lines.append("全日值班：2~4位义工")
    lines.append("🟧 星期六和星期日")
    lines.append("全日值班：2~6位义工")
    lines.append("")

    days_in_month = calendar.monthrange(target_year, target_month)[1]

    for day in range(1, days_in_month + 1):
        date_obj = datetime(target_year, target_month, day)
        weekday = WEEKDAY_CN[date_obj.weekday()]

        day_df = df[df["日期"].dt.day == day].copy()
        job_text = day_df["岗位"].astype(str)

        hygiene_df = day_df[job_text.str.contains("卫生", na=False)]
        full_day_df = day_df[job_text.str.contains("全日", na=False)]

        normal_duty_df = day_df[
            job_text.str.contains("值班|观音堂|活动中心", na=False)
            & ~job_text.str.contains("全日", na=False)
        ]

        hygiene_names = []
        for _, row in hygiene_df.iterrows():
            name = str(row["姓名"]).strip()
            if name and name not in hygiene_names:
                hygiene_names.append(name)

        full_day_names = []
        for _, row in full_day_df.iterrows():
            name = str(row["姓名"]).strip()
            if name and name not in full_day_names:
                full_day_names.append(name)

        normal_lines = []
        for _, row in normal_duty_df.iterrows():
            text = format_name_time(row)
            if text and text not in normal_lines:
                normal_lines.append(text)

        lines.append(f"{day}/{target_month}/{target_year}    {weekday}")

        if hygiene_names:
            lines.append("值日卫生：" + "、".join(hygiene_names))
        else:
            lines.append("值日卫生：")

        if full_day_names:
            lines.append("全日值班：" + "、".join(full_day_names))
        else:
            lines.append("全日值班：")

        if normal_lines:
            lines.extend(normal_lines)

        lines.append("")

    output = "\n".join(lines)

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.write(output)

    return output


def main():
    output = generate_monthly_prebook_message(2026, 5)
    print(f"已输出：{OUT_FILE}")
    print(output)


if __name__ == "__main__":
    main()