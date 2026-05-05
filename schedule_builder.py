import os
import re
from datetime import datetime

import pandas as pd

from lunar_rules import get_special_day_info, get_next_day_remove_info

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

FIXED_FILE = os.path.join(BASE_DIR, "fixed_schedule.xlsx")
SIGNUP_FILE = os.path.join(BASE_DIR, "signup_input.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "schedule_output.txt")
HISTORY_FILE = os.path.join(BASE_DIR, "assignment_history.xlsx")
HISTORY_SHEET = "history"
PREBOOK_FILE = os.path.join(BASE_DIR, "prebook_schedule.xlsx")


def load_prebook_input(target_date_str):
    if not os.path.exists(PREBOOK_FILE):
        return pd.DataFrame(columns=["姓名", "岗位", "开始时间", "结束时间", "优先岗位", "备注"])

    df = pd.read_excel(PREBOOK_FILE, sheet_name="预报名")
    df.columns = df.columns.astype(str).str.strip()

    need_cols = ["姓名", "岗位", "日期", "开始时间", "结束时间"]
    missing = [c for c in need_cols if c not in df.columns]
    if missing:
        raise ValueError(f"prebook_schedule.xlsx 缺少栏位：{missing}")

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce").dt.strftime("%Y-%m-%d")
    df = df[df["日期"] == target_date_str].copy()

    if df.empty:
        return pd.DataFrame(columns=["姓名", "岗位", "开始时间", "结束时间", "优先岗位", "备注"])

    df["姓名"] = df["姓名"].astype(str).str.strip()
    df["岗位"] = df["岗位"].astype(str).str.strip().map(normalize_job_name)
    df["开始时间"] = df["开始时间"].astype(str).str.strip()
    df["结束时间"] = df["结束时间"].astype(str).str.strip()
    df["优先岗位"] = ""
    if "备注" not in df.columns:
        df["备注"] = ""
    else:
        df["备注"] = df["备注"].astype(str).str.strip()

    return df[["姓名", "岗位", "开始时间", "结束时间", "优先岗位", "备注"]].copy()

def merge_signup_and_prebook(signup_df, prebook_df):
    if signup_df is None or signup_df.empty:
        return prebook_df.copy()

    if prebook_df is None or prebook_df.empty:
        return signup_df.copy()

    merged = pd.concat([signup_df, prebook_df], ignore_index=True)

    merged = merged.drop_duplicates(
        subset=["姓名", "岗位", "开始时间", "结束时间"],
        keep="first"
    ).reset_index(drop=True)

    return merged

def save_assignment_history(date_obj, result):
    records = []

    def add_records(names, job):
        for n in names:
            if n:
                records.append({
                    "日期": date_obj.strftime("%Y-%m-%d"),
                    "姓名": n,
                    "岗位": job
                })

    # ===== 值班 =====
    for key in ["橙观音堂", "橙活动中心", "黄观音堂", "黄活动中心"]:
        for item in result[key]:
            name = item[0]
            job = "观音堂" if "观音堂" in key else "活动中心"
            add_records([name], job)

    # ===== 卫生 =====
    add_records(result["佛堂卫生"], "佛堂卫生")
    add_records(result["二楼卫生"], "二楼卫生")
    add_records(result["楼梯卫生"], "楼梯卫生")

    # ===== 佛台 / 供台（可选，一起记下来也行）=====
    add_records(result["整理佛台"], "整理佛台")
    add_records(result["设师父供台"], "设师父供台")

    new_df = pd.DataFrame(records)

    if new_df.empty:
        return

    if os.path.exists(HISTORY_FILE):
        try:
            old_df = pd.read_excel(HISTORY_FILE, sheet_name=HISTORY_SHEET)
            df = pd.concat([old_df, new_df], ignore_index=True)
        except Exception:
            df = new_df
    else:
        df = new_df

    with pd.ExcelWriter(HISTORY_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=HISTORY_SHEET, index=False)

WEEKDAY_MAP = {
    0: "星期一",
    1: "星期二",
    2: "星期三",
    3: "星期四",
    4: "星期五",
    5: "星期六",
    6: "星期日",
}

DEFAULT_TIME = {
    "佛台": ("08:00", "10:00"),
    "整理佛台": ("08:00", "10:00"),
    "佛堂卫生": ("08:00", "10:00"),
    "二楼卫生": ("08:00", "10:00"),
    "楼梯卫生": ("08:00", "10:00"),
    "设师父供台": ("06:00", "08:00"),
    "师父供台": ("06:00", "08:00"),
    "观音堂": ("10:00", "14:00"),
    "活动中心": ("10:00", "14:00"),
    "绿观音堂": ("08:00", "10:00"),
    "绿活动中心": ("08:00", "10:00"),
}

SPECIAL_DEFAULT_TIME = {
    "佛台": ("06:00", "08:00"),
    "整理佛台": ("06:00", "08:00"),
    "佛堂卫生": ("06:00", "08:00"),
    "二楼卫生": ("06:00", "08:00"),
    "楼梯卫生": ("06:00", "08:00"),
    "设师父供台": ("06:00", "08:00"),
    "师父供台": ("06:00", "08:00"),
    "观音堂": ("10:00", "14:00"),
    "活动中心": ("10:00", "14:00"),
    "绿观音堂": ("08:00", "10:00"),
    "绿活动中心": ("08:00", "10:00"),
}

HISTORY_FILE = os.path.join(BASE_DIR, "assignment_history.xlsx")


def load_last_assignment():
    if not os.path.exists(HISTORY_FILE):
        return {}

    df = pd.read_excel(HISTORY_FILE, sheet_name="history")

    if df.empty:
        return {}

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期"])

    last = df.sort_values("日期").groupby("姓名").tail(1)

    return dict(zip(last["姓名"], last["岗位"]))

def save_assignment_history(date_obj, result):
    records = []

    def add_records(names, job):
        for n in names:
            records.append({
                "日期": date_obj.strftime("%Y-%m-%d"),
                "姓名": n,
                "岗位": job
            })

    # ===== 值班 =====
    for key in ["橙观音堂", "橙活动中心", "黄观音堂", "黄活动中心"]:
        for item in result[key]:
            name = item[0]
            job = "观音堂" if "观音堂" in key else "活动中心"
            add_records([name], job)

    # ===== 卫生 =====
    add_records(result["佛堂卫生"], "佛堂卫生")
    add_records(result["二楼卫生"], "二楼卫生")
    add_records(result["楼梯卫生"], "楼梯卫生")

    new_df = pd.DataFrame(records)

    if os.path.exists(HISTORY_FILE):
        old_df = pd.read_excel(HISTORY_FILE, sheet_name="history")
        df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        df = new_df

    with pd.ExcelWriter(HISTORY_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="history", index=False)

def load_yesterday_hygiene_assignment(date_obj):
    """
    读取昨天每个人的卫生岗位：
    返回格式：
    {
        "郑筱頵": "佛堂卫生",
        "陈彩群": "二楼卫生"
    }
    """
    if not os.path.exists(HISTORY_FILE):
        return {}

    df = pd.read_excel(HISTORY_FILE, sheet_name="history")

    if df.empty:
        return {}

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期"])

    yesterday = pd.to_datetime(date_obj) - pd.Timedelta(days=1)

    hygiene_jobs = ["佛堂卫生", "二楼卫生", "楼梯卫生"]

    df = df[
        (df["日期"].dt.date == yesterday.date()) &
        (df["岗位"].astype(str).isin(hygiene_jobs))
    ].copy()

    if df.empty:
        return {}

    df["姓名"] = df["姓名"].astype(str).str.strip()
    df["岗位"] = df["岗位"].astype(str).str.strip()

    last = df.groupby("姓名").tail(1)

    return dict(zip(last["姓名"], last["岗位"]))

def assign_hygiene_no_same_as_yesterday(hygiene_names, date_obj):
    jobs = ["佛堂卫生", "二楼卫生", "楼梯卫生"]
    no_fotang_names = {"葉荔銖", "葉茘銖"}

    yesterday_map = load_yesterday_hygiene_assignment(date_obj)

    result = {
        "佛堂卫生": [],
        "二楼卫生": [],
        "楼梯卫生": []
    }

    used = set()

    for job in jobs:
        candidates = []

        for name in hygiene_names:
            name = str(name).strip()

            if not name or name in used:
                continue

            if job == "佛堂卫生" and name in no_fotang_names:
                continue

            # 核心规则：不要做昨天同一个岗位
            if yesterday_map.get(name) != job:
                candidates.append(name)

        # 如果没人可选，才放宽允许重复昨天岗位
        if not candidates:
            for name in hygiene_names:
                name = str(name).strip()

                if not name or name in used:
                    continue

                if job == "佛堂卫生" and name in no_fotang_names:
                    continue

                candidates.append(name)

        if candidates:
            chosen = candidates[0]
            result[job].append(chosen)
            used.add(chosen)

    return result

def parse_time_str(t):
    if pd.isna(t):
        return None
    s = str(t).strip()
    if not s:
        return None
    return s


def get_weekday_name(date_obj):
    return WEEKDAY_MAP[date_obj.weekday()]


def get_latest_date_from_signup():
    df = pd.read_excel(SIGNUP_FILE, sheet_name="报名")
    df.columns = df.columns.astype(str).str.strip()

    if "日期" not in df.columns:
        raise ValueError("signup_input.xlsx 的【报名】sheet 缺少【日期】栏位")

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期"])

    if df.empty:
        raise ValueError("signup_input.xlsx 的【报名】sheet 里没有有效日期")

    latest_date = df["日期"].max()
    return latest_date.to_pydatetime()


def load_fixed_schedule():
    xls = pd.ExcelFile(FIXED_FILE)
    sheets = xls.sheet_names

    print("检测到 fixed_schedule.xlsx 工作表：", sheets)

    if "佛台固定" in sheets:
        buddhist_df = pd.read_excel(FIXED_FILE, sheet_name="佛台固定")
        buddhist_df.columns = buddhist_df.columns.astype(str).str.strip()
        print("佛台固定栏位：", list(buddhist_df.columns))
    else:
        raise ValueError("❌ 必须有【佛台固定】sheet")

    if "卫生固定" in sheets:
        cleaning_df = pd.read_excel(FIXED_FILE, sheet_name="卫生固定")
        cleaning_df.columns = cleaning_df.columns.astype(str).str.strip()
        print("卫生固定栏位：", list(cleaning_df.columns))
    else:
        print("⚠️ 没有【卫生固定】，将全部由报名决定")
        cleaning_df = pd.DataFrame(columns=["星期", "佛堂卫生", "二楼卫生", "楼梯卫生1", "楼梯卫生2"])

    return buddhist_df, cleaning_df

def fix_time_range(text):
    text = fix_time_format(text)

    # 1030-1230pm → 10:30am~12:30pm
    text = re.sub(
        r"(\d{3,4})(am|pm)?\s*[-~]\s*(\d{3,4})(am|pm)",
        lambda m: f"{fix_time_format(m.group(1)+ (m.group(2) or 'am'))}~{fix_time_format(m.group(3)+m.group(4))}",
        text,
        flags=re.IGNORECASE
    )

    return text

import re

def fix_time_format(text):
    text = str(text)

    def repl(m):
        num = m.group(1)
        ap = m.group(2).lower()

        if len(num) == 4:
            hour = int(num[:2])
            minute = int(num[2:])
        elif len(num) == 3:
            hour = int(num[:1])
            minute = int(num[1:])
        else:
            return m.group(0)

        if hour < 1 or hour > 12 or minute > 59:
            return m.group(0)

        return f"{hour}:{minute:02d}{ap}"

    return re.sub(
        r"(?<!\d)(\d{3,4})\s*(am|pm)(?![a-zA-Z])",
        repl,
        text,
        flags=re.IGNORECASE
    )


def normalize_time_text(text):
    text = str(text).strip()

    text = text.replace("～", "~").replace("—", "-").replace("–", "-").replace("－", "-")

    # 1030am -> 10:30am
    text = fix_time_format(text)

    # 黄丽玲 10 - 2 pm -> 黄丽玲 10am~2pm
    def repl_range(m):
        h1 = int(m.group(1))
        h2 = int(m.group(2))
        ap2 = m.group(3).lower()

        if ap2 == "pm" and h1 >= 7 and h2 <= 6:
            ap1 = "am"
        else:
            ap1 = ap2

        return f"{h1}{ap1}~{h2}{ap2}"

    text = re.sub(
        r"(?<!\d)(\d{1,2})\s*[-~]\s*(\d{1,2})\s*(am|pm)(?![a-zA-Z])",
        repl_range,
        text,
        flags=re.IGNORECASE
    )

    return text

def fix_time_range_v2(text):
    """
    修复：
    10 - 2 pm → 10am~2pm
    10-2pm → 10am~2pm
    10 - 2pm → 10am~2pm
    """

    text = str(text)

    def repl(m):
        h1 = int(m.group(1))
        h2 = int(m.group(2))
        ap2 = m.group(3).lower()

        # 推断第一个时间
        ap1 = "am" if h1 < h2 else ap2

        return f"{h1}{ap1}~{h2}{ap2}"

    text = re.sub(
        r"\b(\d{1,2})\s*[-~]\s*(\d{1,2})\s*(am|pm)\b",
        repl,
        text,
        flags=re.IGNORECASE
    )

    return text


def load_signup_input(target_date_str):
    df = pd.read_excel(SIGNUP_FILE, sheet_name="报名")
    df.columns = df.columns.astype(str).str.strip()

    if "日期" not in df.columns:
        raise ValueError("❌ 报名表需要【日期】栏位")

    if "报名原文" not in df.columns:
        raise ValueError("❌ 报名表需要【报名原文】栏位")

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce").dt.strftime("%Y-%m-%d")
    df = df[df["日期"] == target_date_str].copy()

    records = []

    for idx, row in df.iterrows():
        raw_text = str(row.get("报名原文", "")).strip()
        if not raw_text:
            continue

        # 支持一格多行
        lines = raw_text.split("\n")

        for line in lines:
            line = normalize_time_text(line.strip())
            if not line:
                continue

            parsed_list = parse_signup_line_multi(line)
            if not parsed_list:
                print(f"⚠️ 无法解析：{line}")
                continue

            records.extend(parsed_list)

    result_df = pd.DataFrame(records)

    if result_df.empty:
        return pd.DataFrame(columns=["姓名", "岗位", "开始时间", "结束时间", "优先岗位", "备注"])

    return result_df

def normalize_time_range(text):
    import re

    if not text:
        return None, None

    s = str(text).lower().strip()
    s = s.replace("～", "~").replace("-", "~")
    s = s.replace(".", ":")
    s = re.sub(r"\s+", " ", s)

    # 12:00pm~3:00pm
    m = re.search(r"(\d{1,2}:\d{2}(?:am|pm))\s*~\s*(\d{1,2}:\d{2}(?:am|pm))", s)
    if m:
        return m.group(1), m.group(2)

    # 4~6pm
    m = re.search(r"(\d{1,2})\s*~\s*(\d{1,2})(am|pm)", s)
    if m:
        start_hour = int(m.group(1))
        end_hour = int(m.group(2))
        suffix = m.group(3)

        # 特别规则：
        # 10-12pm / 9-12pm / 11-12pm 这种，前一个通常应视为 am
        if suffix == "pm" and start_hour < end_hour and end_hour == 12:
            return f"{start_hour}:00am", f"{end_hour}:00pm"

        return f"{start_hour}:00{suffix}", f"{end_hour}:00{suffix}"

    # 12pm 5pm
    m = re.findall(r"(\d{1,2})(am|pm)", s)
    if len(m) >= 2:
        return f"{m[0][0]}:00{m[0][1]}", f"{m[1][0]}:00{m[1][1]}"

    return None, None

def get_fixed_people(date_obj, buddhist_df, cleaning_df):
    weekday = get_weekday_name(date_obj)

    result = {
        "整理佛台": [],
        "佛堂卫生": [],
        "二楼卫生": [],
        "楼梯卫生": [],
    }

    if not buddhist_df.empty:
        buddhist_df = buddhist_df.copy()
        buddhist_df.columns = buddhist_df.columns.astype(str).str.strip()

        if "星期" not in buddhist_df.columns:
            raise ValueError(f"❌【佛台固定】缺少栏位：星期。当前栏位：{list(buddhist_df.columns)}")

        buddhist_row = buddhist_df[buddhist_df["星期"].astype(str).str.strip() == weekday]

        if not buddhist_row.empty:
            row = buddhist_row.iloc[0]
            name_cols = [c for c in buddhist_df.columns if str(c).startswith("姓名")]

            for col in name_cols:
                if pd.notna(row[col]) and str(row[col]).strip():
                    result["整理佛台"].append(str(row[col]).strip())

    if not cleaning_df.empty:
        cleaning_df = cleaning_df.copy()
        cleaning_df.columns = cleaning_df.columns.astype(str).str.strip()

        if "星期" in cleaning_df.columns:
            cleaning_row = cleaning_df[cleaning_df["星期"].astype(str).str.strip() == weekday]

            if not cleaning_row.empty:
                row = cleaning_row.iloc[0]

                if "佛堂卫生" in cleaning_df.columns and pd.notna(row.get("佛堂卫生")) and str(row.get("佛堂卫生")).strip():
                    result["佛堂卫生"].append(str(row.get("佛堂卫生")).strip())

                if "二楼卫生" in cleaning_df.columns and pd.notna(row.get("二楼卫生")) and str(row.get("二楼卫生")).strip():
                    result["二楼卫生"].append(str(row.get("二楼卫生")).strip())

                for col in ["楼梯卫生1", "楼梯卫生2"]:
                    if col in cleaning_df.columns and pd.notna(row.get(col)) and str(row.get(col)).strip():
                        result["楼梯卫生"].append(str(row.get(col)).strip())

    return result


def normalize_job_name(job):
    job = str(job).strip()

    mapping = {
        "佛台": "整理佛台",
        "整理佛台": "整理佛台",
        "卫生": "卫生",
        "佛堂卫生": "佛堂卫生",
        "二楼卫生": "二楼卫生",
        "楼梯卫生": "楼梯卫生",
        "设师父供台": "设师父供台",
        "师父供台": "设师父供台",
        "观音堂": "观音堂",
        "活动中心": "活动中心",
        "绿观音堂": "绿观音堂",
        "绿活动中心": "绿活动中心",
    }
    return mapping.get(job, job)

def parse_signup_line_multi(text):
    text = str(text).strip()
    if not text:
        return []

    text = text.replace("　", " ")
    text = re.sub(r"\s+", " ", text)

    # 统一常见字
    text = text.replace("衛生", "卫生")

    # ===== 全日规则 =====
    is_full_day = "全日" in text

    # 全日先去掉，避免被吃进名字
    clean_text = text.replace("全日", "").strip()

    # 抓时间
    start, end = normalize_time_range(clean_text)

    # 如果写了全日但没写时间，默认 10am~6pm
    if is_full_day and (not start or not end):
        start, end = "10:00am", "6:00pm"

    role_text = clean_text
    if start:
        role_text = role_text.replace(start, "")
    if end:
        role_text = role_text.replace(end, "")
    role_text = role_text.replace("~", " ").replace("-", " ")
    role_text = re.sub(r"\s+", " ", role_text).strip()

    role_text = role_text.replace("&", "/").replace("＆", "/").replace("和", "/")

    job_keywords = ["活动中心", "观音堂", "卫生", "佛台", "值班"]

    found = []
    for kw in job_keywords:
        start_pos = 0
        while True:
            pos = role_text.find(kw, start_pos)
            if pos == -1:
                break
            found.append((pos, kw))
            start_pos = pos + len(kw)

    found.sort(key=lambda x: x[0])

    if not found:
        m = re.match(r"^([A-Za-z\u4e00-\u9fff·•.]+)", role_text)
        if not m:
            return []
        name = m.group(1).rstrip(".").strip()
        return [{
            "姓名": name,
            "岗位": "值班",
            "开始时间": start,
            "结束时间": end,
            "优先岗位": "",
            "备注": "",
        }]

    first_pos = found[0][0]
    name = role_text[:first_pos].strip().rstrip(".")

    if not name:
        return []

    records = []

    for _, kw in found:
        if kw == "卫生":
            job = "卫生"
            job_start, job_end = None, None
        elif kw == "观音堂":
            job = "观音堂"
            job_start, job_end = start, end
        elif kw == "活动中心":
            job = "活动中心"
            job_start, job_end = start, end
        elif kw == "佛台":
            job = "整理佛台"
            job_start, job_end = start, end
        elif kw == "值班":
            job = "值班"
            job_start, job_end = start, end
        else:
            continue

        records.append({
            "姓名": name,
            "岗位": job,
            "开始时间": job_start,
            "结束时间": job_end,
            "优先岗位": "",
            "备注": "",
        })

    dedup = []
    seen = set()
    for r in records:
        key = (r["姓名"], r["岗位"], r["开始时间"], r["结束时间"])
        if key not in seen:
            seen.add(key)
            dedup.append(r)

    return dedup

def parse_signup_line(text):
    text = str(text).strip()
    if not text:
        return None

    text = text.replace("　", " ")
    text = re.sub(r"\s+", " ", text)

    # 先定义岗位关键词
    job_keywords = ["活动中心", "观音堂", "卫生", "佛台", "值班"]

    found_job = None
    found_pos = None

    for kw in job_keywords:
        pos = text.find(kw)
        if pos >= 0:
            if found_pos is None or pos < found_pos:
                found_job = kw
                found_pos = pos

    # 时间先抓出来
    start, end = normalize_time_range(text)

    if found_job is not None:
        before_job = text[:found_pos].strip()
        after_job = text[found_pos + len(found_job):].strip()

        # 名字通常在岗位前面；如果岗位前面还有时间，要去掉
        name = before_job

        # 去掉名字后面可能混进去的时间
        name = re.sub(r"\d{1,2}[:.]?\d{0,2}\s*(?:am|pm)?\s*[-~]?\s*\d{0,2}[:.]?\d{0,2}\s*(?:am|pm)?", "", name, flags=re.I).strip()
        name = name.rstrip(".").strip()

        rest = f"{found_job} {after_job}".strip()
    else:
        # 没岗位关键词时，抓开头名字
        m = re.match(r"^([A-Za-z\u4e00-\u9fff·•.]+)", text)
        if not m:
            return None

        name = m.group(1).rstrip(".").strip()
        rest = text[len(m.group(0)):].strip()
        found_job = None

    # 最终岗位映射
    if found_job == "卫生":
        job = "卫生"
    elif found_job == "观音堂":
        job = "观音堂"
    elif found_job == "活动中心":
        job = "活动中心"
    elif found_job == "佛台":
        job = "整理佛台"
    elif found_job == "值班":
        job = "值班"
    else:
        # 没写岗位时，默认值班
        job = "值班"

    return {
        "姓名": name,
        "岗位": job,
        "开始时间": start,
        "结束时间": end,
        "优先岗位": "",
        "备注": "",
    }

def auto_assign_cleaning_roles(result, signup_df):
    cleaning_people = []

    for _, r in signup_df.iterrows():
        if str(r.get("岗位", "")).strip() == "卫生":
            name = str(r.get("姓名", "")).strip()
            if name and name not in cleaning_people:
                cleaning_people.append(name)

    if not cleaning_people:
        return result

    last_job_map = load_last_assignment()

    def pick_job(name, options):
        last = last_job_map.get(name)
        for o in options:
            if o != last:
                return o
        return options[0]

    # 先清空，避免重复
    result["佛堂卫生"] = []
    result["二楼卫生"] = []
    result["楼梯卫生"] = []

    # ===== 1人 =====
    if len(cleaning_people) == 1:
        name = cleaning_people[0]

        if name == "葉荔銖":
            job = pick_job(name, ["二楼卫生", "楼梯卫生"])
        else:
            job = pick_job(name, ["佛堂卫生", "二楼卫生", "楼梯卫生"])

        result[job] = [name]
        return result

    # ===== 2人 =====
    if len(cleaning_people) == 2:
        p1, p2 = cleaning_people

        # 葉荔銖不排佛堂卫生
        if p1 == "葉荔銖" or p2 == "葉荔銖":
            late_person = "葉荔銖"
            other = p2 if p1 == "葉荔銖" else p1

            result["佛堂卫生"] = [other]
            result["二楼卫生"] = [late_person]
            result["楼梯卫生"] = [other, late_person]
            return result

        # 普通2人：一人佛堂，一人二楼，楼梯联合
        job1 = pick_job(p1, ["佛堂卫生", "二楼卫生"])
        job2 = "二楼卫生" if job1 == "佛堂卫生" else "佛堂卫生"

        result[job1] = [p1]
        result[job2] = [p2]
        result["楼梯卫生"] = [p1, p2]
        return result

    # ===== 3人或以上 =====
    remaining = cleaning_people[:]

    # 葉荔銖特殊规则：优先二楼，不排佛堂
    if "葉荔銖" in remaining:
        result["二楼卫生"] = ["葉荔銖"]
        remaining.remove("葉荔銖")

    # 佛堂卫生
    if remaining:
        fotang_person = remaining.pop(0)
        result["佛堂卫生"] = [fotang_person]
    else:
        fotang_person = None

    # 二楼卫生（如果还没安排）
    if not result["二楼卫生"] and remaining:
        second_floor_person = remaining.pop(0)
        result["二楼卫生"] = [second_floor_person]
    else:
        second_floor_person = result["二楼卫生"][0] if result["二楼卫生"] else None

    # 楼梯卫生：有第三人就给第三人单独做
    if remaining:
        result["楼梯卫生"] = [remaining.pop(0)]
    else:
        # 理论上不会进来，因为这里是3人或以上
        stair_fallback = []
        for p in cleaning_people:
            if p not in [fotang_person, second_floor_person]:
                stair_fallback.append(p)
        result["楼梯卫生"] = stair_fallback[:1]

    return result

def normalize_signup(df, is_special_day=False):
    rows = []

    for _, r in df.iterrows():
        name = str(r.get("姓名", "")).strip()
        wishes = str(r.get("岗位意向", "")).strip()
        start = parse_time_str(r.get("开始时间", ""))
        end = parse_time_str(r.get("结束时间", ""))
        priority = str(r.get("优先岗位", "")).strip()
        note = str(r.get("备注", "")).strip()

        if not name or not wishes:
            continue

        jobs = [x.strip() for x in wishes.split(",") if x.strip()]
        for raw_job in jobs:
            job = normalize_job_name(raw_job)

            job_start = start
            job_end = end

            if not job_start or not job_end:
                if is_special_day and job in SPECIAL_DEFAULT_TIME:
                    job_start, job_end = SPECIAL_DEFAULT_TIME[job]
                elif job in DEFAULT_TIME:
                    job_start, job_end = DEFAULT_TIME[job]

            rows.append({
                "姓名": name,
                "岗位": job,
                "开始时间": job_start,
                "结束时间": job_end,
                "优先岗位": priority,
                "备注": note,
            })

    return pd.DataFrame(rows)


def classify_shift_slot(start_time):
    s = parse_time_str(start_time)
    if not s:
        return "橙"

    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            dt = datetime.strptime(s, fmt)
            hm = dt.hour * 60 + dt.minute
            if 8 * 60 <= hm < 10 * 60:
                return "绿"
            elif 10 * 60 <= hm < 14 * 60:
                return "橙"
            else:
                return "黄"
        except ValueError:
            pass

    s2 = s.lower().replace(" ", "")
    if "am" in s2 or "pm" in s2:
        for fmt in ("%I:%M%p", "%I.%M%p"):
            try:
                dt = datetime.strptime(s2, fmt)
                hm = dt.hour * 60 + dt.minute
                if 8 * 60 <= hm < 10 * 60:
                    return "绿"
                elif 10 * 60 <= hm < 14 * 60:
                    return "橙"
                else:
                    return "黄"
            except ValueError:
                pass

    return "橙"


def assign_jobs(fixed_people, signup_df, special_info):
    result = {
        "整理佛台": fixed_people["整理佛台"][:],
        "佛堂卫生": fixed_people["佛堂卫生"][:],
        "二楼卫生": fixed_people["二楼卫生"][:],
        "楼梯卫生": fixed_people["楼梯卫生"][:],
        "设师父供台": [],
        "绿观音堂": [],
        "绿活动中心": [],
        "橙观音堂": [],
        "橙活动中心": [],
        "黄观音堂": [],
        "黄活动中心": [],
    }

    for _, r in signup_df.iterrows():
        name = r["姓名"]
        job = r["岗位"]
        start = r["开始时间"]
        end = r["结束时间"]

        if job == "整理佛台":
            if name not in result["整理佛台"]:
                result["整理佛台"].append(name)

        elif job == "佛堂卫生":
            if name not in result["佛堂卫生"]:
                result["佛堂卫生"].append(name)

        elif job == "二楼卫生":
            if name not in result["二楼卫生"]:
                result["二楼卫生"].append(name)

        elif job == "楼梯卫生":
            if name not in result["楼梯卫生"]:
                result["楼梯卫生"].append(name)

        elif job == "设师父供台":
            if name not in result["设师父供台"]:
                result["设师父供台"].append(name)

        elif job == "绿观音堂":
            result["绿观音堂"].append(name)

        elif job == "绿活动中心":
            result["绿活动中心"].append(name)

        elif job in ["观音堂", "活动中心", "值班"]:
            # 如果没有时间，当作全天或默认时间
            if not start or not end:
                start = "10:00am"
                end = "2:00pm"

            slot = classify_shift_slot(start)

            if job == "值班":
                # 值班 → 先全部丢去观音堂，再让 auto_assign_duty 分
                key = f"{slot}观音堂"
            else:
                key = f"{slot}{job}"

            result[key].append((name, start, end))

    # ✅ 循环结束后才做自动分配
    result = auto_assign_cleaning_roles(result, signup_df)
    result = auto_assign_duty(result)

    return result
    
def auto_assign_duty(result):
    def parse_min(t):
        if not t:
            return None

        s = str(t).strip().lower().replace(" ", "")

        # 2pm -> 2:00pm
        if re.match(r"^\d{1,2}(am|pm)$", s):
            s = s.replace("am", ":00am").replace("pm", ":00pm")

        for fmt in ("%H:%M", "%H:%M:%S", "%I:%M%p"):
            try:
                dt = datetime.strptime(s, fmt)
                return dt.hour * 60 + dt.minute
            except:
                pass

        return None

    def min_to_ampm(m):
        hour = m // 60
        minute = m % 60
        suffix = "am" if hour < 12 else "pm"
        hour12 = hour % 12
        if hour12 == 0:
            hour12 = 12
        return f"{hour12}:{minute:02d}{suffix}"

    def assign_balanced(group1, group2, item):
        name = item[0]

        # 如果这个人已经在 group1 出现过，下一段优先去 group2
        if any(x[0] == name for x in group1) and not any(x[0] == name for x in group2):
            group2.append(item)
            return

        # 如果这个人已经在 group2 出现过，下一段优先去 group1
        if any(x[0] == name for x in group2) and not any(x[0] == name for x in group1):
            group1.append(item)
            return

        # 普通情况：人数平衡
        if len(group1) <= len(group2):
            group1.append(item)
        else:
            group2.append(item)

    def assign_afternoon(item):
        name = item[0]

        # 1）如果这个人上午在活动中心，下午优先去观音堂
        if any(x[0] == name for x in result["橙活动中心"]):
            result["黄观音堂"].append(item)
            return

        # 2）如果这个人上午在观音堂，下午优先去活动中心
        if any(x[0] == name for x in result["橙观音堂"]):
            result["黄活动中心"].append(item)
            return

        # 3）普通情况：下午观音堂优先，但保持平衡
        if len(result["黄观音堂"]) <= len(result["黄活动中心"]):
            result["黄观音堂"].append(item)
        else:
            result["黄活动中心"].append(item)

    duty_pool = []

    for key in ["橙观音堂", "橙活动中心", "黄观音堂", "黄活动中心"]:
        for item in result[key]:
            name, start, end = item
            duty_pool.append({
                "name": name,
                "start": parse_min(start),
                "end": parse_min(end),
                "raw": item
            })

    # 清空重新排
    result["橙观音堂"] = []
    result["橙活动中心"] = []
    result["黄观音堂"] = []
    result["黄活动中心"] = []

    for p in duty_pool:
        if p["start"] is None or p["end"] is None:
            continue

        name = p["name"]
        s = p["start"]
        e = p["end"]

        if e <= s:
            continue

        name = p["name"]
        start_txt = min_to_ampm(s)
        end_txt = min_to_ampm(e)
        duration = e - s

        # ===== 长班（≥4小时）直接锁岗位，但跨2pm要切开 =====
        if duration >= 4 * 60:

            # 如果跨过2pm，例如 10~6 / 11~4
            if s < 14 * 60 < e:
                morning_item = (name, start_txt, "2:00pm")
                afternoon_item = (name, "2:00pm", end_txt)

                # 上午先分
                if len(result["橙观音堂"]) <= len(result["橙活动中心"]):
                    result["橙观音堂"].append(morning_item)
                else:
                    result["橙活动中心"].append(morning_item)

                # 下午一定要走统一逻辑
                assign_afternoon(afternoon_item)

                continue

            # 没跨2pm的长班，例如 10~2
            if e <= 14 * 60:
                if len(result["橙观音堂"]) <= len(result["橙活动中心"]):
                    result["橙观音堂"].append((name, start_txt, end_txt))
                else:
                    result["橙活动中心"].append((name, start_txt, end_txt))

                continue

            # 纯下午长班
            if s >= 14 * 60:
                assign_afternoon((name, start_txt, end_txt))
                continue
        
        # ===== 跨 2pm 短班，例如 1pm~4pm =====
        if s < 14 * 60 < e:
            assign_balanced(
                result["橙观音堂"],
                result["橙活动中心"],
                (name, start_txt, "2:00pm")
            )

            # 下午优先补观音堂
            assign_afternoon((name, "2:00pm", end_txt))

            continue

        # ===== 上午 =====
        if e <= 14 * 60:
            assign_balanced(
                result["橙观音堂"],
                result["橙活动中心"],
                (name, start_txt, end_txt)
            )
            continue

        # ===== 下午 =====
        if s >= 14 * 60:
            assign_afternoon((name, start_txt, end_txt))
            continue

    return result

def format_people_inline(names, sep="  "):
    return sep.join(names) if names else ""


def to_ampm_text(t):
    if pd.isna(t):
        return ""

    s = str(t).strip()
    if not s:
        return ""

    s_lower = s.lower().replace(" ", "")

    if "am" in s_lower or "pm" in s_lower:
        return s_lower

    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            dt = datetime.strptime(s, fmt)
            hour = dt.hour
            minute = dt.minute
            suffix = "am" if hour < 12 else "pm"
            hour12 = hour % 12
            if hour12 == 0:
                hour12 = 12
            return f"{hour12}:{minute:02d}{suffix}"
        except ValueError:
            pass

    return s


def format_shift_block(items):
    if not items:
        return ""

    lines = []
    for name, start, end in items:
        start_txt = to_ampm_text(start)
        end_txt = to_ampm_text(end)
        lines.append(name)
        lines.append(f"{start_txt}~{end_txt}")
    return "\n".join(lines)


def build_normal_message(date_obj, arranged, special_info, remove_info):
    weekday = get_weekday_name(date_obj)
    date_text = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"

    msg = f"""师兄们大家好！

{date_text} ({weekday}) 

十点正请安

8:00am~10:00am  或      
8:00am~完成佛台工作 
整理佛台: 
{format_people_inline(arranged["整理佛台"])}

8:00am~10:00am 或 
8:00am~完成卫生工作 
佛堂卫生: {format_people_inline(arranged["佛堂卫生"])}
二楼卫生: {format_people_inline(arranged["二楼卫生"])}
楼梯卫生: {format_people_inline(arranged["楼梯卫生"], sep="/")}
每日卫生义工请注意：清理完卫生之后，请把卫生用具包括吸尘机放回原位（活动中心store里面的小房间）

10:00am~2:00pm 
🟠 观音堂: 
{format_shift_block(arranged["橙观音堂"])}
🟠 活动中心: 
{format_shift_block(arranged["橙活动中心"])}

2:00pm~6:00pm 
🟡 观音堂: 
{format_shift_block(arranged["黄观音堂"])}
🟡 活动中心: 
{format_shift_block(arranged["黄活动中心"])}

观音堂早晚香 由值班义工带领上香。

观音堂续香
佛友可以续香（黑色无烟香），但是要跟着请安词，必须燃烧完了一支香才续香。如有多位佛友要续香，请先让给先到达观音堂的佛友，轮流续香。

值班义工请注意
1）下雨天记得关上烧送小房子的窗口。
2）在离开观音堂之前，请确保把所有的窗口关上。
3）观音堂第一架的冷气坚决不能调。第二架和第三架可以轮流。

另外，请大家多留意义工群信息，以便大家能够团结一致的护持好观音堂。佛子齐心，普度众生。

非常感恩大家！
大家功德无量！
🙏🙏🙏
"""
    return msg


def build_lunar_1_15_message(date_obj, arranged, special_info, remove_info):
    weekday = get_weekday_name(date_obj)
    date_text = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"
    lunar_text = special_info.get("lunar_text", "")

    msg = f"""师兄们，大家好！

{date_text} ({weekday}) 

{lunar_text}

八点正请安 

6:00am~8:00am 或 
6:00am~完成佛台工作 
整理佛台:  
{format_people_inline(arranged["整理佛台"])}

6:00am~8:00am 或
6:00am~完成供台工作 
师父供台: 
{format_people_inline(arranged["设师父供台"])}

6:00am~8:00am 或 
6:00am~完成卫生工作 
佛堂卫生: {format_people_inline(arranged["佛堂卫生"])}
二楼卫生: {format_people_inline(arranged["二楼卫生"])}
楼梯卫生: {format_people_inline(arranged["楼梯卫生"], sep="/")}

每日卫生义工请注意：清理完卫生之后，请把卫生用具包括吸尘机放回原位（活动中心store里面的小房间）

🟢 8:00am~10:00am
每个岗位1位师兄
观音堂: {format_people_inline(arranged["绿观音堂"])}
活动中心: {format_people_inline(arranged["绿活动中心"])}

10:00am~2:00pm 
🟠 观音堂: 
{format_shift_block(arranged["橙观音堂"])}
🟠 活动中心: 
{format_shift_block(arranged["橙活动中心"])}

2:00pm~6:00pm 
🟡 观音堂: 
{format_shift_block(arranged["黄观音堂"])}
🟡 活动中心: 
{format_shift_block(arranged["黄活动中心"])}

观音堂早晚香
由值班义工带领上香。

观音堂续香
佛友可以续香（黑色无烟香），但是要跟着请安词，必须燃烧完了一支香才续香。如有多位佛友要续香，请先让给先到达观音堂的佛友，轮流续香。

值班义工请注意
1）下雨天记得关上烧送小房子的窗口。
2）在离开观音堂之前，请确保把所有的窗口关上。
3）观音堂第一架的冷气坚决不能调。第二架和第三架可以轮流。

另外，请大家多留意义工群信息，以便大家能够团结一致的护持好观音堂。佛子齐心，普度众生。

非常感恩大家！
大家功德无量！
🙏🙏🙏
"""
    return msg


def build_buddhist_festival_message(date_obj, arranged, special_info, remove_info):
    weekday = get_weekday_name(date_obj)
    date_text = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"
    lunar_text = special_info.get("lunar_text", "")
    festival_text = "\n".join(special_info.get("special_names", []))

    remove_notice = ""
    if remove_info["need_remove_today_after_12"]:
        remove_notice = "今日中午12点后请记得撤供台。\n\n"

    msg = f"""师兄们，大家好！

{date_text} ({weekday}) 

{lunar_text}

{festival_text}

八点正请安 

6:00am~8:00am 或 
6:00am~完成佛台工作
整理佛台: 
{format_people_inline(arranged["整理佛台"])}

6:00am~8:00am 或 
6:00am~完成供台工作
设师父供台: 
{format_people_inline(arranged["设师父供台"])}

6:00am~8:00am 或 
6:00am~完成卫生工作 
佛堂卫生: {format_people_inline(arranged["佛堂卫生"])}
二楼卫生: {format_people_inline(arranged["二楼卫生"])}
楼梯卫生: {format_people_inline(arranged["楼梯卫生"], sep="/")}
每日卫生义工请注意：清理完卫生之后，请把卫生用具包括吸尘机放回原位（活动中心store里面的小房间）

8:00am~10:00am
每个岗位1位师兄
观音堂: {format_people_inline(arranged["绿观音堂"])}
活动中心: {format_people_inline(arranged["绿活动中心"])}

10:00am~2:00pm 
🔴 观音堂: 
{format_shift_block(arranged["橙观音堂"])}
🔴 活动中心: 
{format_shift_block(arranged["橙活动中心"])}

2:00pm~6:00pm 
🟡 观音堂: 
{format_shift_block(arranged["黄观音堂"])}
🟡 活动中心: 
{format_shift_block(arranged["黄活动中心"])}

观音堂早晚香
由值班义工带领上香。

观音堂续香
佛友可以续香（黑色无烟香），但是要跟着请安词，必须燃烧完了一支香才续香。如有多位佛友要续香，请先让给先到达观音堂的佛友，轮流续香。

值班义工请注意
1）下雨天记得关上烧送小房子的窗口。
2）在离开观音堂之前，请确保把所有的窗口关上。
3）观音堂第一架的冷气坚决不能调。第二架和第三架可以轮流。

另外，请大家多留意义工群信息，以便大家能够团结一致的护持好观音堂。佛子齐心，普度众生。

{remove_notice}非常感恩大家！
大家功德无量！
🙏🙏🙏
"""
    return msg


def main():
    try:
        date_obj = get_latest_date_from_signup()
        target_date_str = date_obj.strftime("%Y-%m-%d")
        print(f"自动使用日期：{target_date_str}")
    except Exception as e:
        print(f"❌ 读取日期失败：{e}")
        return

    try:
        buddhist_df, cleaning_df = load_fixed_schedule()
    except Exception as e:
        print(f"❌ 读取 fixed_schedule.xlsx 失败：{e}")
        return

    try:
        fixed_people = get_fixed_people(date_obj, buddhist_df, cleaning_df)
        special_info = get_special_day_info(date_obj)
        remove_info = get_next_day_remove_info(date_obj)

        signup_df = load_signup_input(target_date_str)
        prebook_df = load_prebook_input(target_date_str)
        signup_df = merge_signup_and_prebook(signup_df, prebook_df)

        print("=== 合并后的报名资料（含预报名） ===")
        print(signup_df)

        arranged = assign_jobs(fixed_people, signup_df, special_info)

        if special_info["template_type"] == "normal":
            message = build_normal_message(date_obj, arranged, special_info, remove_info)
        elif special_info["template_type"] == "lunar_1_15":
            message = build_lunar_1_15_message(date_obj, arranged, special_info, remove_info)
        elif special_info["template_type"] == "buddhist_festival":
            message = build_buddhist_festival_message(date_obj, arranged, special_info, remove_info)
        else:
            message = build_normal_message(date_obj, arranged, special_info, remove_info)

        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write(message)

        save_assignment_history(date_obj, arranged)

        print("已生成排班文案：", OUTPUT_FILE)
        print("=" * 50)
        print(message)

    except Exception as e:
        print(f"❌ 生成排班失败：{e}")

def run_schedule_for_date(date_str):
    from datetime import datetime

    date_obj = datetime.strptime(date_str, "%Y-%m-%d")

    try:
        buddhist_df, cleaning_df = load_fixed_schedule()
        fixed_people = get_fixed_people(date_obj, buddhist_df, cleaning_df)

        special_info = get_special_day_info(date_obj)
        remove_info = get_next_day_remove_info(date_obj)

        signup_df = load_prebook_input(date_str)

        arranged = assign_jobs(fixed_people, signup_df, special_info)

        if special_info["template_type"] == "normal":
            message = build_normal_message(date_obj, arranged, special_info, remove_info)
        elif special_info["template_type"] == "lunar_1_15":
            message = build_lunar_1_15_message(date_obj, arranged, special_info, remove_info)
        else:
            message = build_buddhist_festival_message(date_obj, arranged, special_info, remove_info)

        save_assignment_history(date_obj, arranged)

        return message

    except Exception as e:
        return f"❌ 排班失败：{e}"

if __name__ == "__main__":
    main()