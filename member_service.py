# member_service.py

import pandas as pd

from pathlib import Path
from utils import normalize_member_id


BASE = Path(__file__).parent
MEMBERS_FILE = BASE / "members.xlsx"


def clean(v):
    return "" if v is None else str(v).strip()


# =========================
# 读取 members（核心）
# =========================
def load_members(path=MEMBERS_FILE):
    df = pd.read_excel(path)

    # 自动识别月份栏
    month_cols = [c for c in df.columns if str(c).startswith("2026-")]

    members = {}

    for _, row in df.iterrows():
        raw_mid = clean(row.get("月费编号"))
        mid = normalize_member_id(raw_mid)
        name = clean(row.get("姓名"))
        ename = clean(row.get("英文名"))

        if not mid:
            continue

        paid_months = [
            m for m in month_cols
            if clean(row.get(m)) in ["1", "✓", "✔", "✅"]
        ]

        members[mid] = {
            "member_id": mid,
            "name": name,
            "ename": ename,
            "months": paid_months
        }

    return members


# =========================
# 查询（月费编号 or 姓名）
# =========================
def query_member(members, keyword):
    keyword = clean(keyword)
    if not keyword:
        return []

    normalized = normalize_member_id(keyword)
    kw = keyword.lower()

    results = []

    for mid, info in members.items():
        name = clean(info.get("name", ""))
        ename = clean(info.get("ename", ""))

        name_simple = (
            name.replace("蕭", "萧")
                .replace("倫", "伦")
                .replace("陳", "陈")
                .replace("黃", "黄")
                .replace("劉", "刘")
                .replace("張", "张")
                .replace("鄭", "郑")
                .replace("鍾", "钟")
        )

        keyword_simple = (
            keyword.replace("蕭", "萧")
                   .replace("倫", "伦")
                   .replace("陳", "陈")
                   .replace("黃", "黄")
                   .replace("劉", "刘")
                   .replace("張", "张")
                   .replace("鄭", "郑")
                   .replace("鍾", "钟")
        )

        matched = (
            normalized == mid
            or keyword == name
            or keyword in name
            or keyword_simple in name_simple
            or kw == ename.lower()
            or kw in ename.lower()
        )

        if matched:
            months = info["months"]
            latest = months[-1] if months else ""

            results.append({
                "member_id": mid,
                "name": name,
                "ename": ename,
                "months": months,
                "latest": latest,
            })

    return results


# =========================
# 获取“供养到几月”（给系统用）
# =========================
def get_latest_month(members, keyword):
    result = query_member(members, keyword)
    if not result:
        return ""

    return result["latest"]


# =========================
# 获取全部未供养名单（很有用）
# =========================
def get_unpaid_members(members, target_month):
    """
    target_month: "2026-05"
    """
    unpaid = []

    for info in members.values():
        if target_month not in info["months"]:
            unpaid.append(info["name"])

    return unpaid


# =========================
# CLI 查询（本地用）
# =========================
def run_cli():
    members = load_members()

    print("\n📊 月费查询系统")
    print("输入：月费编号 / 姓名 / 英文名")
    print("例如：0160 / 160 / CHE-3 / Anna / 文")
    print("输入 q 退出")

    while True:
        key = input("\n请输入：").strip()

        if key.lower() == "q":
            break

        results = query_member(members, key)

        # ❌ 找不到
        if not results:
            print("❌ 找不到该佛友")
            continue

        # ✅ 只有一个
        if len(results) == 1:
            r = results[0]

            print("\n====== 查询结果 ======")
            print(f"姓名：{r['name']}")
            print(f"编号：{r['member_id']}")

            if r["months"]:
                print(f"已供养月份：{', '.join(r['months'])}")
                print(f"已供养至：{r['latest']}")
            else:
                print("尚未供养")

            print("=====================")

        # 🔎 多个结果
        else:
            print("\n🔎 找到多个结果：")

            for i, r in enumerate(results, 1):
                print(f"{i}. {r['name']} ({r['member_id']})")

            idx = input("请选择编号（输入数字，Enter取消）：").strip()

            if not idx:
                continue

            if idx.isdigit():
                idx = int(idx) - 1

                if 0 <= idx < len(results):
                    r = results[idx]

                    print("\n====== 查询结果 ======")
                    print(f"姓名：{r['name']}")
                    print(f"编号：{r['member_id']}")

                    if r["months"]:
                        print(f"已供养月份：{', '.join(r['months'])}")
                        print(f"已供养至：{r['latest']}")
                    else:
                        print("尚未供养")

                    print("=====================")
                else:
                    print("❌ 选择无效")
            else:
                print("❌ 请输入数字")


# =========================
# 直接运行（测试用）
# =========================
if __name__ == "__main__":
    run_cli()