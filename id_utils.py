# id_utils.py

def normalize_member_id(raw_id, default_branch="CHE"):
    if raw_id is None:
        return ""

    # 1️⃣ 转字符串 + 清理
    raw = str(raw_id).strip()

    # 👉 处理 Excel 160.0
    if raw.endswith(".0"):
        raw = raw[:-2]

    raw = raw.upper()

    if not raw:
        return ""

    # 2️⃣ 已经是 CHE-160 / STW-160
    if "-" in raw:
        return raw

    # 3️⃣ 0160 → STW
    if raw.startswith("0") and raw.isdigit():
        return f"STW-{int(raw)}"

    # 4️⃣ 160 → CHE
    if raw.isdigit():
        return f"{default_branch}-{int(raw)}"

    return raw