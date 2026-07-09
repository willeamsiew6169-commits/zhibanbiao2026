# quote_service.py

import os
import pandas as pd

from datetime import datetime, time, timedelta
from zoneinfo import ZoneInfo

BASE_DIR = os.path.dirname(
    os.path.dirname(
        os.path.dirname(__file__)
    )
)


def get_daily_dharma():
    filepath = os.path.join(
        BASE_DIR,
        "schedule",
        "data",
        "daily_quotes.xlsx"
    )

    print("BASE_DIR =", BASE_DIR)
    print("filepath =", filepath)

    if not os.path.exists(filepath):
        return ""

    try:
        df = pd.read_excel(filepath)

        if df.empty or "content" not in df.columns:
            return ""

        if "active" in df.columns:
            df = df[df["active"] == True]

        df = df.dropna(subset=["content"]).reset_index(drop=True)

        if df.empty:
            return ""

        malaysia_now = datetime.now(ZoneInfo("Asia/Kuala_Lumpur"))

        quote_date = malaysia_now.date()
        if malaysia_now.time() < time(18, 0):
            quote_date = quote_date - timedelta(days=1)

        base_date = datetime(2026, 7, 1).date()
        day_index = (quote_date - base_date).days
        index = day_index % len(df)

        print("index =", index)
        print("content =", df.iloc[index]["content"])

        return str(df.iloc[index]["content"]).strip()

    except Exception as e:
        print("读取每日法语失败：", e)
        return ""