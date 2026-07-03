# quote_service.py

import os
import pandas as pd

BASE_DIR = os.path.dirname(
    os.path.dirname(
        os.path.dirname(__file__)
    )
)


def get_daily_dharma():
    """
    读取每日法语（第一版）
    目前固定读取第一句。
    """

    filepath = os.path.join(
        BASE_DIR,
        "data",
        "daily_quotes.xlsx"
    )

    if not os.path.exists(filepath):
        return ""

    try:

        df = pd.read_excel(filepath)

        if "active" in df.columns:
            df = df[df["active"] == True]

        if df.empty:
            return ""

        return str(df.iloc[0]["content"]).strip()

    except Exception as e:

        print("读取每日法语失败：", e)

        return ""