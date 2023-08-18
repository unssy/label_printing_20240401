import pandas as pd
from datetime import datetime

# 定義一個函數來計算日期差距
def calculate_month_diff_dataframe(df: pd.DataFrame, date_column: str = 'date_code') -> pd.DataFrame:
    """
    計算 dataframe 中指定日期欄位與當前日期的月份差距，並將結果存儲在新的 'out_date' 欄位中。

    參數:
        df (pd.DataFrame): 輸入的 dataframe。
        date_column (str): 包含日期資料的欄位名稱。預設為 'date_code'.

    返回:
        pd.DataFrame: 包含新的 'out_date' 欄位的 dataframe。
    """

    def month_diff(date_str):
        date_format = "%Y/%m/%d"
        date_object = datetime.strptime(date_str, date_format)

        today = datetime.today()
        return (today.year - date_object.year) * 12 + today.month - date_object.month

    df['out_date'] = df[date_column].apply(month_diff)
    return df


# 定義每個 MPQ 值的處理邏輯
def process_MPQ_1000(row):
    # ... (MPQ = 1800 的邏輯)
    return row


def process_MPQ_1800(row):
    quantity = row['quantity']

    label_copies_small = quantity / 1800
    label_copies_medium = abs(label_copies_small / 14)


    if label_copies_medium < 1:
        label_copies_medium_quantity = quantity
    else:
        label_copies_medium_quantity = quantity - (25200 * (label_copies_medium - 1))
        label_copies_medium_quantity
        label_copies_medium = 1

    label_copies_large = abs(label_copies_medium / 4)
    if label_copies_large < 1:
        label_copies_large = 3
        quantity = quantity
    else:
        quantity -= 100800 * (label_copies_large - 1)
        label_copies_large = 1

    row['label_copies_small'] = label_copies_small
    row['label_copies_medium'] = label_copies_medium
    row['label_copies_large'] = label_copies_large

    return row


def process_MPQ_2500(row):
    # ... (MPQ = 1800 的邏輯)
    return row


def process_MPQ_3000(row):
    quantity = row['quantity']

    label_copies_small = quantity / 3000
    label_copies_medium = abs(label_copies_small / 20)
    label_copies_large = abs(label_copies_medium / 4)

    row['label_copies_small'] = label_copies_small
    row['label_copies_medium'] = label_copies_medium
    row['label_copies_large'] = label_copies_large

    return row


def process_MPQ_5000(row):
    # ... (MPQ = 1800 的邏輯)
    return row


def process_MPQ_7500(row):
    # ... (MPQ = 3000 的邏輯)
    return row


def process_MPQ_8000(row):
    # ... (MPQ = 3000 的邏輯)
    return row


def process_MPQ_10000(row):
    # ... (MPQ = 3000 的邏輯)
    return row

# 其他 MPQ 處理邏輯也可以在這裡定義

# 建立一個字典，將每個 MPQ 值與其對應的處理函數關聯起來
MPQ_handlers = {
    1000: process_MPQ_1000,
    1800: process_MPQ_1800,
    2500: process_MPQ_2500,
    3000: process_MPQ_3000,
    5000: process_MPQ_5000,
    7500: process_MPQ_7500,
    8000: process_MPQ_8000,
    10000: process_MPQ_10000
}

# 主計算函數
def calculate_label_copies_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    根據提供的業務邏輯計算 'label_copies_small'、'label_copies_medium' 和 'label_copies_large'。

    參數:
        df (pd.DataFrame): 輸入的 dataframe。

    返回:
        pd.DataFrame: 包含新的標籤欄位的 dataframe。
    """

    def calculate_labels(row):
        handler = MPQ_handlers.get(row['MPQ'])  # 根據 MPQ 值獲取對應的處理函數
        if handler:
            return handler(row)
        return row  # 如果沒有對應的處理函數，則返回原始行

    return df.apply(calculate_labels, axis=1)
