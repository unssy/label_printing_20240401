import pandas as pd
from datetime import datetime
import math

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
# def process_MPQ_1000(row):
#     quantity = row['quantity']
#
#     # 小包裝
#     small_label_copies = quantity / 1000
#     small_label_quantity = [small_label_copies]
#
#     # 中包裝
#     label_copies_medium = math.ceil(small_label_copies / 20)
#     if label_copies_medium == 1:
#         medium_label_quantity = [quantity]
#         medium_label_copies = [label_copies_medium]
#     else:
#         medium_label_quantity = [20000, quantity - 20000 * int(quantity / 20000)]
#         medium_label_copies = [int(quantity / 20000), 1]
#
#     # 大包裝
#     label_copies_large = math.ceil(medium_label_copies[0] / 4) if medium_label_copies else 0
#     large_label_quantity = [80000, medium_label_quantity[1]] if medium_label_quantity else [0, 0]
#     large_label_copies = [int(label_copies_large), 1] if label_copies_large else [0, 0]
#
#     # 根據package資訊進行例外處理（如果需要）
#
#     row['small_label_quantity'] = small_label_quantity
#     row['small_label_copies'] = small_label_copies
#     row['medium_label_quantity'] = medium_label_quantity
#     row['medium_label_copies'] = medium_label_copies
#     row['large_label_quantity'] = large_label_quantity
#     row['large_label_copies'] = large_label_copies
#
#     return row
#
# def process_MPQ_1800(row):
#     quantity = row['quantity']
#
#     # 小包裝
#     small_label_copies = quantity // 1800
#
#     # 中包裝
#     label_copies_medium = math.ceil(small_label_copies / 14)
#     medium_label_quantity = [25200, quantity - 25200 * int(quantity / 25200)] if label_copies_medium > 1 else [quantity]
#     medium_label_copies = [int(quantity / 25200), 1] if label_copies_medium > 1 else [label_copies_medium]
#
#     # 大包裝
#     label_copies_large = math.ceil(medium_label_copies[0] / 4) if medium_label_copies and medium_label_copies[
#         0] > 0 else 0
#     large_label_quantity = [100800, medium_label_quantity[1]] if len(medium_label_quantity) > 1 else [
#         medium_label_quantity[0], 0]
#     large_label_copies = [int(label_copies_large), 1] if label_copies_large else [0, 0]
#
#     # 根據package資訊進行例外處理（如果需要）
#
#     row['small_label_quantity'] = small_label_copies
#     row['small_label_copies'] = small_label_copies
#     row['medium_label_quantity'] = medium_label_quantity
#     row['medium_label_copies'] = medium_label_copies
#     row['large_label_quantity'] = large_label_quantity
#     row['large_label_copies'] = large_label_copies
#
#     return row
# def process_MPQ_2500(row):
#     quantity = row['quantity']
#     MPQ = row['MPQ']
#     medium_box = 20
#     large_box = 4
#
#     # 小包裝
#     small_label_copies = quantity / MPQ
#     small_label_quantity = MPQ
#
#     # 中包裝
#     label_copies_medium = math.ceil(small_label_copies / medium_box)
#     if label_copies_medium == 1:
#         medium_label_quantity = [quantity]
#         medium_label_copies = [label_copies_medium]
#     else:
#         medium_label_quantity = [medium_box*MPQ, quantity - medium_box*MPQ * (label_copies_medium-1)]
#         medium_label_copies = [(label_copies_medium-1), 1]
#
#     # 大包裝
#     label_copies_large = math.ceil(label_copies_medium / large_box)
#     if label_copies_large == 1:
#         large_label_quantity = [quantity]
#         large_label_copies = [label_copies_large]
#     else:
#         large_label_quantity = [large_box*medium_box * MPQ, quantity - large_box*medium_box * MPQ * (label_copies_large - 1)]
#         large_label_copies = [(label_copies_large - 1), 1]
#
#     # 根據package資訊進行例外處理（如果需要）
#
#     row['small_label_quantity'] = small_label_quantity
#     row['small_label_copies'] = small_label_copies
#     row['medium_label_quantity'] = medium_label_quantity
#     row['medium_label_copies'] = medium_label_copies
#     row['large_label_quantity'] = large_label_quantity
#     row['large_label_copies'] = large_label_copies
#
#     return row
#
#
# def process_MPQ_3000(row):
#     quantity = row['quantity']
#
#     # 小包裝
#     small_label_copies = quantity / 3000
#     small_label_quantity = [small_label_copies]
#
#     # 中包裝
#     label_copies_medium = math.ceil(small_label_copies / 20)
#     if label_copies_medium == 1:
#         medium_label_quantity = [quantity]
#         medium_label_copies = [label_copies_medium]
#     else:
#         medium_label_quantity = [60000, quantity - 60000 * int(quantity / 60000)]
#         medium_label_copies = [int(quantity / 60000), 1]
#
#     # 大包裝
#     label_copies_large = math.ceil(medium_label_copies[0] / 4) if medium_label_copies else 0
#     large_label_quantity = [240000, medium_label_quantity[1]] if medium_label_quantity else [0, 0]
#     large_label_copies = [int(label_copies_large), 1] if label_copies_large else [0, 0]
#
#     # 根據package資訊進行例外處理（如果需要）
#
#     row['small_label_quantity'] = small_label_quantity
#     row['small_label_copies'] = small_label_copies
#     row['medium_label_quantity'] = medium_label_quantity
#     row['medium_label_copies'] = medium_label_copies
#     row['large_label_quantity'] = large_label_quantity
#     row['large_label_copies'] = large_label_copies
#
#     return row
#
#
# def process_MPQ_5000(row):
#     pass
#
# def process_MPQ_7500(row):
#     pass
#
# def process_MPQ_8000(row):
#     quantity = row['quantity']
#
#     # 小包裝
#     small_label_copies = quantity / 8000
#     small_label_quantity = [small_label_copies]
#
#     # 中包裝
#     label_copies_medium = math.ceil(small_label_copies / 20)
#     if label_copies_medium == 1:
#         medium_label_quantity = [quantity]
#         medium_label_copies = [label_copies_medium]
#     else:
#         medium_label_quantity = [160000, quantity - 160000 * int(quantity / 160000)]
#         medium_label_copies = [int(quantity / 160000), 1]
#
#     # 大包裝
#     label_copies_large = math.ceil(medium_label_copies[0] / 4) if medium_label_copies else 0
#     large_label_quantity = [640000, medium_label_quantity[1]] if medium_label_quantity else [0, 0]
#     large_label_copies = [int(label_copies_large), 1] if label_copies_large else [0, 0]
#
#     # 根據package資訊進行例外處理（如果需要）
#
#     row['small_label_quantity'] = small_label_quantity
#     row['small_label_copies'] = small_label_copies
#     row['medium_label_quantity'] = medium_label_quantity
#     row['medium_label_copies'] = medium_label_copies
#     row['large_label_quantity'] = large_label_quantity
#     row['large_label_copies'] = large_label_copies
#
#     return row
# def process_MPQ_10000(row):
#     quantity = row['quantity']
#
#     # 小包裝
#     small_label_copies = quantity / 10000
#     small_label_quantity = [small_label_copies]
#
#     # 中包裝
#     label_copies_medium = math.ceil(small_label_copies / 20)
#     if label_copies_medium == 1:
#         medium_label_quantity = [quantity]
#         medium_label_copies = [label_copies_medium]
#     else:
#         medium_label_quantity = [200000, quantity - 200000 * int(quantity / 200000)]
#         medium_label_copies = [int(quantity / 200000), 1]
#
#     # 大包裝
#     label_copies_large = math.ceil(medium_label_copies[0] / 4) if medium_label_copies else 0
#     large_label_quantity = [800000, medium_label_quantity[1]] if medium_label_quantity else [0, 0]
#     large_label_copies = [int(label_copies_large), 1] if label_copies_large else [0, 0]
#
#     # 根據package資訊進行例外處理（如果需要）
#
#     row['small_label_quantity'] = small_label_quantity
#     row['small_label_copies'] = small_label_copies
#     row['medium_label_quantity'] = medium_label_quantity
#     row['medium_label_copies'] = medium_label_copies
#     row['large_label_quantity'] = large_label_quantity
#     row['large_label_copies'] = large_label_copies
#
#     return row
#
#
# handlers = {
#     1000: process_MPQ_1000,
#     1800: process_MPQ_1800,
#     2500: process_MPQ_2500,
#     3000: process_MPQ_3000,
#     5000: process_MPQ_5000,
#     7500: process_MPQ_7500,
#     8000: process_MPQ_8000,
#     10000: process_MPQ_10000
# }
#
# def calculate_label_copies_dataframe(df: pd.DataFrame) -> pd.DataFrame:
#     def calculate_labels(row):
#         handler = handlers.get(row['MPQ'])
#         if handler:
#             return handler(row)
#         return row
#
#     return df.apply(calculate_labels, axis=1)


def process_MPQ(row, medium_box, large_box):
    quantity = int(row['quantity'])
    MPQ = int(row['MPQ'])

    # 小包裝
    small_label_copies = quantity / MPQ
    small_label_quantity = MPQ

    # 中包裝
    label_copies_medium = math.ceil(small_label_copies / medium_box)
    if label_copies_medium == 1:
        medium_label_quantity = [quantity]
        medium_label_copies = [label_copies_medium]
    else:
        medium_label_quantity = [medium_box*MPQ, quantity - medium_box*MPQ * (label_copies_medium-1), quantity]
        medium_label_copies = [(label_copies_medium-1), 1, 1]

    # 大包裝
    label_copies_large = math.ceil(label_copies_medium / large_box)
    if label_copies_large == 1:
        large_label_quantity = [quantity]
        large_label_copies = [label_copies_large*3]
    else:
        large_label_quantity = [large_box*medium_box * MPQ, quantity - large_box*medium_box * MPQ * (label_copies_large - 1)]
        large_label_copies = [(label_copies_large - 1)*3, 3]

    # 根據package資訊進行例外處理（如果需要）

    row['small_label_quantity'] = [int(small_label_quantity)]
    row['small_label_copies'] = [int(small_label_copies)]
    row['medium_label_quantity'] = medium_label_quantity
    row['medium_label_copies'] = medium_label_copies
    row['large_label_quantity'] = large_label_quantity
    row['large_label_copies'] = large_label_copies

    return row

def calculate_label_copies_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    def calculate_labels(row):
        # 如果MPQ為0，則跳過處理並返回原始行
        if row['MPQ'] == 0:
            return row

        # 根據MPQ值設置medium_box和large_box
        if row['MPQ'] in [1000, 2500, 3000, 8000, 10000]:
            medium_box = 20
            large_box = 4
        elif row['MPQ'] == 1800:
            medium_box = 14
            large_box = 4
        elif row['MPQ'] in [5000, 7500]:
            medium_box = 2
            large_box = 8
        else:
            # 可選：處理其他情況或引發錯誤
            raise ValueError("未知的MPQ值：" + str(row['MPQ']))

        # 調用通用處理函數並返回結果
        return process_MPQ(row, medium_box, large_box)

    # 使用DataFrame的apply方法，並將axis設置為1以便按行應用函數
    return df.apply(calculate_labels, axis=1)