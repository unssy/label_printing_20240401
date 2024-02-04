import pandas as pd
from datetime import datetime
import math

# 定義一個函數來計算日期差距
def calculate_month_diff_dataframe(df: pd.DataFrame, date_column: str = 'date_code', threshold_months: int = 24) -> pd.DataFrame:
    """
    計算 dataframe 中指定日期欄位與當前日期的月份差距，並將結果存儲在新的 'out_date' 欄位中。
    同時，新增 'expired' 欄位標記超過指定月份的項目。

    參數:
        df (pd.DataFrame): 輸入的 dataframe。
        date_column (str): 包含日期資料的欄位名稱。預設為 'date_code'.
        threshold_months (int): 超過指定月份被標記為 'expired'。預設為 24.

    返回:
        pd.DataFrame: 包含新的 'out_date' 和 'expired' 欄位的 dataframe。
    """
    today = datetime.today()

    df['out_date'] = df[date_column].apply(lambda date_str: (today.year - datetime.strptime(date_str, "%Y/%m/%d").year) * 12 + today.month - datetime.strptime(date_str, "%Y/%m/%d").month)
    df['expired'] = df['out_date'] > threshold_months
    # 對 'marking_code' 包含 'YM'、'ym'、'YW'、'yw' 的行，將 'expired' 設為 False
    df.loc[df['marking_code'].str.upper().str.contains('|'.join(['YM', 'YW']), na=False), 'expired'] = False

    return df

def update_columns_based_on_expired(df: pd.DataFrame):
    """
    根據 'expired' 欄位的值更新 'lot' 和 'date_code' 欄位。

    參數:
        df (pd.DataFrame): 輸入的 dataframe。

    返回:
        pd.DataFrame: 更新後的 dataframe。
    """
    mapping_dict = {12: 'A', 24: 'B', 36: 'C', 48: 'D', 60: 'E', 72: 'F', 84: 'G', 96: 'H', 108: 'I'}
    df['new_lot'] = df['lot']
    df['new_DC'] = df['DC']
    df['new_date_code'] = df['date_code']
    # 根據 'expired' 的條件更新 'lot' 和 'date_code' 使用字典
    for threshold, label in mapping_dict.items():
        mask = df['expired'] & (df['out_date'] > threshold)
        df.loc[mask, 'new_lot'] = df.loc[mask, 'lot'] + label
        df.loc[mask, 'new_date_code'] = (
                    pd.to_datetime(df.loc[mask, 'date_code']) + pd.DateOffset(months=threshold)).dt.strftime("%Y/%m/%d")
        df.loc[mask, 'new_DC'] = (pd.to_datetime(df.loc[mask, 'DC'].astype(str).str[:2] + '0101', format='%y%m%d') + pd.DateOffset(years=int(threshold/12))).dt.strftime("%y") + df.loc[mask, 'DC'].str[-2:]

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


# def process_MPQ(row, medium_box, large_box):
#     quantity = int(row['quantity'])
#     MPQ = int(row['MPQ'])
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
#         medium_label_quantity = [medium_box*MPQ, quantity - medium_box*MPQ * (label_copies_medium-1), quantity]
#         medium_label_copies = [(label_copies_medium-1), 1, 1]
#
#     # 大包裝
#     label_copies_large = math.ceil(label_copies_medium / large_box)
#     if label_copies_large == 1:
#         large_label_quantity = [quantity]
#         large_label_copies = [label_copies_large*3]
#     else:
#         large_label_quantity = [large_box*medium_box * MPQ, quantity - large_box*medium_box * MPQ * (label_copies_large - 1)]
#         large_label_copies = [(label_copies_large - 1)*3, 3]
#
#     # 根據package資訊進行例外處理（如果需要）
#
#     row['small_label_quantity'] = [int(small_label_quantity)]
#     row['small_label_copies'] = [int(small_label_copies)]
#     row['medium_label_quantity'] = medium_label_quantity
#     row['medium_label_copies'] = medium_label_copies
#     row['large_label_quantity'] = large_label_quantity
#     row['large_label_copies'] = large_label_copies
#
#     return row
#
# def calculate_label_copies_dataframe(df: pd.DataFrame) -> pd.DataFrame:
#     def calculate_labels(row):
#         # 如果MPQ為0，則跳過處理並返回原始行
#         if row['MPQ'] == 0:
#             return row
#
#         # 根據MPQ值設置medium_box和large_box
#         if row['MPQ'] in [1000, 2500, 3000, 8000, 10000]:
#             medium_box = 20
#             large_box = 4
#         elif row['MPQ'] == 1800:
#             medium_box = 14
#             large_box = 4
#         elif row['MPQ'] in [5000, 7500]:
#             medium_box = 2
#             large_box = 8
#         else:
#             # 可選：處理其他情況或引發錯誤
#             raise ValueError("未知的MPQ值：" + str(row['MPQ']))
#
#         # 調用通用處理函數並返回結果
#         return process_MPQ(row, medium_box, large_box)
#
#     # 使用DataFrame的apply方法，並將axis設置為1以便按行應用函數
#     return df.apply(calculate_labels, axis=1)

def calculate_label_copies_dataframe(df):
    mpq_categories = {
        '7_inch_reel': [250, 500, 800, 2000, 2500, 3000, 4000, 5000, 8000, 10000, 12000],
        '7_inch_thick_reel': [1000, 1800],
        '13_inch_reel': [2500, 5000, 7500],
        'TO252': [2500]  # TO252 特殊情況使用 13_inch_reel 的 MPQ
    }

    def process_row(row):
        MPQ = row['MPQ']
        package = row['package']
        df['quantity'] = df['quantity'].astype(int)
        category = find_category(MPQ, package, mpq_categories)
        if category:
            if category == '7_inch_reel':
                process_7_inch_reel(row)
            elif category == '7_inch_thick_reel':
                process_7_inch_thick_reel(row)
            elif category == '13_inch_reel':
                process_13_inch_reel(row)

        return row

    def find_category(MPQ, package, categories):
        if package == 'TO252':
            return '13_inch_reel' if MPQ in categories['TO252'] else None
        else:
            for category, mpq_values in categories.items():
                if MPQ in mpq_values:
                    return category
            return None

    # 定義處理不同分類的子函數
    # 子函數：7 inch reel 的處理邏輯
    def process_7_inch_reel(row):
        quantity = row['quantity']
        MPQ = row['MPQ']
        # 小包裝
        small_label_copies = quantity // MPQ
        # 中包裝
        medium_remainder = quantity - 60000 * int(quantity / 60000)
        medium_label_quantity =[quantity - 60000 * int(quantity / 60000),quantity,60000]
        medium_label_copies =[0 if medium_remainder == 0 else 1,1,max(quantity // 60000, 0)]
        # 大包裝
        large_remainder = quantity - 240000 * int(quantity / 240000)
        large_label_quantity = [quantity - 240000 * int(quantity / 240000),quantity,240000]
        large_label_copies = [0 if large_remainder == 0 else 1,2,max(quantity // 240000, 0)]

        row['small_label_quantity'] = MPQ
        row['small_label_copies'] = small_label_copies
        row['medium_label_quantity'] = medium_label_quantity
        row['medium_label_copies'] = medium_label_copies
        row['large_label_quantity'] = large_label_quantity
        row['large_label_copies'] = large_label_copies

        return row

    def process_7_inch_thick_reel(row):
        quantity = row['quantity']
        MPQ = row['MPQ']
        # 小包裝
        small_label_copies = quantity // MPQ
        # 中包裝
        medium_remainder = quantity - 25200 * int(quantity / 25200)
        medium_label_quantity =[medium_remainder,quantity,25200]
        medium_label_copies =[0 if medium_remainder == 0 else 1,1,max(quantity // 25200, 0)]
        # 大包裝
        large_remainder = quantity - 100800 * int(quantity / 100800)
        large_label_quantity = [large_remainder,quantity,100800]
        large_label_copies = [0 if large_remainder == 0 else 1,2,max(quantity // 100800, 0)]

        row['small_label_quantity'] = MPQ
        row['small_label_copies'] = small_label_copies
        row['medium_label_quantity'] = medium_label_quantity
        row['medium_label_copies'] = medium_label_copies
        row['large_label_quantity'] = large_label_quantity
        row['large_label_copies'] = large_label_copies

        return row

    # 子函數：13 inch reel 的處理邏輯
    def process_13_inch_reel(row):
        quantity = row['quantity']
        MPQ = row['MPQ']
        medium_unit = 15000
        large_unit = 120000
        package = row['package']
        if package == 'TO252':
            medium_unit = 50000
            large_unit = 200000
        # 小包裝
        small_label_copies = quantity // MPQ
        # 中包裝
        medium_remainder = quantity - medium_unit * int(quantity / medium_unit)
        medium_label_quantity = [medium_remainder, quantity, medium_unit]
        medium_label_copies = [0 if medium_remainder == 0 else 1, 1, max(quantity // medium_unit, 0)]
        # 大包裝
        large_remainder = quantity - large_unit * int(quantity / large_unit)
        large_label_quantity = [large_remainder, quantity, large_unit]
        large_label_copies = [0 if large_remainder == 0 else 1, 2, max(quantity // large_unit, 0)]

        row['small_label_quantity'] = MPQ
        row['small_label_copies'] = small_label_copies
        row['medium_label_quantity'] = medium_label_quantity
        row['medium_label_copies'] = medium_label_copies
        row['large_label_quantity'] = large_label_quantity
        row['large_label_copies'] = large_label_copies

        return row

    # 套用主函數到 DataFrame
    df = df.apply(process_row, axis=1)
    desired_order = ['part_number', 'product_number', 'customer_part_number', 'lot', 'DC', 'date_code', 'marking_code',
                     'MPQ', 'package', 'new_lot', 'new_DC', 'new_date_code',
                    'small_label_copies', 'small_label_quantity', 'medium_label_copies', 'medium_label_quantity', 'large_label_copies', 'large_label_quantity', 'out_date', 'expired']

    return df[desired_order]