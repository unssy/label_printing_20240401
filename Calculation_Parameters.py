import pandas as pd
from datetime import datetime


# 定義一個函數來計算日期差距
def calculate_month_diff_dataframe(df: pd.DataFrame, date_column: str = 'date_code',
                                   threshold_months: int = 24) -> pd.DataFrame:
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

    df['out_date'] = df[date_column].apply(lambda date_str: (today.year - datetime.strptime(date_str,
                                                                                            "%Y/%m/%d").year) * 12 + today.month - datetime.strptime(
        date_str, "%Y/%m/%d").month)
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
    df['new_quantity'] = df['quantity']
    # 根據 'expired' 的條件更新 'lot' 和 'date_code' 使用字典
    for threshold, label in mapping_dict.items():
        mask = df['expired'] & (df['out_date'] > threshold)
        df.loc[mask, 'new_lot'] = df.loc[mask, 'lot'] + label
        df.loc[mask, 'new_date_code'] = (
                pd.to_datetime(df.loc[mask, 'date_code']) + pd.DateOffset(months=threshold)).dt.strftime("%Y/%m/%d")
        df.loc[mask, 'new_DC'] = (pd.to_datetime(df.loc[mask, 'DC'].astype(str).str[:2] + '0101',
                                                 format='%y%m%d') + pd.DateOffset(
            years=int(threshold / 12))).dt.strftime("%y") + df.loc[mask, 'DC'].str[-2:]

    return df


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
        medium_label_quantity = [quantity - 60000 * int(quantity / 60000), quantity, 60000]
        medium_label_copies = [0 if medium_remainder == 0 else 1, 1, max(quantity // 60000, 0)]
        # 大包裝
        large_remainder = quantity - 240000 * int(quantity / 240000)
        large_label_quantity = [quantity - 240000 * int(quantity / 240000), quantity, 240000]
        large_label_copies = [0 if large_remainder == 0 else 1, 2, max(quantity // 240000, 0)]

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
        medium_label_quantity = [medium_remainder, quantity, 25200]
        medium_label_copies = [0 if medium_remainder == 0 else 1, 1, max(quantity // 25200, 0)]
        # 大包裝
        large_remainder = quantity - 100800 * int(quantity / 100800)
        large_label_quantity = [large_remainder, quantity, 100800]
        large_label_copies = [0 if large_remainder == 0 else 1, 2, max(quantity // 100800, 0)]

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
    desired_order = ['part_number', 'product_number', 'customer_part_number', 'lot', 'DC', 'date_code', 'quantity',
                     'marking_code',
                     'MPQ', 'package', 'new_lot', 'new_DC', 'new_date_code', 'new_quantity',
                     'small_label_copies', 'small_label_quantity', 'medium_label_copies', 'medium_label_quantity',
                     'large_label_copies', 'large_label_quantity', 'out_date', 'expired']

    return df[desired_order]
