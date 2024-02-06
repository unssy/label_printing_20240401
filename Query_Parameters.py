from  utinities import *
import pandas as pd
import win32com.client as win32
from datetime import datetime

def get_sampling_count(dataframe):

    shipment_bounds = [2, 9, 16, 26, 51, 91, 151, 281, 501, 1201, 3201, 10001, 35001, 150001, 500001]
    sampling_counts = [2, 3, 5, 8, 13, 20, 32, 50, 80, 125, 200, 315, 500, 800, 1250]

    def compute_sampling(shipment_count):
        shipment_bounds = [2, 9, 16, 26, 51, 91, 151, 281, 501, 1201, 3201, 10001, 35001, 150001, 500001]
        sampling_counts = [2, 3, 5, 8, 13, 20, 32, 50, 80, 125, 200, 315, 500, 800, 1250]

        # 嘗試將 shipment_count 轉換為整數，如果轉換失敗則返回 None
        try:
            shipment_count = int(shipment_count)
        except ValueError:
            return None

        # 如果數量超過上限，直接返回1250
        if shipment_count > shipment_bounds[-1]:
            return 1250

        left, right = 0, len(shipment_bounds) - 1
        while left <= right:
            mid = left + (right - left) // 2
            if shipment_bounds[mid] <= shipment_count < shipment_bounds[mid + 1]:
                return sampling_counts[mid]
            elif shipment_count < shipment_bounds[mid]:
                right = mid - 1
            else:
                left = mid + 1
        return None

    dataframe['sampling'] = dataframe['quantity'].apply(compute_sampling)
    return dataframe


def merge_with_reference(dataframe):
    # 读取参照表
    parameters_db_1 = pd.read_csv('parameters_database.csv')
    parameters_db_2 = pd.read_csv('customer_no.csv')
    # 根据'customer_no'查找'customer_name'并加入到merged_dataframe
    merged_dataframe = pd.merge(dataframe, parameters_db_2[['customer_no', 'customer_name']], on='customer_no',
                                how='left')

    # 在合并前将 'part_number' 列转换为小写
    lower_part_number_column = 'lower_part_number'
    try:
        merged_dataframe = pd.merge(merged_dataframe, parameters_db_1.rename(columns={'part_number': lower_part_number_column}),
                                    left_on='part_number', right_on=lower_part_number_column,
                                    how='left', suffixes=('', '_ref'))
    except KeyError:
        # 若 'part_number' 列不存在，填充 NaN 并继续进行合并
        merged_dataframe = pd.merge(dataframe, parameters_db_1, on='part_number', how='left')

    # 填充 NaN 并将 'MPQ' 列转换为整数
    merged_dataframe['MPQ'].fillna(0, inplace=True)
    merged_dataframe['MPQ'] = merged_dataframe['MPQ'].astype(int)

    return merged_dataframe

def process_stock_by_max_and_earliest(part_stock, target_quantity):
    part_stock = part_stock.copy()
    part_stock['date_code'] = part_stock['date_code'].apply(format_date_code)
    part_stock_sorted = part_stock.sort_values(by=['date_code'], ascending=True, na_position='last')
    result_list = []
    accumulated_quantity = 0

    for date_code, group in part_stock_sorted.groupby('date_code'):
        exact_match = group[group['quantity'] == target_quantity]
        if not exact_match.empty:
            result_list.append(exact_match.iloc[0].to_dict())
            accumulated_quantity += target_quantity
            break

        # 找到與 target_quantity 最接近的 row
        closest_lot = group.iloc[(group['quantity'] - target_quantity).abs().argsort()[:1]]
        take_quantity = min(closest_lot['quantity'].values[0], target_quantity - accumulated_quantity)
        formatted_date = format_date_code(date_code)
        result_row = {
            'part_number': str(closest_lot['part_number'].values[0]),
            'lot': str(closest_lot['lot'].values[0]),
            'DC': str(closest_lot['DC'].values[0]),
            'date_code': formatted_date,
            'quantity': str(take_quantity),
            'store': str(closest_lot['store'].values[0]),
            'row_index': str(closest_lot['row_index'].values[0])
        }
        result_list.append(result_row)
        accumulated_quantity += take_quantity
        if accumulated_quantity >= target_quantity:
            break
    return result_list


def main_query(input_dataframe, stock_dataframe):
    all_results = []
    for row in input_dataframe.itertuples(index=False):
        target_part_number = row.part_number
        target_quantity = row.quantity

        part_stock = stock_dataframe[
            (stock_dataframe['part_number'].str.upper() == target_part_number.upper()) &
            (stock_dataframe['quantity'] != 0)]
        # 使用最大和最早 lot 的組合處理
        result_list = process_stock_by_max_and_earliest(part_stock, target_quantity)
        all_results.extend(result_list)

    recommend_df = pd.DataFrame(all_results)
    # 使用 'part_number' 列合併 recommend_df 和 input_dataframe 中的 'product_number' 列
    recommend_df = recommend_df.merge(input_dataframe[['part_number', 'product_number']], on='part_number', how='left')
    input_dataframe['delivery_date'] = input_dataframe['delivery_date'].map(format_date_code)
    recommend_df['delivery_date'] = input_dataframe['delivery_date'].iloc[0]
    recommend_df['customer_no'] = input_dataframe['customer_no'].iloc[0]
    recommend_df['deduct'] = False
    recommend_df = get_sampling_count(recommend_df)
    recommend_df = merge_with_reference(recommend_df)
    desired_order = ['customer_no', 'customer_name', 'part_number', 'product_number','lot', 'DC', 'date_code', 'quantity',
                     'store', 'row_index', 'remark', 'sampling', 'marking_code', 'package','MPQ','delivery_date','customer_no']
    recommend_df = recommend_df[desired_order]
    return recommend_df


def convert_to_mm_dd(date_str):
    """Convert date string from yyyymmdd to mm/dd format."""
    date_obj = datetime.strptime(date_str, "%Y/%m/%d")
    # 直接获取月份和日期，不包括前导零
    month = str(date_obj.month)
    day = str(date_obj.day)
    return f"{month}/{day}"


def find_insert_column(ws, target_date):
    """Find the column index before which the new date column should be inserted."""
    for col in range(ws.UsedRange.Columns.Count, 0, -1):
        cell_value = ws.Cells(1, col).Value
        if cell_value and cell_value < target_date:
            return col + 1
    return None


def open_excel_workbook(file_path, password):
    """Open an Excel workbook with given write password."""

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    try:
        # win32.gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, 9)
        wb = excel.Workbooks.Open(
            file_path,
            UpdateLinks=0,
            ReadOnly=False,
            WriteResPassword=password,
            IgnoreReadOnlyRecommended=True
        )
        if wb.ReadOnly:
            raise Exception("檔案已被開啟")
        return excel, wb
    except Exception as e:
        print(f"Error: {e}")
        excel.Application.Quit()
        exit()


def autofill_formula(ws, source_col, row, target_col):
    """Autofill formula in Excel from source column to target column for the given row."""
    source_cell = ws.Cells(row, source_col)
    dest_range = ws.Range(ws.Cells(row, source_col), ws.Cells(row, target_col))
    source_cell.AutoFill(Destination=dest_range, Type=win32.constants.xlFillDefault)

def deduct_stock(file_path, sheet, data_df, password):
    """Deduct stock based on provided dataframe."""
    # excel = None
    # wb = None  # Initialize wb to None before the try block
    try:
        excel, wb = open_excel_workbook(file_path, password)
        ws = wb.Worksheets[sheet]
        delivery_date = convert_to_mm_dd(str(data_df['delivery_date'].iloc[0]))
        col_to_insert = find_insert_column(ws, delivery_date)
        if col_to_insert is None:
            print("找不到特定日期之前的最後一列。")
            return

        ws.Columns(col_to_insert).Insert()
        ws.Cells(1, col_to_insert).Value = delivery_date
        ws.Cells(2, col_to_insert).Value = data_df['customer_name'].iloc[0]
        for _, row in data_df.iterrows():
            ws.Cells(row['row_index'], col_to_insert).Value = row['quantity']

        autofill_formula(ws, col_to_insert - 1, 3, col_to_insert)

        print("出貨總數為: ", int(ws.Cells(3, col_to_insert).Value))
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        if excel is not None:
            excel.Application.Quit()