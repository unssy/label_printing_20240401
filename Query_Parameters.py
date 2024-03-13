from  utilities import *
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
    parameters_db_1 = parameters_df
    parameters_db_2 = customer_no_df
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

def process_stock_by_max_and_earliest(part_stock, target_quantity, target_purchase_order, stock_dataframe):
    part_stock = part_stock.copy()
    part_stock['date_code'] = part_stock['date_code'].apply(format_date_code)
    part_stock_sorted = part_stock.sort_values(by=['date_code', 'row_index'], ascending=[True, True], na_position='last')
    result_rows = []
    accumulated_quantity = 0
    found_enough_stock = False  # 标志是否找到足够的库存

    # 额外的检查，看是否有完全匹配的库存
    exact_match = part_stock_sorted[part_stock_sorted['quantity'] == target_quantity]
    if not exact_match.empty:
        for index, row in exact_match.iterrows():
            row['purchase_order'] = str(target_purchase_order)
            result_rows.append(row)
            accumulated_quantity += target_quantity
            found_enough_stock = True  # 找到足够的库存
            stock_dataframe.loc[stock_dataframe['row_index'] == row['row_index'], 'quantity'] -= row['quantity']
        return result_rows
    # 没有完全匹配的情况下，继续循环查找部分匹配
    for index, row in part_stock_sorted.iterrows():
        remaining_quantity = target_quantity - accumulated_quantity

        if row['quantity'] >= remaining_quantity:
            result_row = row.copy()
            result_row['quantity'] = remaining_quantity
            result_row['purchase_order'] = str(target_purchase_order)
            result_rows.append(result_row)
            accumulated_quantity += remaining_quantity
            found_enough_stock = True  # 找到足够的库存
            break

        row['purchase_order'] = str(target_purchase_order)
        result_rows.append(row)
        accumulated_quantity += row['quantity']
        stock_dataframe.loc[stock_dataframe['row_index'] == row['row_index'], 'quantity'] -= row['quantity']

    if not found_enough_stock:
        # 在这里添加缺少库存的备注信息
        remark_row = {
            'part_number': str(row['part_number']),
            'lot': 'N/A',
            'DC': 'N/A',
            'date_code': 'N/A',
            'quantity': 0,
            'store': 'N/A',
            'row_index': 'N/A',
            'purchase_order': str(target_purchase_order),
            'remark': '缺少库存' + str(target_quantity - accumulated_quantity)
        }
        result_rows.append(remark_row)

    return result_rows


def main_query(preprocess_input_dataframe, stock_dataframe, parameters_database_path, customer_no_database_path):
    all_results = []
    for row in preprocess_input_dataframe.itertuples(index=False):
        target_part_number = row.part_number
        target_quantity = row.quantity
        target_purchase_order = row.purchase_order

        part_stock = stock_dataframe[
            (stock_dataframe['part_number'].str.upper() == target_part_number.upper()) &
            (stock_dataframe['quantity'] != 0)]
        # 使用最大和最早 lot 的組合處理
        result_list = process_stock_by_max_and_earliest(part_stock, target_quantity, target_purchase_order, stock_dataframe)
        all_results.extend(result_list)

    recommend_df = pd.DataFrame(all_results)
    global parameters_df, customer_no_df
    parameters_df = pd.read_csv(parameters_database_path)
    customer_no_df = pd.read_csv(customer_no_database_path)
    recommend_df = recommend_df.merge(preprocess_input_dataframe, on=['part_number', 'purchase_order'], how='left', suffixes=('', '_x'))
    recommend_df['deduct'] = False
    recommend_df = get_sampling_count(recommend_df)
    recommend_df = merge_with_reference(recommend_df)
    desired_order = ['customer_no', 'customer_name', 'part_number', 'product_number', 'customer_part_number',
                     'lot', 'DC', 'date_code', 'quantity', 'purchase_order', 'unit_price',
                     'store', 'row_index', 'remark', 'sampling', 'marking_code', 'package', 'MPQ', 'delivery_date',
                     'invoice_series', 'currency', 'deduct', 'month', 'day']
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


def print_excel_sheet(workbook_path, sheet_name):
    excel_app = None
    try:
        excel_app = win32.gencache.EnsureDispatch('Excel.Application')
        excel_app.Visible = False  # Set to True if you want Excel to be visible during printing

        workbook = excel_app.Workbooks.Open(workbook_path)
        worksheet = workbook.Sheets[sheet_name]

        # # Set the printer
        printer_name = "KONICA MINOLTA 367SeriesPCL on Ne10:"
        worksheet.PageSetup.Orientation =  2
        worksheet.PageSetup.FitToPagesWide = 1


        worksheet.PrintOut(Copies=1, ActivePrinter= printer_name)

        workbook.Close(SaveChanges=False)
        excel_app.Quit()

        print(f"Successfully printed '{sheet_name}' from '{workbook_path}'")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        if excel_app:
            workbook.Close(SaveChanges=False)
            excel_app.Quit()