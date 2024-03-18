from utilities import *
import pandas as pd
import win32com.client as win32
import win32con
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


def parameters_join_ignore_case(df_left, df_right, columns=None):
    """
    Perform a left join between two DataFrames while ignoring the case of 'part_number' fields.

    Args:
    - df_left (DataFrame): The left DataFrame.
    - df_right (DataFrame): The right DataFrame.
    - columns (list of str, optional): Columns to select from the right DataFrame.

    Returns:
    - DataFrame: The result of the left join operation.
    """
    # Create a new column with uppercase 'part_number' for right DataFrame
    df_right['part_number_upper'] = df_right['part_number'].str.upper()

    # Perform the left join
    merged_df = pd.merge(df_left, df_right[['part_number_upper'] + (columns or [])], on='part_number_upper', how='left')

    # Fill NaN values in 'MPQ' column with 0 and convert it to integer
    merged_df['MPQ'].fillna(0, inplace=True)
    merged_df['MPQ'] = merged_df['MPQ'].astype(int)

    # Drop the temporary uppercase column
    merged_df.drop(columns=['part_number_upper'], inplace=True)

    return merged_df


def left_join_ignore_case(df1, df2):
    """
    Perform a left join between two DataFrames while ignoring the case of 'part_number' fields.

    Args:
    - df1 (DataFrame): The first DataFrame.
    - df2 (DataFrame): The second DataFrame.

    Returns:
    - DataFrame: The result of the left join operation.
    """
    # Convert 'part_number' fields to uppercase in both DataFrames
    df1['part_number_upper'] = df1['part_number'].str.upper()
    df2['part_number_upper'] = df2['part_number'].str.upper()

    # Perform left join
    merged_df = df1.merge(df2,
                          left_on=['part_number_upper', 'purchase_order'],
                          right_on=['part_number_upper', 'purchase_order'],
                          how='left',
                          suffixes=('', '_x'))

    # Remove the additional fields with suffix '_x'
    columns_to_drop = [col for col in merged_df.columns if col.endswith('_x')]
    merged_df.drop(columns=columns_to_drop, inplace=True)

    return merged_df


def process_stock_by_max_and_earliest(part_stock, target_quantity, target_purchase_order, stock_dataframe):
    # Make a copy of the part_stock DataFrame to avoid modifying the original
    part_stock = part_stock.copy()

    # Format the date_code column if needed
    part_stock['date_code'] = part_stock['date_code'].apply(format_date_code)

    # Sort the DataFrame by date_code and row_index in ascending order
    part_stock_sorted = part_stock.sort_values(by=['date_code', 'row_index'], ascending=[True, True],
                                               na_position='last')

    # Initialize an empty list to store result rows
    result_rows = []

    # Initialize a variable to track the accumulated quantity of target parts
    accumulated_target_quantity = 0

    # Check if there is an exact match for the target quantity in the stock
    exact_match = part_stock_sorted[part_stock_sorted['quantity'] == target_quantity]
    if not exact_match.empty:
        # If an exact match is found, update result_rows and stock_dataframe accordingly
        for index, row in exact_match.iterrows():
            row['purchase_order'] = str(target_purchase_order)
            result_rows.append(row)
            accumulated_target_quantity += target_quantity
            stock_dataframe.loc[stock_dataframe['row_index'] == row['row_index'], 'quantity'] -= row['quantity']
        return result_rows

    # Iterate through the sorted part_stock DataFrame to find partial matches
    for index, row in part_stock_sorted.iterrows():
        remaining_target_quantity = target_quantity - accumulated_target_quantity

        # If the current row has enough quantity to fulfill the remaining target quantity
        if row['quantity'] >= remaining_target_quantity:
            # Update the quantity to match the remaining target quantity
            row['quantity'] = remaining_target_quantity
            found_enough_stock = True
        else:
            # If the current row does not have enough quantity, use all of its quantity
            remaining_target_quantity = row['quantity']
            found_enough_stock = False

        # Assign the purchase_order to the current row
        row['purchase_order'] = str(target_purchase_order)
        result_rows.append(row)
        accumulated_target_quantity += remaining_target_quantity

        # If enough stock is found to fulfill the target quantity, exit the loop
        if found_enough_stock:
            break

    # If enough stock is not found to fulfill the target quantity, add a remark row
    if not found_enough_stock:
        remark_row = {
            'part_number': str(row['part_number']),
            'lot': 'N/A',
            'DC': 'N/A',
            'date_code': 'N/A',
            'quantity': 0,
            'store': 'N/A',
            'row_index': 'N/A',
            'purchase_order': str(target_purchase_order),
            'remark': 'Insufficient stock: ' + str(target_quantity - accumulated_target_quantity)
        }
        result_rows.append(remark_row)

    # Update stock_dataframe with the quantities taken from the stock
    for row in result_rows:
        if row['row_index'] != 'N/A':
            stock_dataframe.loc[stock_dataframe['row_index'] == row['row_index'], 'quantity'] -= row['quantity']

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
        result_list = process_stock_by_max_and_earliest(part_stock, target_quantity, target_purchase_order,
                                                        stock_dataframe)
        all_results.extend(result_list)

    recommend_df = pd.DataFrame(all_results)
    global parameters_df, customer_no_df
    parameters_df = pd.read_csv(parameters_database_path)
    customer_no_df = pd.read_csv(customer_no_database_path)
    recommend_df = left_join_ignore_case(recommend_df, preprocess_input_dataframe)
    recommend_df['deduct'] = False
    recommend_df = get_sampling_count(recommend_df)
    recommend_df = pd.merge(recommend_df, customer_no_df[['customer_no', 'customer_name']], on='customer_no',
                            how='left')
    parameters_df['part_number_upper'] = parameters_df['part_number'].str.upper()

    recommend_df = parameters_join_ignore_case(recommend_df, parameters_df,
                                               columns=['marking_code', 'package', 'MPQ'])

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
    excel.Visible = False
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
            current_value = ws.Cells(row['row_index'], col_to_insert).Value
            current_value = 0 if current_value is None else current_value
            ws.Cells(row['row_index'], col_to_insert).Value = int(current_value) + int(row['quantity'])

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
        excel_app.DisplayAlerts = False

        workbook = excel_app.Workbooks.Open(workbook_path)
        worksheet = workbook.Sheets[sheet_name]

        # # Set the printer
        printer_name = "KONICA MINOLTA 367SeriesPCL on Ne10:"
        worksheet.PageSetup.Orientation = win32con.DMORIENT_LANDSCAPE
        worksheet.PageSetup.Zoom = False
        worksheet.PageSetup.FitToPagesWide = 1

        worksheet.PrintOut(Copies=1, ActivePrinter=printer_name)

        workbook.Close(SaveChanges=False)
        excel_app.Quit()

        print(f"Successfully printed '{sheet_name}' from '{workbook_path}'")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        if excel_app:
            workbook.Close(SaveChanges=False)
            excel_app.Quit()
