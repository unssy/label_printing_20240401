import pandas as pd
from openpyxl import load_workbook
import os


def clear_worksheet(workbook_path, sheet_name):
    # Load the workbook
    workbook = load_workbook(workbook_path)

    # Get the specific worksheet
    worksheet = workbook[sheet_name]

    # Determine the last row with values
    last_row_with_value = worksheet.max_row

    # Clear the values from row 2 to the last row with values
    for row in worksheet.iter_rows(min_row=2, max_row=last_row_with_value):
        for cell in row:
            cell.value = None


def read_stock_data(workbook_path, sheet_name):
    workbook = load_workbook(workbook_path, data_only=True)
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Sheet name '{sheet_name}' does not exist in the workbook.")
    worksheet = workbook[sheet_name]

    # Read specific range (e.g., from B3 to J[last_row]) into a list of lists
    data = [[cell.value for cell in row] for row in worksheet['B3':'J' + str(worksheet.max_row)]]

    # Create a DataFrame using the first row as the header
    new_columns = ['supplier_code', 'part_number', 'lot', "date_code", 'DC', 'QTY', 'quantity', 'Y', 'store']
    df = pd.DataFrame(data[1:], columns=new_columns)

    return df


def write_to_excel(result_df, workbook_path, sheet_name):
    # 打開現有的工作簿
    book = load_workbook(workbook_path)
    writer = pd.ExcelWriter(workbook_path, engine='openpyxl')
    writer.book = book

    # 獲取要寫入的起始行
    start_row = writer.book[sheet_name].max_row

    # 將DataFrame寫入Excel
    result_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, header=False, index=False)

    writer.save()


def format_date_code(date_code):
    parsed_date = pd.to_datetime(date_code, errors='coerce')
    if pd.notna(parsed_date):
        return parsed_date.strftime('%Y/%m/%d')
    else:
        return str(date_code)

def process_stock_by_max_and_earliest(part_stock, target_quantity):
    part_stock_sorted = part_stock.sort_values(by=['date_code'], ascending=True)
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
            'store': str(closest_lot['store'].values[0])
        }
        result_list.append(result_row)
        accumulated_quantity += take_quantity
        if accumulated_quantity >= target_quantity:
            break
    return result_list

def main_query(workbook_path, sheet_name):
    df_input_parameters = pd.read_excel(workbook_path, sheet_name=sheet_name, engine='openpyxl')
    all_results = []
    for row in df_input_parameters.itertuples(index=False):
        target_part_number = row.part_number
        target_quantity = row.quantity

        part_stock = stock_data[(stock_data['part_number'] == target_part_number) & (stock_data['quantity'] != 0)]

        # 使用最大和最早 lot 的組合處理
        result_list = process_stock_by_max_and_earliest(part_stock, target_quantity)
        all_results.extend(result_list)

    recommend_df = pd.DataFrame(all_results)
    return recommend_df


if __name__ == "__main__":
    # 獲取當前文件的絕對路徑
    current_file_path = os.path.abspath(__file__)
    # 獲取當前文件的所在目錄
    current_directory = os.path.dirname(current_file_path)
    # 設置parameters_worksheet_path和stock_workbook_path的相對路徑
    parameters_worksheet_path = os.path.join(current_directory, 'parameters_dataframe.xlsx')
    stock_workbook_path = os.path.join(current_directory, '(NEW)2023-2.xlsx')
    # clear_worksheet(workbook_path=parameters_worksheet_path, sheet_name='Input_Parameters')
    stock_data = read_stock_data(workbook_path=stock_workbook_path, sheet_name='20230105')
    result_df = main_query(workbook_path=parameters_worksheet_path, sheet_name='Input_Parameters')
    write_to_excel(result_df, workbook_path=parameters_worksheet_path, sheet_name='Query_Parameters')
