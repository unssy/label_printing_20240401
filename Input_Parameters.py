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


def main_query(workbook_path, sheet_name):
    df_input_parameters = pd.read_excel(workbook_path, sheet_name=sheet_name, engine='openpyxl')

    result_list = []
    for row in df_input_parameters.itertuples(index=False):
        target_part_number = row.part_number
        target_quantity = row.quantity

        part_stock = stock_data[(stock_data['part_number'] == target_part_number) & (stock_data['quantity'] != 0)]

        # 按lot分組
        grouped_by_lot = part_stock.groupby('lot', group_keys=False)

        accumulated_quantity = 0

        # 遍歷每個lot
        for lot, group in grouped_by_lot:
            # 在每個lot內部按照date_code排序
            group_sorted_by_date = group.sort_values(by='date_code')

            # 逐個選擇庫存
            for index, stock_row in group_sorted_by_date.iterrows():
                take_quantity = min(stock_row['quantity'], target_quantity - accumulated_quantity)
                formatted_date = format_date_code(stock_row['date_code'])
                result_row = {
                    'part_number': str(stock_row['part_number']),
                    'lot': str(stock_row['lot']),
                    'DC': str(stock_row['DC']),
                    'date_code': formatted_date,
                    'quantity': str(take_quantity),  # 直接設置為take_quantity
                    'store': str(stock_row['store'])
                }
                result_list.append(result_row)
                accumulated_quantity += take_quantity
                if accumulated_quantity >= target_quantity:
                    break

            if accumulated_quantity >= target_quantity:
                break
    recommend_df = pd.DataFrame(result_list)
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
