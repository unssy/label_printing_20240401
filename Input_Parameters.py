import pandas as pd
import openpyxl  # 如果需要更精細的Excel操作

worksheet_path = 'path_to_workbook.xlsx'
stock_workbook_path = 'path_to_stock_workbook.xlsx'


def clear_worksheet(sheet_name):
    # 讀取工作表
    workbook = pd.read_excel(worksheet_path, sheet_name=sheet_name)

    # 清除工作表
    workbook.iloc[2:] = ''
    return workbook

def read_stock_data(stock_workbook_path, stock_sheet_name):
    return pd.read_excel(stock_workbook_path, sheet_name=stock_sheet_name)


def main_query(sheet_name):
    query_input_data = pd.read_excel(worksheet_path, sheet_name=sheet_name)

    for index, row in query_input_data.iterrows():
        # 區域變數
        part_no = row['part_no_column']
        qty = row['qty_column']
        # 其他處理


if __name__ == "__main__":
    clear_worksheet('查詢結果')
    stock_data = read_stock_data(stock_workbook_path, '20230105')
    main_query('查詢輸入')