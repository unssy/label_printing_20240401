import pandas as pd
from openpyxl import load_workbook


def clear_worksheet(workbook_path, sheet_name):
    # Read the worksheet into a DataFrame
    df = pd.read_excel(workbook_path, sheet_name=sheet_name, engine='openpyxl')

    # Clear the values from row 2 onward
    df.iloc[1:] = None

    # Save the changes
    df.to_excel(workbook_path, sheet_name=sheet_name, index=False, engine='openpyxl')


def read_stock_data(workbook_path, sheet_name):
    """

    :type workbook_path: object
    """
    # Define the columns you want to read (from B to J)
    cols = "B:J"

    # Skip the first 2 rows to start reading from row 3
    skip_rows = 2

    # Read the specific range into a DataFrame
    df = pd.read_excel(workbook_path, sheet_name=sheet_name, usecols=cols, skiprows=skip_rows, engine='openpyxl')

    return df


def main_query(workbook_path, sheet_name):
    query_input_data = pd.read_excel(workbook_path, sheet_name=sheet_name)

    for index, row in query_input_data.iterrows():
        # 區域變數
        part_no = row['part_no_column']
        qty = row['qty_column']
        # 其他處理


if __name__ == "__main__":
    parameters_worksheet_path = r'C:\Users\windows user\Desktop\parameters_dataframe.xlsx'
    stock_workbook_path: str = r'\\192.168.1.220\神錡\sample標籤\出貨標籤\客戶出貨規範\崧騰\出貨明細\2023倉庫用\(NEW)2023.xlsx'
    clear_worksheet(workbook_path=parameters_worksheet_path, sheet_name='Input_Parameters')
    stock_data = read_stock_data(workbook_path=stock_workbook_path, sheet_name='20230105')
    main_query(workbook_path=parameters_worksheet_path, sheet_name='Query_Parameters')
