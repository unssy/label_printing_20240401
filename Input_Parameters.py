import pandas as pd
from openpyxl import load_workbook
import os
from  utinities import format_date_code

def read_stock_data(workbook_path, sheet_name):
    workbook = load_workbook(workbook_path, data_only=True)

    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"工作表名稱 '{sheet_name}' 在工作簿中不存在。")

    worksheet = workbook[sheet_name]

    # 使用原始方法提取數據值
    data = [[cell.value for cell in row] for row in worksheet['B3':'J' + str(worksheet.max_row)]]

    # 使用單獨的循環來提取註釋
    remarks = []
    for row in worksheet['C3':'C' + str(worksheet.max_row)]:
        remark = row[0].comment.text if row[0].comment else None
        remarks.append(remark)

    columns = ['supplier_code', 'part_number', 'lot', "date_code", 'DC', 'QTY', 'quantity', 'Y', 'store']
    df = pd.DataFrame(data[1:], columns=columns)
    df = df.drop(columns=['supplier_code', 'QTY', 'Y'])
    df['remark'] = remarks[1:]

    return df





