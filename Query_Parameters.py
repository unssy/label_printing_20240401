import pandas as pd

# 全域變數
sheet_name = "20230105"

# 打開工作簿並檢查是否為唯讀
def open_workbook(file_path):
    try:
        return pd.ExcelFile(file_path)
    except Exception as e:
        print("檔案已被開啟")
        exit()

# 定義數據範圍並清理格式
def read_data_range(wb, sheet_name, skiprows):
    return pd.read_excel(wb, sheet_name=sheet_name, skiprows=skiprows)

# 收集出貨信息
def collect_shipment_info():
    input_date = input("請輸入出貨日期:")
    customer = input("請輸入出貨客戶:")
    return input_date, customer

# 尋找特定日期之前的最後一列
def find_last_column_before_date(data, input_date):
    for i in range(data.columns.size - 1, -1, -1):
        if pd.to_datetime(data.columns[i]) < pd.to_datetime(input_date):
            return i
    return -1

# 插入新列並存儲出貨信息
def insert_shipment_column(data_range, input_date, customer):
    column_number = find_last_column_before_date(data_range, input_date)
    data_range.insert(column_number + 1, input_date, 0)
    data_range.loc[0, column_number + 1] = customer

# 遍歷扣帳表並更新庫存數據
debit_sheet = pd.read_excel(wb, sheet_name="扣帳表")
def update_stock_data(debit_sheet, data_range):
    for index, row in debit_sheet.iterrows():
        inputPartNo = row[2]
        inputlot = row[3]
        inputdate = row[5]
        inputQty = row[6]
    # ...
    pass

# 顯示出貨總數
def display_total_shipment(data_range, input_date):
    total_shipment = data_range[input_date].sum()
    print("出貨總數為:", total_shipment)

# 保存工作簿
def save_workbook(file_path, data_range, sheet_name):
    with pd.ExcelWriter(file_path) as writer:
        data_range.to_excel(writer, sheet_name=sheet_name, index=False)

# 主函數封裝整個流程
def stock_update(file_path):
    wb = open_workbook(file_path)
    data_range = read_data_range(wb)
    input_date, customer = collect_shipment_info()
    insert_shipment_column(data_range, input_date, customer)
    update_stock_data(debit_sheet, data_range)
    display_total_shipment(data_range, input_date)
    save_workbook(file_path, data_range)

if __name__ == "__main__":
    file_path = r"\\192.168.1.220\\神錡\\sample標籤\\出貨標籤\\客戶出貨規範\\崧騰\\出貨明細\\2023倉庫用\\(NEW)2023.xlsx"
    stock_update(file_path)
