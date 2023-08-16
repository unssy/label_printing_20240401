import pandas as pd

# 全域變數
sheet_name = "20230105"

def get_sampling_count(shipment_count):
    shipment_bounds = [2, 9, 16, 26, 51, 91, 151, 281, 501, 1201, 3201, 10001, 35001, 150001, 500001]
    sampling_counts = [2, 3, 5, 8, 13, 20, 32, 50, 80, 125, 200, 315, 500, 800, 1250]

    # 二分搜索找到合適的區間
    left, right = 0, len(shipment_bounds) - 1
    while left <= right:
        mid = left + (right - left) // 2
        if shipment_bounds[mid] <= shipment_count < shipment_bounds[mid + 1]:
            return sampling_counts[mid]
        elif shipment_count < shipment_bounds[mid]:
            right = mid - 1
        else:
            left = mid + 1

    return None  # 如果數量不在任何區間內，返回 None

def merge_with_reference(your_dataframe):
    # 讀取參照表
    parameters_db = pd.read_csv('parameters_database.csv')

    # 將 part_number 設置為參照表的索引
    parameters_db.set_index('part_number', inplace=True)

    # 確保 your_dataframe 有 part_number 列
    if 'part_number' not in your_dataframe.columns:
        raise ValueError("Input DataFrame must contain a 'part_number' column.")

    # 將 part_number 設置為您的 DataFrame 的索引
    your_dataframe.set_index('part_number', inplace=True)

    # 使用 join 方法進行合併
    merged_dataframe = your_dataframe.join(parameters_db, how='left')

    # 將 part_number 轉回普通列
    merged_dataframe.reset_index(inplace=True)

    return merged_dataframe

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
    debit_sheet = pd.read_excel(wb, sheet_name="扣帳表")
    update_stock_data(debit_sheet, data_range)
    display_total_shipment(data_range, input_date)
    save_workbook(file_path, data_range)

if __name__ == "__main__":
    file_path = r"\\192.168.1.220\\神錡\\sample標籤\\出貨標籤\\客戶出貨規範\\崧騰\\出貨明細\\2023倉庫用\\(NEW)2023.xlsx"
    stock_update(file_path)
