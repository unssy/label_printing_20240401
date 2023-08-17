from  utinities import *
import pandas as pd
import win32com.client as win32

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
    # 讀取參照表
    parameters_db = pd.read_csv('parameters_database.csv')

    # 將 part_number 設置為參照表的索引
    parameters_db.set_index('part_number', inplace=True)

    # 確保 your_dataframe 有 part_number 列
    if 'part_number' not in dataframe.columns:
        raise ValueError("Input DataFrame must contain a 'part_number' column.")

    # 將 part_number 設置為您的 DataFrame 的索引
    dataframe.set_index('part_number', inplace=True)

    # 使用 join 方法進行合併
    merged_dataframe = dataframe.join(parameters_db, how='left')

    # 將 part_number 轉回普通列
    merged_dataframe.reset_index(inplace=True)
    merged_dataframe['MPQ'].fillna(0, inplace=True)
    merged_dataframe['MPQ'] = merged_dataframe['MPQ'].astype(int)

    return merged_dataframe

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
    return recommend_df


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

def convert_to_mm_dd(date_str):
    # 提取月份和日期
    month = date_str[5:6]
    day = date_str[6:8]
    return f"{month}/{day}"


# 主函數封裝整個流程
def deduct_stock(workbook_path, sheet_name, dataframe):
    # 打开 Excel
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    wb = excel.Workbooks.Open(workbook_path, ReadOnly=True, UpdateLinks=0)
    ws = wb.Worksheets(sheet_name)

    # 尋找特定日期 delivery_date之前的最後一列 column=k
    delivery_date = dataframe['delivery_date'].iloc[0]
    delivery_date_format = convert_to_mm_dd(str(delivery_date))
    col_k = None
    for col in range(ws.UsedRange.Columns.Count, 0, -1):  # Start from the last used column and move backward
        cell_value = ws.Cells(1, col).Value
        if cell_value and cell_value < delivery_date_format:
            col_k = col+1
            break

    if col_k is None:
        print("找不到特定日期之前的最後一列。")
        return

    # 插入新列
    ws.Columns(col_k).Insert()

    # 存储出货信息
    ws.Cells(1, col_k).Value = delivery_date_format
    ws.Cells(2, col_k).Value = dataframe['customer_no']
    for i, row in dataframe.iterrows():
        row_idx = row['row_index']
        qty = row['quantity']
        ws.Cells(row_idx, col_k).Value = qty

    # wb.Save()
    excel.Application.Quit()

