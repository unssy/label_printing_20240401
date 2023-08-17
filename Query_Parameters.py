import pandas as pd

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

