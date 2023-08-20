import pandas as pd

def fill_delivery_record(input_dataframe, query_dataframe, calculation_dataframe):
    # 首先根據 'part_number' 和 'product_number' 合併 calculation_dataframe 和 query_dataframe
    columns_to_drop = ['MPQ','package']
    query_dataframe = query_dataframe.drop(columns=columns_to_drop, errors='ignore')
    merged_with_query = pd.merge(calculation_dataframe, query_dataframe, on=['part_number', 'product_number'], how='inner')
    # 然後根據 'part_number' 和 'product_number' 合併 merged_with_query 和 input_dataframe
    columns_to_drop_2 = ['customer_part_number','quantity','package','delivery_date','customer_no']
    input_dataframe = input_dataframe.drop(columns=columns_to_drop_2, errors='ignore')
    final_merged_dataframe = pd.merge(merged_with_query, input_dataframe, on=['part_number', 'product_number'], how='inner')
    # 選擇所需的列並重新排列
    desired_order = ['delivery_date','invoice_series','customer_no','part_number', 'lot', 'DC', 'date_code', 'quantity', 'customer_part_number', 'product_number', 'new_lot', 'new_DC', 'new_date_code','new_quantity', 'marking_code', 'package', 'unit_price', 'currency']
    final_merged_dataframe = final_merged_dataframe[desired_order]

    return final_merged_dataframe

