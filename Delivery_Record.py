import pandas as pd

def fill_delivery_record(input_dataframe, calculation_dataframe):
    # 以['part_number', 'product_number']列合并两个DataFrame
    merged_dataframe = pd.merge(calculation_dataframe, input_dataframe, on=['part_number', 'product_number'],
                                how='left', suffixes=('', '_ref'))
    # 拆分delivery_date并添加month和day列
    merged_dataframe['month'] = pd.to_datetime(merged_dataframe['delivery_date']).dt.month
    merged_dataframe['day'] = pd.to_datetime(merged_dataframe['delivery_date']).dt.day

    # 选择所需的列并重新排列
    desired_order = ['month', 'day', 'invoice_series', 'customer_no', 'customer_name' , 'part_number', 'lot', 'DC',
                     'quantity', 'customer_part_number', 'product_number', 'new_lot', 'new_DC',
                     'new_quantity', 'marking_code', 'package', 'unit_price', 'currency']
    final_merged_dataframe = merged_dataframe[desired_order]

    return final_merged_dataframe

