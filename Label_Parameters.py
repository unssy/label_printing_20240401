import pandas as pd


def send_parameters_to_printer(input_dataframe, calculation_dataframe):
    columns_to_drop = ['customer_part_number']
    input_dataframe = input_dataframe.drop(columns=columns_to_drop, errors='ignore')
    # 根據 'part_number' 和 'product_number' 合併 input_dataframe 和 calculation_dataframe
    merged_dataframe = pd.merge(calculation_dataframe, input_dataframe, on=['part_number', 'product_number'],
                                how='inner')
    # 按照指定的列順序重新排列
    desired_order = ['product_number', 'customer_part_number', 'new_lot', 'new_DC', 'new_date_code', 'new_quantity',
                     'small_label_quantity', 'small_label_copies', 'medium_label_quantity', 'medium_label_copies',
                     'large_label_quantity', 'large_label_copies', 'purchase_order', 'delivery_date']
    final_dataframe = merged_dataframe.reindex(columns=desired_order)

    return final_dataframe
