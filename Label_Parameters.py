import pandas as pd


def send_parameters_to_printer(input_dataframe,calculation_dataframe):
    # 以['part_number', 'product_number']列合并两个DataFrame
    merged_dataframe = pd.merge(calculation_dataframe, input_dataframe, on=['part_number', 'product_number'],
                                how='left', suffixes=('', '_ref'))
    # 按照指定的列順序重新排列
    desired_order = ['purchase_order', 'delivery_date', 'product_number', 'customer_part_number', 'new_lot', 'new_DC', 'new_date_code', 'new_quantity',
                     'small_label_quantity', 'small_label_copies', 'medium_label_quantity', 'medium_label_copies',
                     'large_label_quantity', 'large_label_copies']
    final_dataframe = merged_dataframe[desired_order]

    return final_dataframe
