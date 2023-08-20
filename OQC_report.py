import pandas as pd

def fill_oqc_report(input_dataframe, delivery_record_dataframe):
    # 根據 'part_number' 和 'product_number' 合併 input_dataframe 和 delivery_record_dataframe
    input_subset = input_dataframe[['part_number', 'product_number', 'customer_part_number', 'purchase_order']]
    # 使用 'part_number', 'product_number', 和 'customer_part_number' 作為連接鍵進行左連接
    merged_dataframe = pd.merge(delivery_record_dataframe, input_subset,
                                on=['part_number', 'product_number', 'customer_part_number'], how='left')
    desired_order = ['purchase_order','product_number','customer_part_number','new_lot','new_quantity','marking_code']
    merged_dataframe = merged_dataframe[desired_order]
    # 如果您有特定的列順序需求，您可以在此處添加代碼來重新排列列，就像之前的函數一樣。

    return merged_dataframe
