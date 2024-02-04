import pandas as pd

def fill_oqc_report(input_dataframe,query_dataframe, delivery_record_dataframe):
    merged_dataframe = pd.merge(input_dataframe, query_dataframe, on=['part_number', 'product_number'], how='left',
                                suffixes=('', '_ref'))
    merged_dataframe = pd.merge(merged_dataframe, delivery_record_dataframe, on=['part_number', 'product_number'],
                                how='left', suffixes=('', '_record'))
    desired_order = ['purchase_order','product_number','customer_part_number','new_lot','new_quantity','marking_code','sampling']
    merged_dataframe = merged_dataframe[desired_order]
    # 如果您有特定的列順序需求，您可以在此處添加代碼來重新排列列，就像之前的函數一樣。

    return merged_dataframe
