import pandas as pd
from openpyxl import load_workbook
from  utinities import *
import os
from Input_Parameters import *
from Query_Parameters import *
from Calculation_Parameters import *
from Delivery_Record import *
from Label_Parameters import *
from OQC_report import *


def get_workbook_path(filename):
    """
    根據給定的文件名返回在當前腳本所在目錄的絕對路徑。

    參數:
        filename (str): 要查找的文件名。

    返回:
        str: 文件的絕對路徑。
    """
    current_file_path = os.path.abspath(__file__)
    current_directory = os.path.dirname(current_file_path)

    workbook_path = os.path.join(current_directory, filename)

    return workbook_path



if __name__ == "__main__":
    # Step 1: Get Input Parameters
    # clear_worksheet(workbook_path=parameters_worksheet_path, sheet_name='Input_Parameters')
    parameters_worksheet_path = get_workbook_path('parameters_dataframe.xlsx')
    stock_workbook_path = get_workbook_path('(NEW)2023-2.xlsx')
    input_dataframe = read_excel_data(workbook_path=parameters_worksheet_path, sheet_name='Input_Parameters')
    stock_dataframe = read_stock_data(workbook_path=stock_workbook_path, sheet_name='20230105')

    # clear_worksheet()

    # Step 2: Query Stock
    query_dataframe = main_query(input_dataframe, stock_dataframe)
    query_dataframe = get_sampling_count(query_dataframe)
    query_dataframe = merge_with_reference(query_dataframe)
    query_dataframe['delivery_date'] = input_dataframe['delivery_date'].iloc[0]
    query_dataframe['customer_no'] = input_dataframe['customer_no'].iloc[0]
    output_data(workbook_path=parameters_worksheet_path,sheet_name='Query_Parameters',dataframe=query_dataframe)

    # Step 3: Deduct Stock
    if ask_deduct_stock():
        print("Executing the desired function...")
        deduct_stock(file_path=stock_workbook_path, sheet='20230105', data_df=query_dataframe)
    else:
        print("Exiting the program...")
        exit()

    # Step 4: Get Calculation Parameters
    calculation_dataframe = query_dataframe.merge(input_dataframe[['part_number', 'product_number', 'customer_part_number']], on='part_number', how='left')
    columns_to_remove = ['store', 'row_index', 'remark', 'sampling', 'marking_code',]
    desired_order = ['part_number','product_number','customer_part_number','lot','DC','date_code','quantity','label_copies_large','label_copies_medium','label_copies_small','out_date','MPQ']
    calculation_dataframe = calculation_dataframe[desired_order]
    calculation_dataframe = calculation_dataframe.drop(columns=columns_to_remove, errors='ignore')
    calculation_dataframe = calculate_month_diff_dataframe(calculation_dataframe)
    calculation_dataframe = calculate_label_copies_dataframe(calculation_dataframe)
    output_data('parameters_dataframe.xlsx','Calculation_Parameters',calculation_dataframe)

    # Step 5: Fill Delivery Record
    columns_to_remove = ['small_label_quantity','small_label_copies','medium_label_quantity','medium_label_copies','large_label_quantity','large_label_copies','out_date','MPQ']
    calculation_dataframe = calculation_dataframe.drop(columns=columns_to_remove, errors='ignore')
    desired_order = ['part_number','product_number','customer_part_number','lot','DC','date_code','quantity','label_copies_large','label_copies_medium','label_copies_small','out_date','MPQ']
    delivery_record_dataframe = fill_delivery_record(input_dataframe, query_dataframe)
    output_data('parameters_dataframe.xlsx','delivery_record',delivery_record_dataframe)

    # Step 6: Send Parameters to Printer
    label_parameters_dataframe = send_parameters_to_printer(input_dataframe, query_dataframe, calculation_dataframe)
    output_data('parameters_dataframe.xlsx', 'label_parameters', label_parameters_dataframe)

    # Step 7: Fill OQC Report
    oqc_report_dataframe = fill_oqc_report(input_dataframe, delivery_record_dataframe)
    output_data('parameters_dataframe.xlsx', 'OQC_report', oqc_report_dataframe)

