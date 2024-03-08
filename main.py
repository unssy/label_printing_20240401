from Input_Parameters import *
from Query_Parameters import *
from Calculation_Parameters import *
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

def slice_dataframe(main_dataframe, dataframe_name):
    parameters_mapping = {
        'delivery_record_dataframe': ['month', 'day', 'invoice_series', 'purchase_order', 'customer_no', 'customer_name',
                                      'lot', 'part_number', 'DC', 'date_code', 'quantity'],
        'label_parameters_dataframe': ['delivery_date', 'invoice_series', 'customer_no', 'customer_name', 'part_number',
                                       'lot', 'DC', 'quantity', 'customer_part_number', 'product_number', 'new_lot', 'new_DC',
                                       'new_quantity', 'marking_code', 'package', 'unit_price', 'currency'],
        'oqc_report_dataframe': ['purchase_order', 'product_number', 'customer_part_number', 'new_lot', 'new_quantity',
                                 'marking_code', 'sampling', 'DC'],
    }

    parameters = parameters_mapping.get(dataframe_name, [])
    sliced_dataframe = main_dataframe[parameters]
    return sliced_dataframe

if __name__ == "__main__":

    column_order = ['delivery_date', 'invoice_series', 'purchase_order', 'customer_no', 'customer_name',
                    'lot', 'part_number', 'DC', 'date_code', 'quantity', 'customer_part_number', 'product_number',
                    'new_lot', 'new_DC', 'new_date_code', 'new_quantity', 'marking_code', 'package', 'unit_price',
                    'currency', 'small_label_quantity', 'small_label_copies', 'medium_label_quantity',
                    'medium_label_copies',
                    'large_label_quantity', 'large_label_copies',
                    'out_date', 'expired', 'store', 'row_index', 'remark', 'sampling', 'MPQ', 'deduct',
                    'month', 'day']
    main_dataframe = pd.DataFrame(columns=column_order)
    config = load_config(get_workbook_path('config.json'))
    parameters_worksheet_path = config.get('parameters_worksheet_path', '')
    stock_workbook_path = config.get('stock_workbook_path', '')
    stock_workbook_sheet = config.get('stock_workbook_sheet', '')
    stock_password = config.get('stock_password', '')
    parameters_database = config.get('parameters_database.csv', '')
    input_dataframe = read_excel_data(workbook_path=parameters_worksheet_path, sheet_name='Input_Parameters')
    preprocess_input_dataframe=preprocess_input_dataframe(input_dataframe)
    stock_dataframe = read_stock_data(workbook_path=stock_workbook_path, sheet_name=stock_workbook_sheet)
    query_dataframe = main_query(preprocess_input_dataframe, stock_dataframe)
    calculation_dataframe = query_dataframe
    calculation_dataframe = calculate_month_diff_dataframe(calculation_dataframe, threshold_months=24)
    calculation_dataframe = update_columns_based_on_expired(calculation_dataframe)
    calculation_dataframe = custom_standardization(calculation_dataframe)
    calculation_dataframe = calculate_label_copies_dataframe(calculation_dataframe)
    main_dataframe[column_order] = calculation_dataframe[column_order]

    delivery_record_dataframe=slice_dataframe(main_dataframe, 'delivery_record_dataframe')
    label_parameters_dataframe = slice_dataframe(main_dataframe, 'label_parameters_dataframe')
    oqc_report_dataframe = slice_dataframe(main_dataframe, 'oqc_report_dataframe')


    if is_file_locked(parameters_worksheet_path):
        response = input("The file is currently open. Do you want to continue? (Enter 1 for yes, 2 for no): ").strip()
        if response != '1':
            print("Operation canceled.")
    else:
        output_data('parameters_dataframe.xlsx', 'main_dataframe', main_dataframe)
        output_data('parameters_dataframe.xlsx', 'delivery_record', delivery_record_dataframe)
        output_data('parameters_dataframe.xlsx', 'label_parameters', label_parameters_dataframe)
        output_data('parameters_dataframe.xlsx', 'OQC_report', oqc_report_dataframe)


    if ask_deduct_stock():
        print("Executing the desired function...")
        main_dataframe = read_excel_data(workbook_path=parameters_worksheet_path, sheet_name='main_dataframe')
        deduct_stock(file_path=stock_workbook_path, sheet=stock_workbook_sheet, data_df=query_dataframe, password=stock_password)
        main_dataframe['deduct'] = True
    else:
        print("Exiting the program...")
        exit()