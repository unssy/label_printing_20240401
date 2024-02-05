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
    # 指定你的配置文件路径
    config = load_config(get_workbook_path('config.json'))

    parameters_worksheet_path = config.get('parameters_worksheet_path', '')
    stock_workbook_path = config.get('stock_workbook_path', '')
    stock_workbook_sheet = config.get('stock_workbook_sheet', '')
    stock_password = config.get('stock_password', '')
    parameters_database = config.get('parameters_database.csv', '')
    input_dataframe = read_excel_data(workbook_path=parameters_worksheet_path, sheet_name='Input_Parameters')
    stock_dataframe = read_stock_data(workbook_path=stock_workbook_path, sheet_name=stock_workbook_sheet)
    # clear_worksheet()

    # Step 2: Query Stock
    query_dataframe = main_query(input_dataframe, stock_dataframe)
    if is_file_locked(parameters_worksheet_path):
        response = input("The file is currently open. Do you want to continue? (Enter 1 for yes, 2 for no): ").strip()
        if response != '1':
            print("Operation canceled.")
    else:
        output_data(workbook_path=parameters_worksheet_path,sheet_name='Query_Parameters',dataframe=query_dataframe)

    # # Step 3: Deduct Stock
    if ask_deduct_stock():
        print("Executing the desired function...")
        deduct_stock(file_path=stock_workbook_path, sheet=stock_workbook_sheet, data_df=query_dataframe, password=stock_password)
    else:
        print("Exiting the program...")
        exit()

    # Step 4: Get Calculation Parameters
    calculation_dataframe = query_dataframe.merge(input_dataframe[['part_number', 'customer_part_number']], on='part_number', how='left')
    calculation_dataframe = calculate_month_diff_dataframe(calculation_dataframe, threshold_months=24)
    calculation_dataframe = update_columns_based_on_expired(calculation_dataframe)
    calculation_dataframe = custom_standardization(calculation_dataframe)
    calculation_dataframe = calculate_label_copies_dataframe(calculation_dataframe)

    output_data('parameters_dataframe.xlsx','Calculation_Parameters',calculation_dataframe)

    # Step 5: Fill Delivery Record
    calculation_dataframe = read_excel_data(workbook_path=parameters_worksheet_path, sheet_name='Calculation_Parameters')
    delivery_record_dataframe = fill_delivery_record(input_dataframe,calculation_dataframe)
    output_data('parameters_dataframe.xlsx','delivery_record',delivery_record_dataframe)

    # Step 6: Send Parameters to Printer
    label_parameters_dataframe = send_parameters_to_printer(input_dataframe,calculation_dataframe)
    output_data('parameters_dataframe.xlsx', 'label_parameters', label_parameters_dataframe)

    # Step 7: Fill OQC Report
    oqc_report_dataframe = fill_oqc_report(input_dataframe,query_dataframe, delivery_record_dataframe,)
    output_data('parameters_dataframe.xlsx', 'OQC_report', oqc_report_dataframe)

