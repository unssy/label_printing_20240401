from Input_Parameters import *
from Query_Parameters import *
from utilities import *
from Calculation_Parameters import *


def get_workbook_path(filename):
    """
    Returns the absolute path of the file in the current script directory.
    Args:
        filename (str): The name of the file to find.
    Returns:
        str: The absolute path of the file.
    """
    current_file_path = os.path.abspath(__file__)
    current_directory = os.path.dirname(current_file_path)
    workbook_path = os.path.join(current_directory, filename)
    return workbook_path


def slice_dataframe(main_dataframe, dataframe_name):
    parameters_mapping = {
        'delivery_record_dataframe': ['month', 'day', 'invoice_series', 'purchase_order', 'customer_no',
                                      'customer_name',
                                      'lot', 'part_number', 'DC', 'date_code', 'quantity'],
        'label_parameters_dataframe': ['delivery_date', 'invoice_series', 'customer_no', 'customer_name', 'part_number',
                                       'lot', 'DC', 'quantity', 'customer_part_number', 'product_number', 'new_lot',
                                       'new_DC',
                                       'new_quantity', 'marking_code', 'small_label_quantity', 'small_label_copies',
                                       'medium_label_quantity', 'medium_label_copies', 'large_label_quantity',
                                       'large_label_copies', ],
        'oqc_report_dataframe': ['purchase_order', 'product_number', 'customer_part_number', 'new_lot', 'new_quantity',
                                 'marking_code', 'sampling', 'DC'],
        'query_dataframe': ['delivery_date', 'invoice_series', 'customer_name', 'part_number', 'lot', 'DC', 'quantity',
                            'customer_part_number', 'marking_code', 'package', 'MPQ', 'store', 'out_date', 'expired',
                            'row_index']
    }

    parameters = parameters_mapping.get(dataframe_name, [])
    sliced_dataframe = main_dataframe[parameters]
    return sliced_dataframe


def save_to_database(dataframe, excel_file_path, sheet_name='工作表1'):
    try:
        # 读取现有的 Excel 文件，如果不存在则创建一个新的 DataFrame
        existing_data = pd.read_excel(excel_file_path, sheet_name=sheet_name) if pd.ExcelFile(excel_file_path).sheet_names else pd.DataFrame()

        # 将新的 dataframe 添加到现有数据库，确保新的行索引从零开始
        updated_data = pd.concat([existing_data, dataframe], ignore_index=True)

        # 存入 Excel 文件
        updated_data.to_excel(excel_file_path, index=False, sheet_name=sheet_name)

        print(f"数据成功保存到 {excel_file_path} 的 {sheet_name} 工作表。")

    except Exception as e:
        print(f"发生错误：{str(e)}")



def initialize_and_generate_forms(config_path):
    global main_dataframe
    global query_dataframe
    column_order = ['delivery_date', 'invoice_series', 'purchase_order', 'customer_no', 'customer_name',
                    'lot', 'part_number', 'DC', 'date_code', 'quantity', 'customer_part_number', 'product_number',
                    'new_lot', 'new_DC', 'new_date_code', 'new_quantity', 'marking_code', 'package', 'unit_price',
                    'currency', 'small_label_quantity', 'small_label_copies', 'medium_label_quantity',
                    'medium_label_copies',
                    'large_label_quantity', 'large_label_copies',
                    'out_date', 'expired', 'store', 'row_index', 'remark', 'sampling', 'MPQ', 'deduct',
                    'month', 'day']
    main_dataframe = pd.DataFrame(columns=column_order)

    config = load_config(get_workbook_path(config_path))
    parameters_worksheet_path = config.get('parameters_worksheet_path', '')
    stock_workbook_path = config.get('stock_workbook_path', '')
    stock_workbook_sheet = config.get('stock_workbook_sheet', '')
    main_dataframe_database_path =config.get('main_dataframe_database_path', '')
    stock_password = config.get('stock_password', '')
    parameters_database = config.get('parameters_database.csv', '')

    input_dataframe = read_excel_data(workbook_path=parameters_worksheet_path, sheet_name='Input_Parameters')
    preprocess_input_dataframe = new_preprocess_input_data(input_dataframe)
    stock_dataframe = read_stock_data(workbook_path=stock_workbook_path, sheet_name=stock_workbook_sheet)
    query_dataframe = main_query(preprocess_input_dataframe, stock_dataframe)
    calculation_dataframe = preprocess_calculation_dataframe(query_dataframe)
    main_dataframe[column_order] = calculation_dataframe[column_order]

    delivery_record_dataframe = slice_dataframe(main_dataframe, 'delivery_record_dataframe')
    label_parameters_dataframe = slice_dataframe(main_dataframe, 'label_parameters_dataframe')
    oqc_report_dataframe = slice_dataframe(main_dataframe, 'oqc_report_dataframe')
    query_dataframe = slice_dataframe(main_dataframe, 'query_dataframe')

    if is_file_locked(parameters_worksheet_path):
        response = input("The file is currently open. Do you want to continue? (Enter 1 for yes, 2 for no): ").strip()
        if response == '1':
            output_data('parameters_dataframe.xlsx', 'main_dataframe', main_dataframe)
            output_data('parameters_dataframe.xlsx', 'delivery_record', delivery_record_dataframe)
            output_data('parameters_dataframe.xlsx', 'label_parameters', label_parameters_dataframe)
            output_data('parameters_dataframe.xlsx', 'OQC_report', oqc_report_dataframe)
            output_data('parameters_dataframe.xlsx', 'query_dataframe', query_dataframe)
        elif response == '2':
            print("Operation canceled.")
        else:
            print("Invalid input. Operation canceled.")
    else:
        output_data('parameters_dataframe.xlsx', 'main_dataframe', main_dataframe)
        output_data('parameters_dataframe.xlsx', 'delivery_record', delivery_record_dataframe)
        output_data('parameters_dataframe.xlsx', 'label_parameters', label_parameters_dataframe)
        output_data('parameters_dataframe.xlsx', 'OQC_report', oqc_report_dataframe)
        output_data('parameters_dataframe.xlsx', 'query_dataframe', query_dataframe)
    save_to_database(main_dataframe, excel_file_path=main_dataframe_database_path, sheet_name='工作表1')


def preprocess_calculation_dataframe(query_dataframe):
    df = query_dataframe.copy()
    df = calculate_month_diff(df, threshold_months=24)
    df = update_columns_based_on_expired(df)
    df = custom_standardization(df)
    df = calculate_label_copies(df)
    return df


def main():

    while True:
        print("Please select an operation:")
        print("1. Initialize and generate forms")
        print("2. Manually update read forms")
        print("3. Deduct")
        print("4. Exit")

        choice = input("Enter the operation number: ")

        if choice == "1":
            print('initialize_and_generate_forms')
            main_dataframe = initialize_and_generate_forms('config.json')
        elif choice == "2":
            main_dataframe = read_excel_data(workbook_path=config.get('parameters_worksheet_path', ''), sheet_name='main_dataframe')
            print_query_dataframe = input(
                "Do you want to print the query_dataframe before updating forms? Enter 1 for yes: ")
            if print_query_dataframe == 1:
                global query_dataframe
                print(query_dataframe.head(5))  # Display the first five rows of query_dataframe

        elif choice == "3":
            if main_dataframe is None:
                print("Please perform the initialize and generate forms operation first (select 1).")
            else:
                stock_workbook_path = config.get('stock_workbook_path', '')
                stock_workbook_sheet = config.get('stock_workbook_sheet', '')
                stock_password = config.get('stock_password', '')
                query_dataframe = slice_dataframe(main_dataframe, 'query_dataframe')
                parameters_worksheet_path = config.get('parameters_worksheet_path', '')
                if ask_deduct_stock():
                    print("Executing the desired function...")
                    main_dataframe = read_excel_data(workbook_path=parameters_worksheet_path,
                                                     sheet_name='main_dataframe')
                    deduct_stock(file_path=stock_workbook_path, sheet=stock_workbook_sheet, data_df=query_dataframe,
                                 password=stock_password)
                    main_dataframe['deduct'] = True
                else:
                    print("Exiting the program...")
                    exit()
        elif choice == "4":
            break
        else:
            print("Invalid operation number, please try again.")


if __name__ == "__main__":
    global query_dataframe, main_dataframe, config
    query_dataframe = None
    main_dataframe = None
    config = load_config(get_workbook_path('config.json'))
    invoice_series = input("Enter the invoice_series: ")
    main()
