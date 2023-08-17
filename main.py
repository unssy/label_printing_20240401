import os
import pandas as pd
from Input_Parameters import get_input_parameters
from Query_Parameters import query_stock, deduct_stock
from Calculation_Parameters import get_calculation_parameters
from Delivery_Record import fill_delivery_record
from Label_Parameters import send_parameters_to_printer
from OQC_report import fill_oqc_report


def main():
    # Step 1: Get Input Parameters
    input_dataframe = get_input_parameters('parameters_dataframe.xlsx', 'Input_Parameters')
    stock_dataframe = get_input_parameters('(NEW)2023-2.xlsx', '20230105')
    clear_worksheet()

    # Step 2: Query Stock
    query_dataframe = query_stock(stock_dataframe, input_dataframe)
    query_dataframe = get_sampling_count(query_dataframe)
    query_dataframe = merge_with_reference(query_dataframe)
    output_data('parameters_dataframe.xlsx','Query_Parameters',query_dataframe)

    # Step 3: Deduct Stock
    if input("Do you want to deduct stock? (yes/no): ").strip().lower() == 'yes':
        deduct_stock(query_dataframe, '(NEW)2023-2.xlsx')

    # Step 4: Get Calculation Parameters
    calculation_dataframe = get_calculation_parameters(input_dataframe, query_dataframe)
    output_data('parameters_dataframe.xlsx','Calculation_Parameters',calculation_dataframe)

    # Step 5: Fill Delivery Record
    delivery_record_dataframe = fill_delivery_record(input_dataframe, query_dataframe)
    output_data('parameters_dataframe.xlsx','delivery_record',delivery_record_dataframe)

    # Step 6: Send Parameters to Printer
    label_parameters_dataframe = send_parameters_to_printer(input_dataframe, query_dataframe, calculation_dataframe)
    output_data('parameters_dataframe.xlsx', 'label_parameters', label_parameters_dataframe)

    # Step 7: Fill OQC Report
    oqc_report_dataframe = fill_oqc_report(input_dataframe, delivery_record_dataframe)
    output_data('parameters_dataframe.xlsx', 'OQC_report', oqc_report_dataframe)

if __name__ == "__main__":
    main()
