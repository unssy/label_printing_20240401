import os
from datetime import datetime
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label

def modify_and_print_ADDA_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """
    # Check if 'ADDA_datecode' is not a string
    if not isinstance(label_data['ADDA_datecode'], str):
        # Convert it to a string
        label_data['ADDA_datecode'] = str(label_data['ADDA_datecode'])
        date = datetime.strptime(label_data['ADDA_datecode'], "%Y-%m-%d %H:%M:%S")
        label_data['ADDA_datecode'] = date.strftime("%Y/%m/%d")

    if not isinstance(label_data['delivery_date'], str):
        # Convert it to a string
        label_data['delivery_date'] = str(label_data['delivery_date'])
        date = datetime.strptime(label_data['delivery_date'], "%Y-%m-%d %H:%M:%S")
        label_data['delivery_date'] = date.strftime("%Y/%m/%d")

    # Extract parameters from label_data
    ADDA_pn = label_data['ADDA_pn']
    ADDA_month = label_data['ADDA_month']
    ADDA_po = label_data['ADDA_po']
    ADDA_lot = label_data['ADDA_lot']
    ADDA_datecode = label_data['ADDA_datecode']
    delivery_date = label_data['delivery_date']
    ADDA_qty = label_data['ADDA_qty']
    label_copies = label_data['label_copies']


    # Full path to the label file
    file_path = os.path.join(label_address, f"{ADDA_pn}.nlbl")

    try:
        process = launch_nicelabel(file_path)

        # Modify parameters
        modify_text_field(852, 468, ADDA_po, type_i=False, first_call=True)
        modify_text_field(914, 814, ADDA_lot, type_i=False)
        modify_text_field(858, 688, ADDA_datecode, type_i=False)
        modify_text_field(848, 643, delivery_date, type_i=False)
        modify_text_field(844, 598, ADDA_qty, type_i=False)
        modify_text_field(944, 367, ADDA_month, type_i=False)

        # Print and save the label
        print_and_save_label(label_copies, printer='A')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path}")

