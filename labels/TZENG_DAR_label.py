import os
import pyautogui
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label
from datetime import datetime
def modify_and_print_TZENG_DAR_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """
    # Check if 'TZENG_DAR_datecode' is not a string
    if not isinstance(label_data['TZENG_DAR_datecode'], str):
        # Convert it to a string
        label_data['TZENG_DAR_datecode'] = str(label_data['TZENG_DAR_datecode'])
        date = datetime.strptime(label_data['TZENG_DAR_datecode'], "%Y-%m-%d %H:%M:%S")
        label_data['TZENG_DAR_datecode'] = date.strftime("%Y/%m/%d")

    if not isinstance(label_data['delivery_date'], str):
        # Convert it to a string
        label_data['delivery_date'] = str(label_data['delivery_date'])
        date = datetime.strptime(label_data['delivery_date'], "%Y-%m-%d %H:%M:%S")
        label_data['delivery_date'] = date.strftime("%Y/%m/%d")

    # Extract parameters from label_data
    TZENG_DAR_pn = label_data['TZENG_DAR_pn']
    TZENG_DAR_month = label_data['TZENG_DAR_month']
    TZENG_DAR_po = label_data['TZENG_DAR_po']
    TZENG_DAR_lot = label_data['TZENG_DAR_lot']
    TZENG_DAR_datecode = label_data['TZENG_DAR_datecode']
    delivery_date = label_data['delivery_date']
    TZENG_DAR_qty = label_data['TZENG_DAR_qty']
    label_copies = label_data['label_copies']

       # Full path to the label file
    file_path = os.path.join(label_address, f"{TZENG_DAR_pn}.nlbl")

    try:
        process = launch_nicelabel(file_path)

        # Modify parameters
        modify_text_field(844, 457, TZENG_DAR_po, type_i=False, first_call=True)
        modify_text_field(887, 772, TZENG_DAR_lot, type_i=False)
        modify_text_field(999, 670, TZENG_DAR_datecode, type_i=False)
        modify_text_field(769, 678, delivery_date, type_i=False)
        modify_text_field(760, 612, TZENG_DAR_qty, type_i=False)
        modify_text_field(977, 345, TZENG_DAR_month, type_i=False)

        # Print and save the label
        print_and_save_label(label_copies, printer='A')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path}")

