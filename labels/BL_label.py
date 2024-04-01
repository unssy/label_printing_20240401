import os
import pyautogui
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label
from datetime import datetime
def modify_and_print_BL_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """
    # Check if 'psi_datecode' is not a string
    if not isinstance(label_data['BL_datecode'], str):
        # Convert it to a string
        label_data['BL_datecode'] = str(label_data['BL_datecode'])
        date = datetime.strptime(label_data['BL_datecode'], "%Y-%m-%d %H:%M:%S")
        label_data['BL_datecode'] = date.strftime("%Y/%m/%d")
    # Extract parameters from label_data
    BL_pn = label_data['BL_pn']
    BL_lot = label_data['BL_lot']
    BL_datecode = label_data['BL_datecode']
    BL_qty = label_data['BL_qty']
    label_copies = label_data['label_copies']


    # Full path to the label file
    file_path = os.path.join(label_address, f"{BL_pn}.nlbl")

    try:
        process = launch_nicelabel(file_path)

        # Modify parameters
        modify_text_field(759, 597, BL_lot, type_i=False, first_call=True)
        modify_text_field(761, 693, BL_datecode, type_i=False)
        modify_text_field(725, 643, BL_qty, type_i=False)

        # Print and save the label
        print_and_save_label(label_copies, printer='D')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path}")

