import os
import pyautogui
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label
from datetime import datetime

def modify_and_print_B1_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """
    # Check if 'psi_datecode' is not a string
    if not isinstance(label_data['psi_datecode'], str):
        # Convert it to a string
        label_data['psi_datecode'] = str(label_data['psi_datecode'])
        date = datetime.strptime(label_data['psi_datecode'], "%Y-%m-%d %H:%M:%S")
        label_data['psi_datecode'] = date.strftime("%Y/%m/%d")
    # Extract parameters from label_data
    psi_pn = label_data['psi_pn']
    psi_lot = label_data['psi_lot']
    psi_datecode = label_data['psi_datecode']
    psi_qty = label_data['psi_qty']
    label_copies = label_data['label_copies']

    # Full path to the label file
    file_path_1 = os.path.join(label_address, psi_pn, f"{psi_pn}-PSI短版.nlbl")
    file_path_2 = os.path.join(label_address, psi_pn, f"{psi_pn}-全漢標籤(小、中).nlbl")
    try:
        process = launch_nicelabel(file_path_1)
        # Modify parameters
        modify_text_field(749, 601, psi_lot, type_i=False, first_call=True)
        modify_text_field(826, 684, psi_datecode, type_i=False)
        modify_text_field(737, 646, psi_qty, type_i=False)

        # Print and save the label
        print_and_save_label(label_copies, printer='D')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path_1}")

    try:
        process = launch_nicelabel(file_path_2)

        psi_datecode = datetime.strptime(psi_datecode, "%Y/%m/%d").strftime("%y%m%d")
        # Modify parameters
        modify_text_field(700, 720, psi_lot, type_i=True, first_call=True)
        modify_text_field(950, 526, psi_datecode, type_i=True)
        modify_text_field(950, 450, psi_qty, type_i=True)

        # Print and save the label
        print_and_save_label(label_copies, printer='D')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path_2}")

