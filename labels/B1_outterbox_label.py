import os
import pyautogui
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label
from datetime import datetime

def modify_and_print_B1_outterbox_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """

    if not isinstance(label_data['B1_datecode'], str):
        # Convert it to a string
        label_data['B1_datecode'] = str(label_data['B1_datecode'])
        date = datetime.strptime(label_data['B1_datecode'], "%Y-%m-%d %H:%M:%S")
        label_data['B1_datecode'] = date.strftime("%y%m%d")
    # Extract parameters from label_data
    B1_pn = label_data['B1_pn']
    B1_po = label_data['B1_po']
    B1_org = label_data['B1_org']
    B1_lot = label_data['B1_lot']
    B1_datecode = label_data['B1_datecode']
    B1_qty = label_data['B1_qty']
    B1_NO = str(label_data['B1_NO']).zfill(6)
    label_copies = label_data['label_copies']

    # Full path to the label file
    file_path = os.path.join(label_address, B1_pn, f"{B1_pn}-外箱標籤.nlbl")

    try:
        process = launch_nicelabel(file_path)

        # 向右End
        pyautogui.press('End')

        # Modify parameters
        modify_text_field(550, 357, B1_org, type_i=True, first_call=True)
        modify_text_field(600, 403, B1_po, type_i=True)
        modify_text_field(550, 500, B1_qty, type_i=True)
        modify_text_field(600, 560, B1_datecode, type_i=True)
        modify_text_field(600, 656, B1_NO, type_i=True)
        modify_text_field(600, 736, B1_lot, type_i=True)

        # Print and save the label
        print_and_save_label(label_copies, printer='D')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path}")

