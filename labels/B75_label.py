import os
import pyautogui
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label

def modify_and_print_B75_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """
    # Extract parameters from label_data
    B75_pn = label_data['B75_pn']
    B75_lot = label_data['B75_lot']
    B75_datecode = label_data['B75_datecode']
    B75_qty = label_data['B75_qty']
    if len(str(B75_qty)) <= 4:
        B75_qty = f'0{B75_qty}'
    B75_po = label_data['B75_po']
    B75_serial_number = label_data['B75_serial_number']
    label_copies = label_data['label_copies']

    # Full path to the label file
    file_path = os.path.join(label_address, f"{B75_pn}.nlbl")

    try:
        process = launch_nicelabel(file_path)

        # 向右End
        pyautogui.press('End')

        # Modify parameters
        modify_text_field(600, 442, B75_po,type_i=True,first_call = True)
        modify_text_field(600, 527, B75_lot,type_i=True)
        modify_text_field(600, 577, f'20{B75_datecode}',type_i=True)
        modify_text_field(600, 615, B75_qty,type_i=True)
        modify_text_field(600, 664, B75_serial_number,type_i=True)

        # Print and save the label
        print_and_save_label(label_copies, printer='D')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path}")