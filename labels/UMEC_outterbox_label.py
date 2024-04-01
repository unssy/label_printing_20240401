import os
import pyautogui
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label
def modify_and_print_UMEC_outterbox_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """
    # Extract parameters from label_data
    UMEC_pn = label_data['UMEC_pn']
    UMEC_datecode = label_data['UMEC_datecode']
    delivery_date = label_data['delivery_date']
    UMEC_qty = label_data['UMEC_qty']
    total_qty = label_data['total_qty']
    UMEC_po = label_data['UMEC_po']
    label_copies = label_data['label_copies']

    # Full path to the label file
    file_path = os.path.join(label_address, f"{UMEC_pn}.nlbl")

    try:
        process = launch_nicelabel(file_path)

        # 向右End
        pyautogui.press('End')

        # Modify parameters
        modify_text_field(600, 573, delivery_date, type_i=True, first_call=True)
        modify_text_field(600, 522, UMEC_datecode, type_i=True)
        modify_text_field(600, 645, UMEC_qty, type_i=True)
        modify_text_field(600, 688, total_qty, type_i=True)
        modify_text_field(600, 730, UMEC_po, type_i=True)

        # Print and save the label
        print_and_save_label(label_copies, printer='D')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path}")
                
