import os
import pyautogui
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label
def modify_and_print_lision_outterbox_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """
    # Extract parameters from label_data
    lision_pn = label_data['lision_pn']
    lision_po = label_data['lision_po']
    lision_qty = label_data['lision_qty']
    label_copies = label_data['label_copies']

       # Full path to the label file
    file_path = os.path.join(label_address, f"{lision_pn}.nlbl")

    try:
        process = launch_nicelabel(file_path)

        # Modify parameters
        modify_text_field(600, 476, lision_po, type_i=True, first_call=True)
        modify_text_field(600, 755, lision_qty, type_i=True)

        # Print and save the label
        print_and_save_label(label_copies, printer='D')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path}")

