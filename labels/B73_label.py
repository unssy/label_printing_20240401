import os
import pyautogui
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label
from datetime import datetime
def modify_and_print_B73_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """
    if not isinstance(label_data['delivery_date'], str):
        # Convert it to a string
        label_data['delivery_date'] = str(label_data['delivery_date'])
        date = datetime.strptime(label_data['delivery_date'], "%Y-%m-%d %H:%M:%S")
        label_data['delivery_date'] = date.strftime("%Y/%m/%d")
    # Extract parameters from label_data
    B73_pn = label_data['B73_pn']
    B73_lot = label_data['B73_lot']
    B73_datecode = label_data['B73_datecode']
    B73_qty = label_data['B73_qty']
    B73_po = label_data['B73_po']
    delivery_date = label_data['delivery_date']
    label_copies = label_data['label_copies']

    # Full path to the label file
    file_path = os.path.join(label_address, f"{B73_pn}.nlbl")

    try:
        process = launch_nicelabel(file_path)

        # Modify parameters
        modify_text_field(800, 489, B73_po, type_i=True, first_call=True)
        modify_text_field(800, 663, B73_lot, type_i=True)
        modify_text_field(800, 709, B73_datecode, type_i=True)
        modify_text_field(731, 753, B73_qty, type_i=True)
        modify_text_field(752, 801, delivery_date, type_i=True)

        # Print and save the label
        print_and_save_label(label_copies, printer='D')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path}")
