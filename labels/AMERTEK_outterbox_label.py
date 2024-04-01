import os
from datetime import datetime
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label

def modify_and_print_AMERTEK_outterbox_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """
    # Check if 'AMERTEK_label' is not a string
    if not isinstance(label_data['AMERTEK_datecode'], str):
        # Convert it to a string
        label_data['AMERTEK_datecode'] = str(label_data['AMERTEK_datecode'])
        date = datetime.strptime(label_data['AMERTEK_datecode'], "%Y-%m-%d %H:%M:%S")
        label_data['AMERTEK_datecode'] = date.strftime("%Y%m%d")
    # Extract parameters from label_data
    AMERTEK_pn = label_data['AMERTEK_pn']
    AMERTEK_PO = label_data['AMERTEK_PO']
    AMERTEK_datecode = label_data['AMERTEK_datecode']
    AMERTEK_package_qty = label_data['AMERTEK_package_qty']
    AMERTEK_shipping_qty = label_data['AMERTEK_shipping_qty']
    label_copies = label_data['label_copies']


    # Full path to the label file
    file_path = os.path.join(label_address, f"{AMERTEK_pn}.nlbl")

    try:
        process = launch_nicelabel(file_path)

        # Modify parameters
        modify_text_field(920, 348, AMERTEK_PO, type_i=False, first_call=True)
        modify_text_field(904, 555, AMERTEK_datecode, type_i=True)
        modify_text_field(911, 604, AMERTEK_package_qty, type_i=True)
        modify_text_field(912, 656, AMERTEK_shipping_qty, type_i=True)

        # Print and save the label
        print_and_save_label(label_copies, printer='A')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path}")

