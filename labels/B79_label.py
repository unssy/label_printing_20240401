import clipboard
from datetime import datetime
import time
import os
import pyautogui
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label


def modify_and_print_B79_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """
    # Check if 'ADDA_datecode' is not a string
    if not isinstance(label_data['B79_datecode'], str):
        # Convert it to a string
        label_data['B79_datecode'] = str(label_data['B79_datecode'])
        date = datetime.strptime(label_data['B79_datecode'], "%Y-%m-%d %H:%M:%S")
        label_data['B79_datecode'] = date.strftime("%Y/%m/%d")

    if not isinstance(label_data['delivery_date'], str):
        # Convert it to a string
        label_data['delivery_date'] = str(label_data['delivery_date'])
        date = datetime.strptime(label_data['delivery_date'], "%Y-%m-%d %H:%M:%S")
        label_data['delivery_date'] = date.strftime("%Y/%m/%d")

    # Extract parameters from label_data
    B79_pn = label_data['B79_pn']
    B79_qty = label_data['B79_qty']
    B79_po = label_data['B79_po']
    B79_box = label_data['B79_box']
    total_weight = label_data['total_weight']
    B79_datecode = label_data['B79_datecode']
    delivery_date = label_data['delivery_date']
    label_copies = label_data['label_copies']
    last_box_text = label_data['last_box_text']

    if label_data['last_box_text'] == '\u662f\u25a0\u5426\u25a1':
        last_box_text = '\u662f\u25a0\u5426\u25a1'  # '是■否□'
    else:
        last_box_text = '\u662f\u25a1\u5426\u25a0'  # '是□否■'



    # Full path to the label file
    file_path = os.path.join(label_address, f"{B79_pn}.nlbl")

    try:
        process = launch_nicelabel(file_path)

        # Modify parameters
        modify_text_field(719, 461, B79_qty, type_i=False, first_call=True)
        modify_text_field(719, 552, B79_qty, type_i=False)
        modify_text_field(719, 499, B79_po, type_i=False)
        modify_text_field(758, 673, B79_box, type_i=False)
        modify_text_field(1140, 460, total_weight, type_i=False)
        modify_text_field(1070, 505, B79_datecode, type_i=False)
        modify_text_field(1070, 551, delivery_date, type_i=False)
        modify_text_field(1070, 674, '', type_i=False)

        pyautogui.click(1598, 592)
        pyautogui.hotkey('ctrl', 'a')
        # Copy last_box_text to clipboard and paste
        time.sleep(0.5)
        clipboard.copy(last_box_text)
        pyautogui.hotkey('ctrl', 'v')

        # Print and save the label
        print_and_save_label(label_copies, printer='A')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path}")


