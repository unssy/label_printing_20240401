import os
import pyautogui
from .utilities import launch_nicelabel, modify_text_field, print_and_save_label
def modify_and_print_lision_custom_pn_label(label_data, label_address):
    """
    Modify and print a label according to the given data.

    Args:
    label_data: A dictionary that contains the label parameters.
                Keys are parameter names and values are the corresponding parameter values.
    label_address: The file path of the label file.
    """
    # Extract parameters from label_data
    lision_pn = label_data['lision_pn']
    label_copies = label_data['label_copies']

    # Full path to the label file
    file_path = os.path.join(label_address, f"{lision_pn}(5x1).nlbl")
    try:
        process = launch_nicelabel(file_path)

        pyautogui.click(1082,507)
        pyautogui.hotkey('ctrl', 'space')

        # Print and save the label
        print_and_save_label(label_copies, printer='0')
        process.wait()  # Wait for the process to finish

    except FileNotFoundError:
        print(f"Could not find the file {file_path}")