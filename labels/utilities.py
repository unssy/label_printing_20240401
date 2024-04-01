import os
import time
import subprocess
import pyautogui
import ctypes

# Path to NiceLabel Designer executable
nicelabel_path = r'C:\Program Files\NiceLabel\NiceLabel 2017\bin.net\NiceLabelDesigner.exe'


def get_user_input(prompt: str) -> str:
    while True:
        choice = input(prompt)
        if choice in ['1', '2']:
            return choice
        else:
            print("Invalid input. Please enter 1 or 2.")

def check_caps_lock():
    # Virtual key code for Caps Lock
    VK_CAPITAL = 0x14

    # Get the current state of Caps Lock
    caps_lock_state = ctypes.windll.user32.GetKeyState(VK_CAPITAL)

    # Check the least significant bit; if 1, Caps Lock is currently ON
    if caps_lock_state & 1:
        print('Caps Lock is ON.')

        # Simulate pressing and releasing the Caps Lock key to turn it OFF
        ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 2, 0)
        ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 0, 0)
    else:
        print('Caps Lock is OFF.')


def check_edit_button():
    # Get the start time
    start_time = time.time()
    # Edit button coordinates
    edit_button_coord = (436, 30)
    # Expected color of the edit button
    expected_edit_button_rgb = (129, 174, 234)

    while True:
        # Get the pixel color at the specified coordinates
        rgb = pyautogui.pixel(*edit_button_coord)

        # If the pixel color matches the expected color, return True
        if rgb == expected_edit_button_rgb:
            return True

        # If it has been more than 1 second, raise an exception
        if time.time() - start_time > 3:
            raise Exception("The edit button did not appear within 3 second.")

def check_print_button():
    # Get the start time
    start_time = time.time()
    # Edit button coordinates
    edit_print_coord = (1740, 156)
    # Expected color of the edit button
    expected_print_button_rgb = (23, 168, 93)

    while True:
        # Get the pixel color at the specified coordinates
        rgb = pyautogui.pixel(*edit_print_coord)

        # If the pixel color matches the expected color, return True
        if rgb == expected_print_button_rgb:
            return True

        # If it has been more than 1 second, raise an exception
        if time.time() - start_time > 3:
            raise Exception("The edit button did not appear within 3 second.")


def is_nicelabel_running(file_path: str) -> bool:
    """Check if NiceLabel process is already running for the given file."""
    process_name = "NiceLabelDesigner.exe"  # Update with the actual process name

    # Run 'tasklist' command and check if the process is in the output
    output = subprocess.run(['tasklist'], capture_output=True, text=True)
    return process_name.lower() in output.stdout.lower()

def is_nicelabel_ready():
    # Get the start time
    start_time = time.time()
    # Correction button coordinates
    coorection_button_coord = (1352, 45)
    # Expected color of the correction button
    coorection_button_rgb = (0, 45, 89)

    while True:
        windows = pyautogui.getWindowsWithTitle("NiceLabel Designer")
        if len(windows) > 0:
            window = windows[0]
            window.activate()
            window.maximize()  # Maximizing the window

            # Get the pixel color at the specified coordinates
            rgb = pyautogui.pixel(*coorection_button_coord)

            # If the pixel color matches the expected color, return True
            if rgb == coorection_button_rgb:
                # 縮放至文件
                pyautogui.click(1554, 976)
                return True

        # If it has been more than 10 seconds, raise an exception
        if time.time() - start_time > 10:
            raise Exception("The correction button did not appear within 10 seconds.")

def launch_nicelabel(file_path):
    if os.path.exists(file_path):
        if is_nicelabel_running(file_path):
            user_input = get_user_input(f"The file is already open in NiceLabel. Continue launching NiceLabel? (1: Yes, 2: No): ")
            if user_input == '2':
                print("Launching NiceLabel aborted.")
                return None  # 返回 None 表示终止程序

        process = subprocess.Popen([nicelabel_path, file_path])
        is_nicelabel_ready()
        return process
    else:
        raise FileNotFoundError(f"File does not exist: {file_path}")

def modify_text_field(x, y, text, type_i=False,first_call = False):
    pyautogui.click(x, y)
    check_edit_button()
    if first_call:  # If this is the first call, switch the input method
        pyautogui.hotkey('ctrl', 'space')
    if type_i:  # If type_i is True, type 'i'
        pyautogui.typewrite('i')
    pyautogui.typewrite(str(text))
    time.sleep(0.5)


def print_and_save_label(label_copies, printer='A'):
    """
    Print and save a label.

    Args:
    label_copies: The number of copies to print.
    printer: The printer to select. Options are 'A' or 'D'. Default is 'A'.
    """
    # Printer coordinates
    printer_0_coord = (389, 337)
    printer_A_coord = (415, 366)
    printer_D_coord = (407, 464)

    # Select the printer coordinates based on the printer argument
    if printer == 'A':
        printer_coord = printer_A_coord
    elif printer == 'D':
        printer_coord = printer_D_coord
    elif printer == '0':
        printer_coord = printer_0_coord
    else:
        raise ValueError(f"Invalid printer: {printer}. Options are 'A' or 'D'.")

    pyautogui.hotkey('ctrl', 'p')  # Press print
    check_print_button()
    pyautogui.click(570, 225)  # Select printer
    time.sleep(0.5)
    pyautogui.moveTo(530, 490)
    pyautogui.scroll(-100)  # Scroll down 100 units
    pyautogui.click(*printer_coord)  # Select printer A or D
    time.sleep(0.5)
    pyautogui.click(618, 494)  # Change the number of copies
    pyautogui.typewrite(str(label_copies))
    pyautogui.press('enter')  # Press enter
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 's')  # Press Ctrl + S to save
    pyautogui.hotkey('alt', 'f4')  # Press Alt + F4 to close NiceLabel Designer

