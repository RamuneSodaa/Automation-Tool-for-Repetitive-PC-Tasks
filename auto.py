import pyautogui
import time
import xlrd
import pyperclip

# Define mouse events

def mouse_click(click_times, button, img, retry):
    """Handles mouse click events."""
    while True:
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=click_times, interval=0.2, duration=0.2, button=button)
            if retry != -1:
                break
        else:
            print(f"Image not found. Retrying in 0.1 seconds: {img}")
            time.sleep(0.1)
        
        if retry > 1:
            retry -= 1
            if retry == 0:
                break

def data_check(sheet1):
    """Validates the data from the sheet for proper command structure."""
    if sheet1.nrows < 2:
        print("No data found.")
        return False
    
    for i in range(1, sheet1.nrows):
        cmd_type = sheet1.row(i)[0]
        cmd_value = sheet1.row(i)[1]
        
        # Check for valid command types (1 to 6)
        if cmd_type.ctype != 2 or cmd_type.value not in [1.0, 2.0, 3.0, 4.0, 5.0, 6.0]:
            print(f"Row {i + 1}: Invalid command type in column 1.")
            return False

        # Check for valid content based on command type
        if cmd_type.value in [1.0, 2.0, 3.0] and cmd_value.ctype != 1:  # Click actions must have a string (image file)
            print(f"Row {i + 1}: Invalid content in column 2.")
            return False
        if cmd_type.value == 4.0 and cmd_value.ctype == 0:  # Input type must have non-empty content
            print(f"Row {i + 1}: Input value cannot be empty.")
            return False
        if cmd_type.value in [5.0, 6.0] and cmd_value.ctype != 2:  # Wait and scroll must have numeric values
            print(f"Row {i + 1}: Invalid content in column 2.")
            return False

    return True

def execute_commands(sheet1):
    """Executes commands based on the content of the sheet."""
    for i in range(1, sheet1.nrows):
        cmd_type = sheet1.row(i)[0]
        cmd_value = sheet1.row(i)[1].value

        if cmd_type.value == 1.0:  # Single left click
            retries = sheet1.row(i)[2].value if sheet1.row(i)[2].ctype == 2 else 1
            mouse_click(1, "left", cmd_value, retries)
            print(f"Single left click on: {cmd_value}")

        elif cmd_type.value == 2.0:  # Double left click
            retries = sheet1.row(i)[2].value if sheet1.row(i)[2].ctype == 2 else 1
            mouse_click(2, "left", cmd_value, retries)
            print(f"Double left click on: {cmd_value}")

        elif cmd_type.value == 3.0:  # Right click
            retries = sheet1.row(i)[2].value if sheet1.row(i)[2].ctype == 2 else 1
            mouse_click(1, "right", cmd_value, retries)
            print(f"Right click on: {cmd_value}")

        elif cmd_type.value == 4.0:  # Input text
            pyperclip.copy(cmd_value)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print(f"Entered text: {cmd_value}")

        elif cmd_type.value == 5.0:  # Wait
            wait_time = cmd_value
            time.sleep(wait_time)
            print(f"Waited for {wait_time} seconds")

        elif cmd_type.value == 6.0:  # Scroll
            scroll_distance = int(cmd_value)
            pyautogui.scroll(scroll_distance)
            print(f"Scrolled by {scroll_distance} units")

if __name__ == '__main__':
    file = 'cmd.xls'
    # Open the Excel file
    wb = xlrd.open_workbook(filename=file)
    # Get the first sheet by index
    sheet1 = wb.sheet_by_index(0)
    print('RPA Automation Starting...')

    # Validate the data
    if data_check(sheet1):
        option = input('Choose option: 1. Execute once  2. Loop execution \n')
        
        if option == '1':
            execute_commands(sheet1)
        elif option == '2':
            while True:
                execute_commands(sheet1)
                time.sleep(0.1)
                print("Waiting for 0.1 seconds between iterations.")
    else:
        print("Invalid input or program has exited.")
