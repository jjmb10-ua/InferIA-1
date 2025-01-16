import pyautogui
import os
import subprocess
import time
import sys
import openpyxl

# Open a file, take screenshot and save the image
def takeScreenshot(filename, output_dir="images"):
    # Verify if the file exists
    if not os.path.exists(filename):
        print(f"File does not exists: {filename}")
        return

    # Create the directory
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Load the Excel file to obtain the sheets
    try:
        wb = openpyxl.load_workbook(filename, read_only=True)
        sheets = wb.sheetnames
        print(f"Sheets founds: {sheets}")
    except Exception as e:
        print(f"Error processing sheets: {e}")
        return

    # Open the file
    try:
        print(f"Opening the file: {filename}. Waiting for Excel to fully load...")
        subprocess.Popen(['start', filename], shell=True)
        time.sleep(5)  # Wait for the file to open completely
    except Exception as e:
        print(f"Error opening the file: {e}")
        return

    # Put Excel in full screen
    try:
        pyautogui.hotkey("alt", "space")  # Open the window menu
        time.sleep(1)
        pyautogui.press("x")  # Maximize de window
        time.sleep(1)
        pyautogui.hotkey("ctrl", "f1")  # Hide options
        time.sleep(2)
    except Exception as e:
        print(f"Error maximizing excel: {e}")
        return

    # Take the screenshot of each sheet
    for i, sh in enumerate(sheets):
        try:
            # Change the sheet using (Ctrl+PgDown o Ctrl+PgUp)
            if i > 0:
                pyautogui.hotkey("ctrl", "pgdn")  # Navigate to the next sheet
                time.sleep(2)  # Wait for the sheet to load

            # Take the screenshot
            screenshot = pyautogui.screenshot()
            filename = f"sh{i + 1}_{sh}.png"
            filepath = os.path.join(output_dir, filename)
            screenshot.save(filepath)
            print(f"Screenshot saved in: {filepath}")

        except Exception as e:
            print(f"Error taking the screenshot for the sheet {sh}: {e}")


    # Close the file
    try:
        print("Closing the file...")
        pyautogui.hotkey("alt", "f4")
        print("File closed correctly.")
    except Exception as e:
        print(f"Error closing the file: {e}")
    finally:
        wb.close()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py <ruta_del_archivo>")
        sys.exit(1)

    filename = sys.argv[1]
    
    # Capture the entire screen after opening the file
    takeScreenshot(filename)