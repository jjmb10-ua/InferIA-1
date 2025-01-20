#Author: Juan José Martínez Berná
import os
import sys
import re
import json
import pandas as pd
from openpyxl import Workbook, load_workbook
import pyautogui
import subprocess
import time
import openai
from pydantic import BaseModel
from typing import Optional
import base64

# Configurar la API Key
client = openai.Client(api_key="")

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
        wb = load_workbook(filename, read_only=True)
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

# Función para crear una clase dinámica de Pydantic
def createDynamicClass(class_name: str, attributes: list):
    # Usamos `__annotations__` para definir los tipos correctamente
    fields = {attr: Optional[str] for attr in attributes}  # Solo anotación de tipo
    defaults = {attr: None for attr in attributes}  # Valores predeterminados
    
    # Crear la clase con `type`
    dynamic_class = type(class_name, (BaseModel,), {"__annotations__": fields, **defaults})
    return dynamic_class

# Function to encode the image
def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

# Convierte la respuesta en diccionario y la guarda como JSON
def save_dict_to_json(response_model, outputFile):
    try:
        # Convertir la respuesta del modelo Pydantic a un diccionario
        response_dict = response_model.model_dump()

        # Guardar el diccionario en un archivo JSON
        with open(outputFile, "w", encoding="utf-8") as json_file:
            json.dump(response_dict, json_file, indent=4, ensure_ascii=False)

        print(f"Datos guardados en {outputFile}")
    except Exception as e:
        print(f"Error al guardar los datos en JSON: {e}")

def processMetadata(sh_index, sh, attributes):
    # Create dynamic class with specified attributes
    DynamicModel = createDynamicClass("DynamicModel", attributes)

    # Path to your image
    image_path = f"images/sh{sh_index}_{sh.title}.png"
    # Getting the base64 string
    base64_image = encode_image(image_path)

    completion = client.beta.chat.completions.parse(
        model="gpt-4o-2024-08-06",
        messages=[
            {"role": "system", "content": "You are an expert at structured data extraction. You will be given an image from an Excel sheet and should convert it into the given structure."},
            {"role": "user", "content": [
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{base64_image}"},
                    },
                ],
            }
        ],
        response_format=DynamicModel,
    )
    research_paper = completion.choices[0].message.parsed
    print("\nResponse from the API:")
    print(research_paper)
    outputFile = f"sh{sh_index}_{sh.title}_metada.json"
    save_dict_to_json(research_paper, outputFile)

def isValidHeader(row_cleaned):
    # Comprobar si cada celda parece un encabezado válido
    return all(cell.isalpha() or cell.replace(" ", "").isalnum() for cell in row_cleaned)

def getMissingHeader(fname):
    # Remove the file extension
    base_name = os.path.splitext(fname)[0]
    # Remove numbers and non-alphabetic characters
    words = re.findall(r'[a-zA-Z]+', base_name)
    # Capitalize the words
    cleaned_title = " ".join(word.capitalize() for word in words)

    return cleaned_title

def processData(fname, sh_index, sh, headers):
    data = []
    headers = None

    # Iterate through all rows in the sheet and adds an index
    for row_index, row in enumerate(sh.iter_rows(values_only=True), start=1):
        # Clean up the row: filter out None values and convert to strings
        row_cleaned = [str(cell) for cell in row if cell is not None]

        if not table_begin and len(row_cleaned) > 1:
            # Trying to access to the next row
            try:
                next_row = next(sh.iter_rows(min_row=row_index+1, max_row=row_index+1, values_only=True))
                next_row_cleaned = [str(cell) for cell in next_row if cell is not None]

                # Check if the next row has the same number of cells or one more
                if len(next_row_cleaned) == len(row_cleaned) and isValidHeader(row_cleaned):
                    headers = row_cleaned
                    table_begin = True
                    print(f"Table starts at row {row_index} with headers: {headers}")
                
                elif len(next_row_cleaned) == len(row_cleaned) + 1:
                    # We have to search the missing header on the tittle or metadata
                    missing_header = getMissingHeader(fname)
                    headers = [missing_header] + row_cleaned
                    table_begin = True
                    print(f"Table starts at row {row_index} with headers: {headers} (missing header added)")
                else:
                    print(f"Skipping row {row_index}: does not match the expected structure for headers.")
            except StopIteration:
                print(f"Row {row_index} is the last row or next row does not exists.")
        
        else:
            # Append remaining rows as data
            data.append(row)

        # Create filenames for the sheet
        sheet_name = sh.title.replace(" ", "_")
        data_filename = f"{sheet_name}_data.csv"

        # Save data as CSV if headers and data exist
        if headers and data:
            # Ensure all rows have the same length as headers
            data = [row[:len(headers)] for row in data]

            df = pd.DataFrame(data, columns=headers)
            df.to_csv(data_filename, index=False, encoding="utf-8")
            print(f"Data saved to {data_filename}")
        else:
            print(f"No table data found in sheet '{sh.title}'")

# Processes each sheet and distinguish between metadata and data
def processSheet(fname, sh_index, sh, attributes):
    # Separate metadata and data
    processMetadata(sh_index, sh, attributes)

    # processData(fname, sh_index, sh)

# Try to open the file and iterate all sheets
def openFile(filename, attributes):
    try:
        wb = load_workbook(filename, read_only=True)

        for sh_index, sheet in enumerate(wb, start=1):
            processSheet(filename, sh_index, sheet, attributes)

    except FileNotFoundError:
        print(f"Error: File '{filename}' not found.")
    except Exception as e:
        print(f"Oops... error : {e}")

def main():
    if len(sys.argv) < 3:
        print("Usage: python .\Code.py <EXCEL_SHEET> <ATTRIBUTES...>")
        sys.exit(1)

    # First argument is Excel sheet
    excel_path = sys.argv[1]
    # The following arguments are the metadata attributes
    attributes = sys.argv[2:]

    # Capture the entire screen from all sheets after opening the file
    takeScreenshot(excel_path)

    openFile(excel_path, attributes)

######################### MAIN ##########################

if __name__ == "__main__":
    main()