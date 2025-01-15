#Author: Juan José Martínez Berná
import os
import sys
import re
import json
import pandas as pd
from openpyxl import Workbook, load_workbook

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

def processMetadata(row_cleaned):
    # If there is only one cell we look for key-value in it
    if len(row_cleaned) == 1:
        text = row_cleaned[0]

        # Divide by ":" or "=" if there are
        if ":" in text:
            parts = text.split(":", 1)
            key = parts[0].strip()
            value = parts[1].strip()
        elif "=" in text:
            parts = text.split("=", 1)
            key = parts[0].strip()
            value = parts[1].strip()
        # If there is no specific separator, assume key-value separated by spaces
        else:
            parts = text.split(" ", 1)
            key = parts[0].strip()
            # if there is only one word, there is no more information
            value = " ".join(parts[1:]).strip() or "No information"
    
    # With more than one cell, the first is the key and the rest are values
    elif len(row_cleaned) >= 2:
        key = row_cleaned[0].strip()
        # Join the elements of the resulting list into a single string, separating them with a space
        value = " ".join(row_cleaned[1:]).strip()

    else:
        key, value = "Desconocido", ""

    # Normalizar el formato
    key = key.capitalize()
    return key, value

# Processes each sheet and distinguish between metadata and data
def processSheet(fname,sh):
    # Separate metadata and data
    metadata = {}
    table_begin = False
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
        
        elif table_begin:
            # Append remaining rows as data
            data.append(row)
        else:
            # Add row to metadata
            key, value = processMetadata(row_cleaned)
            metadata[key] = value

    # Create filenames for the sheet
    sheet_name = sh.title.replace(" ", "_")
    metadata_filename = f"{sheet_name}_metadata.json"
    data_filename = f"{sheet_name}_data.csv"

    # Save metadata as JSON
    metadata_dict = {
        "sheet_name": sh.title,
        "metadata_rows": len(metadata),
        "metadata": metadata
    }
    with open(metadata_filename, "w", encoding="utf-8") as json_file:
        json.dump(metadata_dict, json_file, ensure_ascii=False, indent=4)
    print(f"Metadata saved to {metadata_filename}")

    # Save data as CSV if headers and data exist
    if headers and data:
        # Ensure all rows have the same length as headers
        data = [row[:len(headers)] for row in data]

        df = pd.DataFrame(data, columns=headers)
        df.to_csv(data_filename, index=False, encoding="utf-8")
        print(f"Data saved to {data_filename}")
    else:
        print(f"No table data found in sheet '{sh.title}'")

# Try to open the file and iterate all sheets
def openFile(filename):
    try:
        wb = load_workbook(filename, read_only=True)

        for sheet in wb:
            processSheet(filename,sheet)

    except FileNotFoundError:
        print(f"Error: File '{filename}' not found.")
    except Exception as e:
        print(f"Oops... error : {e}")

def main():
    if len(sys.argv) == 2:
        filename = sys.argv[1]
        openFile(filename)
    else:
        print("Usage: python Code.py <file.xlsx>")
        sys.exit(1)

######################### MAIN ##########################

if __name__ == "__main__":
    main()