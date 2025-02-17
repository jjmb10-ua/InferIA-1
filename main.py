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
from typing import Optional, List
import base64
import Code
import findData
import metaData

# Processes each sheet and distinguish between metadata and data
def processSheet(fname, sh_index, sh, attributes, previous_headers):

    data_start = findData.findDataStart(sh)

    # Separate metadata and data
    metaData.processMetadata(fname, sh_index, sh, attributes, data_start)
    # headers = metaData.processHeaders(sh_index, sh, attributes, data_start)
    headers = metaData.findHeaders(sh_index, sh, previous_headers=previous_headers)
    findData.processData(fname, sh_index, sh, headers, data_start)

    previous_headers.extend(headers)

# Try to open the file and iterate all sheets
def openFile(filename, attributes):
    try:
        wb = load_workbook(filename, read_only=True)

        accumulated_headers = []

        for sh_index, sheet in enumerate(wb, start=1):
            processSheet(filename, sh_index, sheet, attributes, accumulated_headers)

    except FileNotFoundError:
        print(f"Error: File '{filename}' not found.")
    except Exception as e:
        print(f"Oops... error : {e}")


def main():
    if len(sys.argv) < 3:
        print("Usage: python Code.py <EXCEL_SHEET> <ATTRIBUTES...>")
        sys.exit(1)

    # First argument is Excel sheet
    excel_path = sys.argv[1]
    # The following arguments are the metadata attributes
    attributes = sys.argv[2:]

    # Capture the entire screen from all sheets after opening the file
    metaData.takeScreenshot(excel_path)

    openFile(excel_path, attributes)

######################### MAIN ##########################

if __name__ == "__main__":
    main()