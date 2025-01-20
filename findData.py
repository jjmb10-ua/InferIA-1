from openpyxl import load_workbook

# Finds the starting line of the data based on the most frequent number of non-empty columns
def findDataStart(sheet, expected_headers):
    # Minimum number of expected columns (based on the provided headers)
    min_columns = len(expected_headers)

    # Dictionary to store statistics
    column_stats = {}

    # Normalize expected headers to lowercase
    normalized_headers = [header.lower() for header in expected_headers]

    # Iterate over the rows in the sheet
    for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        # Count the number of non-empty cells in the row
        non_empty_cells = [cell for cell in row if cell is not None]

        # Normalize the row cells to lowercase
        normalized_row = [str(cell).lower() for cell in non_empty_cells]

        # Check if the row matches the expected headers, to skip it
        if set(normalized_row) == set(normalized_headers):
            # Skip this row because it appears to be a header
            continue

        # Only record rows with at least the minimum number of expected columns
        if len(non_empty_cells) >= min_columns:
            if len(non_empty_cells) not in column_stats:
                # If this number of columns is not in the dictionary, initialize it
                column_stats[len(non_empty_cells)] = {
                    "first_row": row_idx,  # Store the first row where it appears
                    "count": 0  # Initial counter
                }
            # Increment the counter
            column_stats[len(non_empty_cells)]["count"] += 1

    # If no rows were found, return -1
    if not column_stats:
        return -1

    # Find the number of columns with the highest occurrences
    most_frequent = max(column_stats.items(), key=lambda x: x[1]["count"])
    most_frequent_row = most_frequent[1]["first_row"]

    # print(f"Column statistics: {column_stats}")
    return most_frequent_row


if __name__ == "__main__":
    import sys

    # Collect the Excel file path and headers from command-line arguments
    if len(sys.argv) < 3:
        print("Usage: python script.py <excel_file_path> <header1> <header2> ...")
        sys.exit(1)

    # Excel file path
    excel_path = sys.argv[1]

    # List of expected headers
    expected_headers = sys.argv[2:]  # Headers as a list

    try:
        # Load the Excel file
        wb = load_workbook(excel_path, read_only=True)
        sheet = wb.active  # Use the first sheet by default

        # Find the most representative row
        data_start_row = findDataStart(sheet, expected_headers)

        if data_start_row == -1:
            print("No rows with data were found.")
        else:
            print(f"Data starts on row: {data_start_row}")

    except FileNotFoundError:
        print(f"Error: The file '{excel_path}' was not found.")
    except Exception as e:
        print(f"Error while processing the file: {e}")
