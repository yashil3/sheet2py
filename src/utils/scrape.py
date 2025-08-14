import os
from openpyxl import load_workbook

def extract_data_and_formulas_from_excel(file_path):
    """Extracts formulas, cell data, and key-value mappings from an Excel file.
    Detects sheets with a Key/Value table (two columns labeled 'Key' and 'Value')
    and returns:
      - formulas: list of {cell, formula}
      - data: dict of cell_reference -> value
      - key_values: dict of key -> value
      - cell_to_key: dict of value_cell_reference -> key (e.g., 'B12' -> 'Engine')
    """
    workbook = load_workbook(filename=file_path, data_only=True)  # Use data_only=True to get values
    formulas_workbook = load_workbook(filename=file_path, data_only=False)

    extracted_data = {}

    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        formulas_ws = formulas_workbook[sheet_name]

        formulas_in_sheet = []
        data_in_sheet = {}
        key_values = {}
        cell_to_key = {}

        # Collect raw cell values
        for row in ws.iter_rows():
            for cell in row:
                data_in_sheet[cell.coordinate] = cell.value

        # Collect formulas (use non-data_only workbook)
        for row in formulas_ws.iter_rows():
            for cell in row:
                if cell.data_type == 'f' and cell.value:
                    formulas_in_sheet.append({
                        "cell": cell.coordinate,
                        "formula": cell.value
                    })

        # Try to detect a Key/Value header row within the first few rows
        # We scan the first 10 rows for any pair of columns that are 'Key' and 'Value'
        header_found = False
        header_row_idx = None
        key_col_letter = None
        value_col_letter = None

        max_scan_rows = min(10, ws.max_row)
        for r in range(1, max_scan_rows + 1):
            row_values = {cell.column_letter: (cell.value.strip() if isinstance(cell.value, str) else cell.value)
                          for cell in ws[r] if cell.value is not None}
            # Find columns labeled 'Key' and 'Value' (case-insensitive)
            key_cols = [col for col, val in row_values.items() if isinstance(val, str) and val.lower() == 'key']
            value_cols = [col for col, val in row_values.items() if isinstance(val, str) and val.lower() == 'value']
            if key_cols and value_cols:
                key_col_letter = key_cols[0]
                value_col_letter = value_cols[0]
                header_row_idx = r
                header_found = True
                break

        # If key/value layout found, build the mappings
        if header_found and key_col_letter and value_col_letter and header_row_idx:
            for r in range(header_row_idx + 1, ws.max_row + 1):
                key_cell = f"{key_col_letter}{r}"
                val_cell = f"{value_col_letter}{r}"
                key = ws[key_cell].value
                val = ws[val_cell].value
                if key is None:
                    continue
                if isinstance(key, str):
                    key = key.strip()
                # Skip empty keys
                if key == "":
                    continue
                key_values[key] = val
                cell_to_key[val_cell] = key

        extracted_data[sheet_name] = {
            "formulas": formulas_in_sheet,
            "data": data_in_sheet,
            "key_values": key_values,
            "cell_to_key": cell_to_key,
        }

    return extracted_data