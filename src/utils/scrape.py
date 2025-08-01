import os
from openpyxl import load_workbook

def extract_data_and_formulas_from_excel(file_path):
    """Extracts formulas and their cell references from an Excel file."""
    workbook = load_workbook(filename=file_path, data_only=True)  # Use data_only=True to get values
    formulas_workbook = load_workbook(filename=file_path, data_only=False)

    extracted_data = {}

    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        formulas_ws = formulas_workbook[sheet_name]
        
        formulas_in_sheet = []
        data_in_sheet = {}

        for row in ws.iter_rows():
            for cell in row:
                data_in_sheet[cell.coordinate] = cell.value

        for row in formulas_ws.iter_rows():
            for cell in row:
                if cell.data_type == 'f' and cell.value:
                    formulas_in_sheet.append({
                        "cell": cell.coordinate,
                        "formula": cell.value
                    })
        
        extracted_data[sheet_name] = {
            "formulas": formulas_in_sheet,
            "data": data_in_sheet
        }

    return extracted_data