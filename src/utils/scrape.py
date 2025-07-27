import os
from openpyxl import load_workbook

def extract_formulas_from_excel(file_path):
    """Extracts formulas and their cell references from an Excel file."""
    workbook = load_workbook(filename=file_path, data_only=False)
    extracted_data = {}

    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        formulas_in_sheet = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == 'f' and cell.value:
                    formulas_in_sheet.append({
                        "cell": cell.coordinate,
                        "formula": cell.value
                    })
        extracted_data[sheet_name] = formulas_in_sheet

    return extracted_data