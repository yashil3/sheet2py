"""
Excel-like helper functions for formula evaluation.
These functions provide Python implementations of Excel operations.
"""

import math
import re
from datetime import datetime, timedelta
from calendar import monthrange

def get_cell(data, sheet, cell):
    """Get cell value from data structure"""
    try:
        return data[sheet][cell]
    except KeyError:
        return 0

def sum_range(data, sheet, start_cell, end_cell):
    """Sum a range of cells"""
    values = []
    start_col, start_row = parse_cell_ref(start_cell)
    end_col, end_row = parse_cell_ref(end_cell)
    
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell_ref = f"{chr(ord('A') + col - 1)}{row}"
            try:
                value = data[sheet][cell_ref]
                if isinstance(value, (int, float)):
                    values.append(value)
            except KeyError:
                continue
    return sum(values)

def count_range(data, sheet, start_cell, end_cell):
    """Count non-empty cells in a range"""
    count = 0
    start_col, start_row = parse_cell_ref(start_cell)
    end_col, end_row = parse_cell_ref(end_cell)
    
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell_ref = f"{chr(ord('A') + col - 1)}{row}"
            try:
                value = data[sheet][cell_ref]
                if value is not None and value != "":
                    count += 1
            except KeyError:
                continue
    return count

def average_range(data, sheet, start_cell, end_cell):
    """Calculate average of a range"""
    values = []
    start_col, start_row = parse_cell_ref(start_cell)
    end_col, end_row = parse_cell_ref(end_cell)
    
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell_ref = f"{chr(ord('A') + col - 1)}{row}"
            try:
                value = data[sheet][cell_ref]
                if isinstance(value, (int, float)):
                    values.append(value)
            except KeyError:
                continue
    return sum(values) / len(values) if values else 0

def count_if_range(data, sheet, start_cell, end_cell, criteria):
    """Count cells meeting criteria"""
    count = 0
    start_col, start_row = parse_cell_ref(start_cell)
    end_col, end_row = parse_cell_ref(end_cell)
    
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell_ref = f"{chr(ord('A') + col - 1)}{row}"
            try:
                value = data[sheet][cell_ref]
                if evaluate_criteria(value, criteria):
                    count += 1
            except KeyError:
                continue
    return count

def count_if(data, sheet, start_cell, end_cell, criteria):
    """Alias for count_if_range for backward compatibility"""
    return count_if_range(data, sheet, start_cell, end_cell, criteria)

def sum_if(data, sheet, start_cell, end_cell, criteria):
    """Sum cells meeting criteria"""
    total = 0
    start_col, start_row = parse_cell_ref(start_cell)
    end_col, end_row = parse_cell_ref(end_cell)
    
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell_ref = f"{chr(ord('A') + col - 1)}{row}"
            try:
                value = data[sheet][cell_ref]
                if evaluate_criteria(value, criteria):
                    total += value if isinstance(value, (int, float)) else 0
            except KeyError:
                continue
    return total

def safe_execute(func, error_value):
    """Execute function safely, return error_value on exception"""
    try:
        return func()
    except:
        return error_value

def is_error(func):
    """Check if function execution results in error"""
    try:
        func()
        return False
    except:
        return True

def concat(*args):
    """Concatenate strings"""
    return ''.join(str(arg) for arg in args)

def round_down(value, digits):
    """Round down to specified digits"""
    multiplier = 10 ** digits
    return math.floor(value * multiplier) / multiplier

def rows_count(start_cell, end_cell):
    """Count rows in range"""
    _, start_row = parse_cell_ref(start_cell)
    _, end_row = parse_cell_ref(end_cell)
    return end_row - start_row + 1

def find_text(search_text, search_in):
    """Find text position (1-based)"""
    try:
        return str(search_in).find(str(search_text)) + 1
    except:
        return 0

# Complex functions that should be identified for user replacement
def vlookup(lookup_value, data, sheet, range_text, col_index, exact_match=True):
    """VLOOKUP - complex function for user replacement"""
    return f"VLOOKUP({lookup_value}, {sheet}!{range_text}, {col_index}, {exact_match})"

def index(data, sheet, range_text, row, col=1):
    """INDEX - complex function for user replacement"""
    return f"INDEX({sheet}!{range_text}, {row}, {col})"

def indirect(data, ref):
    """INDIRECT - complex function for user replacement"""
    return f"INDIRECT({ref})"

def countifs(data, ranges_criteria):
    """COUNTIFS - complex function for user replacement"""
    return f"COUNTIFS({ranges_criteria})"

def eomonth(start_date, months):
    """EOMONTH - complex function for user replacement"""
    return f"EOMONTH({start_date}, {months})"

# Helper functions
def parse_cell_ref(cell_ref):
    """Parse cell reference like 'A1' into column number and row number"""
    match = re.match(r'([A-Z]+)(\d+)', cell_ref)
    if match:
        col_str, row_str = match.groups()
        col_num = 0
        for char in col_str:
            col_num = col_num * 26 + (ord(char) - ord('A') + 1)
        return col_num, int(row_str)
    return 1, 1

def evaluate_criteria(value, criteria):
    """Evaluate criteria string like '>10' against value"""
    criteria = criteria.strip('"')
    if criteria.startswith('>='):
        return value >= float(criteria[2:])
    elif criteria.startswith('<='):
        return value <= float(criteria[2:])
    elif criteria.startswith('>'):
        return value > float(criteria[1:])
    elif criteria.startswith('<'):
        return value < float(criteria[1:])
    elif criteria.startswith('='):
        return value == criteria[1:]
    else:
        return value == criteria