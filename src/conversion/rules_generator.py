def generate_python_rules_file(converter, converted_formulas, shared_data, sorted_cells):
    """Generate complete Python file with all rules, respecting topological order."""
    header = '''# Auto-generated Python rule functions
# Generated from Excel formulas

# Import Excel-like helper functions
from src.conversion.excel_functions import *

# Shared data (pre-calculated values)
shared_data = {}

# Generated rule functions
'''
    
    # Generate rule functions in topological order
    rule_functions = []
    for cell_ref in sorted_cells:
        # Find the ConvertedFormula object that matches the cell_ref
        formula = next((f for f in converted_formulas if (f"{f.sheet}!{f.cell_reference}" if f.sheet else f.cell_reference) == cell_ref), None)
        if formula:
            func_code = converter.generate_python_function(formula)
            rule_functions.append(func_code)
        else:
            print(f"Warning: Could not find formula for cell {cell_ref}")
    
    # Combine all parts
    complete_code = header + "\n\n".join(rule_functions)
    
    return complete_code