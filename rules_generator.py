def generate_python_rules_file(converter, converted_formulas, shared_data):
    """Generate complete Python file with all rules."""
    header = '''# Auto-generated Python rule functions
# Generated from Excel formulas

# Import Excel-like helper functions
from excel_functions import *

# Shared data (pre-calculated values)
shared_data = {}

# Generated rule functions
'''
    
    # Generate all rule functions
    rule_functions = []
    for converted in converted_formulas:
        func_code = converter.generate_python_function(converted)
        rule_functions.append(func_code)
    
    # Combine all parts
    complete_code = header + "\n\n".join(rule_functions)
    
    return complete_code