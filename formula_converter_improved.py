import re
import json
from typing import Dict, Any, List, Optional
from dataclasses import dataclass
import time
import logging
from tqdm import tqdm  # Import tqdm for progress bar

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

@dataclass
class ConvertedFormula:
    """Represents a converted formula with metadata."""
    original_formula: str
    python_expression: str
    cell_reference: str
    sheet: str
    dependencies: List[str]
    description: str
    rule_type: str

class ExcelToPythonConverter:
    """Fixed converter for Excel formulas to Python expressions."""
    
    def __init__(self):
        self.converted_cells = set()
        
    def convert_cell_references(self, text: str) -> str:
        """Convert cell references to Python data access."""
        logging.info(f"Converting cell references in: {text}")
        # Pattern for cell references like "Source!B14" or "B14"
        # Use a negative lookbehind to avoid matching function names like 'find_text'
        cell_pattern = r'(?<![a-zA-Z0-9_])([A-Za-z_]+!)?\$?([A-Z]+)\$?(\d+)'
        
        def replace_cell(match):
            sheet_part = match.group(1)
            col = match.group(2)
            row = match.group(3)
            
            if sheet_part:
                sheet_name = sheet_part.rstrip('!')
                return f"get_cell(data, '{sheet_name}', '{col}{row}')"
            else:
                return f"get_cell(data, 'Formulas', '{col}{row}')"
        
        result = re.sub(cell_pattern, replace_cell, text)
        logging.info(f"Cell references conversion result: {result}")
        return result
    
    def convert_sum_function(self, formula: str) -> str:
        """Convert Excel SUM functions to Python sum_range()."""
        logging.info(f"Converting SUM function in: {formula}")
        # Handle both sheet-qualified and unqualified ranges
        sum_pattern = r'SUM\s*\(\s*(?:([A-Za-z_]+)!)?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\s*\)'
        
        def replace_sum(match):
            sheet_name = match.group(1)
            start_col = match.group(2)
            start_row = match.group(3)
            end_col = match.group(4)
            end_row = match.group(5)
            
            if sheet_name:
                return f"sum_range(data, '{sheet_name}', '{start_col}{start_row}', '{end_col}{end_row}')"
            else:
                return f"sum_range(data, 'Formulas', '{start_col}{start_row}', '{end_col}{end_row}')"
        
        result = re.sub(sum_pattern, replace_sum, formula)
        logging.info(f"SUM function conversion result: {result}")
        return result
    
    def convert_if_statement(self, formula: str) -> str:
        """Convert Excel IF statements to Python conditional expressions."""
        logging.info(f"Converting IF statement in: {formula}")
        
        max_iterations = 10
        iteration = 0
        
        while 'IF(' in formula and iteration < max_iterations:
            iteration += 1
            logging.info(f"IF statement iteration: {iteration}")
            
            # Find the outermost IF statement
            if_pattern = r'IF\s*\('
            match = re.search(if_pattern, formula)
            
            if not match:
                logging.info("No IF statement found, breaking loop.")
                break
            
            start_index = match.end()
            
            # Extract arguments
            args = []
            current_arg = ''
            paren_count = 0
            
            for i in range(start_index, len(formula)):
                char = formula[i]
                
                if char == '(':
                    paren_count += 1
                elif char == ')':
                    paren_count -= 1
                elif char == ',' and paren_count == 0:
                    args.append(current_arg.strip())
                    current_arg = ''
                elif char == ')':
                    if paren_count == -1:
                        args.append(current_arg.strip())
                        end_index = i
                        break
                    else:
                        current_arg += char
                else:
                    current_arg += char
            
            if len(args) == 3:
                condition = args[0].strip()
                true_val = args[1].strip()
                false_val = args[2].strip()
                
                # Replace the IF statement with the Python equivalent
                python_if = f"({true_val} if {condition} else {false_val})"
                formula = formula[:match.start()] + python_if + formula[end_index + 1:]
                logging.info(f"Replaced IF statement with: {python_if}")
            else:
                logging.error(f"Incorrect number of arguments in IF statement: {args}")
                break
        
        logging.info(f"IF statement conversion result: {formula}")
        return formula
    
    def convert_operators(self, text: str) -> str:
        """Convert Excel operators to Python operators."""
        logging.info(f"Converting operators in: {text}")
        # Handle comparison operators
        text = text.replace('<>', '!=')
        
        # Replace = with == for comparisons, but be careful with existing == or other operators
        text = re.sub(r'(?<![<>=!])=(?!=)', '==', text)
        
        logging.info(f"Operators conversion result: {text}")
        return text
    
    def convert_or_function(self, formula: str) -> str:
        """Convert Excel OR function to Python or operator."""
        logging.info(f"Converting OR function in: {formula}")
        or_pattern = r'OR\s*\(\s*([^)]+)\s*\)'
        
        def replace_or(match):
            args_str = match.group(1)
            args = self.split_function_args(args_str)
            
            python_conditions = []
            for arg in args:
                arg = arg.strip()
                arg = self.convert_operators(arg)
                python_conditions.append(f"({arg})")
            
            return f"({' or '.join(python_conditions)})"
        
        result = re.sub(or_pattern, replace_or, formula)
        logging.info(f"OR function conversion result: {result}")
        return result
    
    def convert_and_function(self, formula: str) -> str:
        """Convert Excel AND function to Python and operator."""
        logging.info(f"Converting AND function in: {formula}")
        and_pattern = r'AND\s*\(\s*([^)]+)\s*\)'
        
        def replace_and(match):
            args_str = match.group(1)
            args = self.split_function_args(args_str)
            
            python_conditions = []
            for arg in args:
                arg = arg.strip()
                arg = self.convert_operators(arg)
                python_conditions.append(f"({arg})")
            
            return f"({' and '.join(python_conditions)})"
        
        result = re.sub(and_pattern, replace_and, formula)
        logging.info(f"AND function conversion result: {result}")
        return result
    
    def convert_find_function(self, formula: str) -> str:
        """Convert Excel FIND function to Python find_text."""
        logging.info(f"Converting FIND function in: {formula}")
        find_pattern = r'FIND\s*\(\s*([^,]+?)\s*,\s*([^)]+?)\s*\)\s*>\s*0'
        
        def replace_find(match):
            search_text = match.group(1).strip()
            search_in = match.group(2).strip()
            return f"find_text({search_text}, {search_in})"
        
        result = re.sub(find_pattern, replace_find, formula)
        logging.info(f"FIND function conversion result: {result}")
        return result
    
    def convert_countif_function(self, formula: str) -> str:
        """Convert Excel COUNTIF functions to Python equivalents."""
        logging.info(f"Converting COUNTIF function in: {formula}")
        countif_pattern = r'COUNTIF\s*\(\s*([^,]+?)\s*,\s*([^)]+?)\s*\)'
        
        def replace_countif(match):
            range_ref = match.group(1).strip()
            criteria = match.group(2).strip()
            
            if ':' in range_ref:
                # Handle range like "Source!B8:B10"
                parts = range_ref.split(':')
                start = parts[0].strip()
                end = parts[1].strip()
                
                # Extract sheet name from start
                if '!' in start:
                    sheet_name = start.split('!')[0]
                    start_cell = start.split('!')[1]
                    end_cell = end
                    return f"count_if_range(data, '{sheet_name}', '{start_cell}', '{end_cell}', {criteria})"
                else:
                    return f"count_if_range(data, 'Formulas', '{start}', '{end}', {criteria})"
            else:
                return f"(1 if {range_ref} == {criteria} else 0)"
        
        result = re.sub(countif_pattern, replace_countif, formula)
        logging.info(f"COUNTIF function conversion result: {result}")
        return result
    
    def convert_iferror_function(self, formula: str) -> str:
        """Convert Excel IFERROR to Python try-except."""
        logging.info(f"Converting IFERROR function in: {formula}")
        iferror_pattern = r'IFERROR\s*\(\s*([^,]+?)\s*,\s*([^)]+?)\s*\)'
        
        def replace_iferror(match):
            try_expr = match.group(1).strip()
            error_value = match.group(2).strip()
            return f"safe_execute(lambda: ({try_expr}), {error_value})"
        
        result = re.sub(iferror_pattern, replace_iferror, formula)
        logging.info(f"IFERROR function conversion result: {result}")
        return result
    
    def convert_rows_function(self, formula: str) -> str:
        """Convert Excel ROWS function."""
        logging.info(f"Converting ROWS function in: {formula}")
        rows_pattern = r'ROWS\s*\(\s*([^:]+):([^)]+)\s*\)'
        
        def replace_rows(match):
            start_ref = match.group(1).strip()
            end_ref = match.group(2).strip()
            
            # Clean up the references - remove sheet names for counting
            start_clean = start_ref.split('!')[-1] if '!' in start_ref else start_ref
            end_clean = end_ref.split('!')[-1] if '!' in end_ref else end_ref
            
            return f"rows_count('{start_clean}', '{end_clean}')"
        
        result = re.sub(rows_pattern, replace_rows, formula)
        logging.info(f"ROWS function conversion result: {result}")
        return result
    
    def split_function_args(self, args_str: str) -> List[str]:
        """Split function arguments, handling nested parentheses."""
        logging.info(f"Splitting function arguments: {args_str}")
        args = []
        current_arg = ""
        paren_count = 0
        
        for char in args_str:
            if char == '(':
                paren_count += 1
            elif char == ')':
                paren_count -= 1
            elif char == ',' and paren_count == 0:
                args.append(current_arg.strip())
                current_arg = ""
                continue
            
            current_arg += char
        
        if current_arg.strip():
            args.append(current_arg.strip())
        
        logging.info(f"Split arguments: {args}")
        return args
    
    def convert_expression(self, expr: str, shared_data: Dict[str, Any]) -> str:
        """Convert a complete Excel expression to Python."""
        logging.info(f"Starting conversion of expression: {expr}")
        # Remove leading = sign
        if expr.startswith('='):
            expr = expr[1:]
        
        # Handle string concatenation (&) before other conversions
        expr = expr.replace(' & ', ' + ')
        
        # Apply conversions in specific order
        expr = self.convert_rows_function(expr)
        expr = self.convert_iferror_function(expr)
        expr = self.convert_find_function(expr)
        expr = self.convert_countif_function(expr)
        expr = self.convert_sum_function(expr)
        
        # Convert logical functions before IF statements
        expr = self.convert_or_function(expr)
        expr = self.convert_and_function(expr)
        
        # Convert operators before IF statements to handle conditions correctly
        expr = self.convert_operators(expr)
        
        # Convert IF statements after operators
        expr = self.convert_if_statement(expr)

        # Convert cell references last
        expr = self.convert_cell_references(expr)
        
        logging.info(f"Final converted expression: {expr}")
        return expr
    
    def classify_formula(self, formula: str) -> str:
        """Classify the type of rule based on formula pattern."""
        if 'SUM(' in formula and '/' in formula:
            return "weight_calculation"
        elif 'IF(' in formula and any(op in formula for op in ['>=', '<=', '>', '<']):
            return "condition_check"
        elif 'COUNTIF(' in formula:
            return "count_matching"
        elif 'FIND(' in formula:
            return "text_search"
        elif 'SUM(' in formula and '*' in formula:
            return "weighted_sum"
        elif '"' in formula and 'IF(' in formula:
            return "scoring_rule"
        else:
            return "general_calculation"
    
    def generate_description(self, formula: str, rule_type: str) -> str:
        """Generate a human-readable description of the formula."""
        descriptions = {
            "weight_calculation": "Calculate relative weight as percentage of total",
            "condition_check": "Check if condition is met and return 1 or 0",
            "count_matching": "Count how many items match the criteria",
            "text_search": "Search for text within another text field",
            "weighted_sum": "Calculate weighted sum of multiple values",
            "scoring_rule": "Assign score based on value ranges",
            "general_calculation": "General calculation or transformation"
        }
        return descriptions.get(rule_type, "Unknown formula type")
    
    def analyze_formula(self, formula: str, cell: str, sheet: str, shared_data: Dict[str, Any]) -> ConvertedFormula:
        """Analyze and convert a single formula."""
        original = formula
        
        # Extract dependencies (cell references from original formula)
        dependencies = re.findall(r'([A-Za-z_]+!)?\$?([A-Z]+)\$?(\d+)', formula)
        dep_list = []
        for dep in dependencies:
            sheet_name = dep[0].rstrip('!') if dep[0] else None
            cell_ref = f"{dep[1]}{dep[2]}"
            if sheet_name:
                dep_list.append(f"'{sheet_name}'!{cell_ref}")
            else:
                dep_list.append(cell_ref)
        
        # Remove duplicates while preserving order
        seen = set()
        unique_deps = [x for x in dep_list if not (x in seen or seen.add(x))]

        python_expr = self.convert_expression(formula, shared_data)
        
        # Determine rule type based on formula pattern
        rule_type = self.classify_formula(formula)
        
        # Generate description
        description = self.generate_description(formula, rule_type)
        
        return ConvertedFormula(
            original_formula=original,
            python_expression=python_expr,
            cell_reference=cell,
            sheet=sheet,
            dependencies=unique_deps,
            description=description,
            rule_type=rule_type
        )
    
    def generate_python_function(self, converted: ConvertedFormula) -> str:
        """Generate a complete Python function from converted formula."""
        # Create function name from cell reference
        func_name = f"rule_{converted.sheet.lower()}_{converted.cell_reference.lower()}"
        
        # Create dependencies string
        deps_str = ", ".join(converted.dependencies)
        
        # Generate function code
        function_code = f'''def {func_name}(data, shared_data):
    """
    {converted.description}
    Original Excel formula: {converted.original_formula}
    Cell: {converted.sheet}!{converted.cell_reference}
    Rule type: {converted.rule_type}
    Dependencies: {deps_str}
    """
    try:
        result = {converted.python_expression}
        return result
    except Exception as e:
        print(f"Error in {func_name}: {{e}}")
        return None'''
        
        return function_code

def generate_python_rules_file(converter, converted_formulas, shared_data):
    """Generate complete Python file with all rules."""
    header = '''# Auto-generated Python rule functions
# Generated from Excel formulas

# Helper functions for Excel-like operations
import re

def get_cell(data, sheet, cell):
    """Get a cell value from the data structure."""
    try:
        if sheet in data:
            return data[sheet].get(cell, 0)
        return 0
    except:
        return 0

def sum_range(data, sheet, start_cell, end_cell):
    """Sum a range of cells like Excel B2:B9."""
    try:
        # Extract row numbers
        start_row = int(re.search(r'(\\d+)', start_cell).group(1))
        end_row = int(re.search(r'(\\d+)', end_cell).group(1))
        
        # Extract column (assume same column for range)
        col = re.search(r'([A-Z]+)', start_cell).group(1)
        
        total = 0
        for row in range(start_row, end_row + 1):
            cell_ref = f"{col}{row}"
            total += get_cell(data, sheet, cell_ref)
        
        return total
    except:
        return 0

def count_if_range(data, sheet, start_cell, end_cell, criteria):
    """Count items in range that match criteria."""
    try:
        # Extract row numbers
        start_row = int(re.search(r'(\\d+)', start_cell).group(1))
        end_row = int(re.search(r'(\\d+)', end_cell).group(1))
        
        # Extract column
        col = re.search(r'([A-Z]+)', start_cell).group(1)
        
        count = 0
        for row in range(start_row, end_row + 1):
            cell_ref = f"{col}{row}"
            if get_cell(data, sheet, cell_ref) == criteria:
                count += 1
        
        return count
    except:
        return 0

def find_text(search_text, search_in):
    """Find text within another text (like Excel FIND)."""
    try:
        if str(search_text) in str(search_in):
            return True
        return False
    except:
        return False

def safe_execute(func, default_value):
    """Execute function safely, return default on error."""
    try:
        return func()
    except:
        return default_value

def rows_count(start_cell, end_cell):
    """Count rows in range."""
    try:
        start_row = int(re.search(r'(\\d+)', start_cell).group(1))
        end_row = int(re.search(r'(\\d+)', end_cell).group(1))
        return end_row - start_row + 1
    except:
        return 1

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

def main():
    """Main conversion function."""
    start_time = time.time()
    with open('extracted_data.json', 'r') as f:
        extracted_data = json.load(f)
    
    converter = ExcelToPythonConverter()
    
    # Pre-calculate shared values
    shared_data = {}
    
    converted_formulas = []
    
    # Use tqdm to create a progress bar
    formulas = []
    for filename, file_data in extracted_data.items():
        formulas.extend(file_data.get("formulas", []))
    
    with tqdm(total=len(formulas), desc="Converting Formulas") as pbar:
        for formula_data in formulas:
            try:
                converted = converter.analyze_formula(
                    formula_data["formula"], 
                    formula_data["cell"], 
                    formula_data["sheet"],
                    shared_data  # Pass shared_data
                )
                converted_formulas.append(converted)
                logging.info(f"✓ Converted {formula_data['cell']}: {formula_data['formula']}")
            except Exception as e:
                logging.error(f"✗ Error converting {formula_data['cell']}: {e}")
            pbar.update(1)  # Update progress bar after each formula
    
    print(f"\nSuccessfully converted {len(converted_formulas)} formulas")
    
    # Generate Python file with all rule functions
    if converted_formulas:
        python_code = generate_python_rules_file(converter, converted_formulas, shared_data)
        
        # Write to file
        output_file = "converted_rules_fixed.py"
        with open(output_file, 'w') as f:
            f.write(python_code)
        print(f"✓ Generated Python rules file: {output_file}")
        
        # Also generate a summary JSON file
        summary = []
        for conv in converted_formulas:
            summary.append({
                "cell": conv.cell_reference,
                "sheet": conv.sheet,
                "rule_type": conv.rule_type,
                "description": conv.description,
                "original_formula": conv.original_formula,
                "python_expression": conv.python_expression,
                "dependencies": conv.dependencies
            })
        
        with open("conversion_summary_fixed.json", 'w') as f:
            json.dump(summary, f, indent=2)
        print(f"✓ Generated conversion summary: conversion_summary_fixed.json")
    
    end_time = time.time()
    print(f"Total execution time: {end_time - start_time:.2f} seconds")

if __name__ == "__main__":
    main()