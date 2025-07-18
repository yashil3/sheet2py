import re
import json
from typing import Dict, Any, List, Optional
from dataclasses import dataclass
import time
import logging
from tqdm import tqdm  # Import tqdm for progress bar
import antlr4
from ANTLR.ExcelFormulaLexer import ExcelFormulaLexer
from ANTLR.ExcelFormulaParser import ExcelFormulaParser
from ANTLR.FormulaConverterVisitor import FormulaConverterVisitor
from rules_generator import generate_python_rules_file

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

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
    
    def __init__(self, data: Dict[str, Any]):
        self.data = data
        self.converted_cells = set()
        
    # All conversion is now handled in analyze_formula via the ANTLR grammar+visitor.
    def convert_expression(
        self,
        expr: str,
        cell: str,
        sheet: str,
        shared_data: Dict[str, Any]
    ) -> str:
        """
        Convert a complete Excel expression to Python by parsing with ANTLR
        and visiting the parse tree.
        """
        return self.analyze_formula(expr, cell, sheet, shared_data).python_expression

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
        """Analyze and convert a single formula using ANTLR."""
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

        # strip leading '=' so grammar can parse it
        formula_body = formula[1:] if formula.startswith('=') else formula
        input_stream = antlr4.InputStream(formula_body)

        lexer = ExcelFormulaLexer(input_stream)
        stream = antlr4.CommonTokenStream(lexer)
        parser = ExcelFormulaParser(stream)
        
        # remove default listeners and bail on first syntax error
        parser.removeErrorListeners()
        from antlr4.error.Errors import ParseCancellationException
        from antlr4.error.ErrorStrategy import BailErrorStrategy
        parser._errHandler = BailErrorStrategy()
        try:
            tree = parser.formula()  # start rule
        except ParseCancellationException as e:
            raise Exception(f"Invalid formula: {formula}") from e

        visitor = FormulaConverterVisitor(self.data, shared_data)
        python_expression = visitor.visit(tree)  # now works because visitor is a real ParseTreeVisitor

        # Classify the formula based on its structure
        rule_type = self.classify_formula(formula)

        # Generate description
        description = self.generate_description(formula, rule_type)

        return ConvertedFormula(
            original_formula=original,
            python_expression=python_expression,
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

def main():
    """Main conversion function."""
    start_time = time.time()
    with open('extracted_data.json', 'r') as f:
        extracted_data = json.load(f)

    converter = ExcelToPythonConverter(extracted_data)

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