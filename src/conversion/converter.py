import re
import json
from typing import Dict, Any, List, Optional
from dataclasses import dataclass
import time
import logging
from tqdm import tqdm
import antlr4
from src.antlr_files.ExcelFormulaLexer import ExcelFormulaLexer
from src.antlr_files.ExcelFormulaParser import ExcelFormulaParser
from src.antlr_files.FormulaConverterVisitor import FormulaConverterVisitor
from .rules_generator import generate_python_rules_file
import networkx as nx

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

@dataclass
class ConvertedFormula:
    """Represents a converted formula with metadata."""
    original_formula: str
    python_expression: str
    cell_reference: str
    sheet: str
    dependencies: nx.DiGraph
    description: str
    rule_type: str

class ExcelToPythonConverter:
    """Fixed converter for Excel formulas to Python expressions."""

    def __init__(self, data: Dict[str, Any]):
        self.data = data
        self.converted_cells = set()

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

        formula_body = formula[1:] if formula.startswith('=') else formula
        input_stream = antlr4.InputStream(formula_body)

        lexer = ExcelFormulaLexer(input_stream)
        stream = antlr4.CommonTokenStream(lexer)
        parser = ExcelFormulaParser(stream)

        parser.removeErrorListeners()
        from antlr4.error.Errors import ParseCancellationException
        from antlr4.error.ErrorStrategy import BailErrorStrategy
        parser._errHandler = BailErrorStrategy()
        try:
            tree = parser.formula()
        except ParseCancellationException as e:
            raise Exception(f"Invalid formula: {formula}") from e

        visitor = FormulaConverterVisitor(self.data, shared_data, sheet)
        python_expression = visitor.visit(tree)

        dependencies = visitor.dependencies
        dep_list = []
        for dep in dependencies:
            if ':' in dep:
                sheet_name, cells = dep.split('!')
                start, end = cells.split(':')
                # Expand range into individual cells
                col_start = ''.join(filter(str.isalpha, start))
                row_start = int(''.join(filter(str.isdigit, start)))
                col_end = ''.join(filter(str.isalpha, end))
                row_end = int(''.join(filter(str.isdigit, end)))

                for col in range(ord(col_start), ord(col_end) + 1):
                    for row in range(row_start, row_end + 1):
                        dep_list.append(f"{sheet_name}!{chr(col)}{row}")
            else:
                dep_list.append(dep)

        seen = set()
        unique_deps = [x for x in dep_list if not (x in seen or seen.add(x))]

        deps_graph = nx.DiGraph()
        node_id = f"{sheet}!{cell}"
        deps_graph.add_node(node_id)
        for dep in unique_deps:
            deps_graph.add_edge(dep, node_id)

        rule_type = self.classify_formula(formula)
        description = self.generate_description(formula, rule_type)

        return ConvertedFormula(
            original_formula=original,
            python_expression=python_expression,
            cell_reference=cell,
            sheet=sheet,
            dependencies=deps_graph,
            description=description,
            rule_type=rule_type
        )

    def generate_python_function(self, converted: ConvertedFormula) -> str:
        """Generate a complete Python function from converted formula."""
        func_name = f"rule_{converted.sheet.lower()}_{converted.cell_reference.lower()}"

        node_id = f"{converted.sheet}!{converted.cell_reference}"
        deps = list(converted.dependencies.predecessors(node_id))
        deps_str = ", ".join(deps)

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

def build_dependency_graph(converted_formulas: List[ConvertedFormula]) -> nx.DiGraph:
    """Builds a dependency graph from a list of converted formulas using networkx."""
    all_graphs = [f.dependencies for f in converted_formulas]
    return nx.compose_all(all_graphs)

def topological_sort(graph: nx.DiGraph) -> List[str]:
    """Topologically sorts a dependency graph using networkx."""
    try:
        return list(nx.topological_sort(graph))
    except nx.NetworkXUnfeasible:
        raise ValueError("Circular dependency detected!")


