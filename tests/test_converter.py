import pytest
import networkx as nx
from src.conversion.converter import ExcelToPythonConverter, ConvertedFormula, build_dependency_graph, topological_sort
from src.conversion.excel_functions import *  # Import helper functions for testing

# Mock data for testing
mock_data = {
    "Formulas": {"A1": 15, "B1": 5, "A2": 20, "A3": 30, "A4": 5, "A5": 12, "B2": "=A1*2", "C1": "=B2+A2"},
    "Source": {"C1": 1, "C2": 2}
}

# Initialize the converter once for all tests
converter = ExcelToPythonConverter(mock_data)
shared_data = {}

CONVERSION_TEST_CASES = [
    ("=A1", "get_cell(self.data, 'TestSheet', 'A1')"),
    ("=SUM(B2:B9)", "sum_range(self.data, 'TestSheet', 'B2', 'B9')"),
    ("=SUM(Source!C1:C10)", "sum_range(self.data, 'Source', 'C1', 'C10')"),
    ('=IF(A1>10,"High","Low")', '("High" if (get_cell(self.data, \'TestSheet\', \'A1\') > 10) else "Low")'),
    ('=AND(A1>0,B1<10)', '((get_cell(self.data, \'TestSheet\', \'A1\') > 0) and (get_cell(self.data, \'TestSheet\', \'B1\') < 10))'),
    ('=IFERROR(A1/B1,0)', "safe_execute(lambda: ((get_cell(self.data, 'TestSheet', 'A1') / get_cell(self.data, 'TestSheet', 'B1'))), 0)"),
    ('=IF(OR(A1="X",A1="Y"),1,0)', '((1 if ((get_cell(self.data, \'TestSheet\', \'A1\') == "X") or (get_cell(self.data, \'TestSheet\', \'A1\') == "Y")) else 0))'),
    ("=NOT(A1=0)", "not (get_cell(self.data, 'TestSheet', 'A1') == 0)"),
    ("=OR(A1>5,A1<2,B1=3)", "((get_cell(self.data, 'TestSheet', 'A1') > 5) or (get_cell(self.data, 'TestSheet', 'A1') < 2) or (get_cell(self.data, 'TestSheet', 'B1') == 3))"),
    ("=AVERAGE(A1:A3)", "average_range(self.data, 'TestSheet', 'A1', 'A3')"),
    ('=SUMIF(A1:A5,">10")', 'sum_if(self.data, \'TestSheet\', \'A1\', \'A5\', ">10")'),
    ('=COUNTIF(A1:A5,">5")', 'count_if(self.data, \'TestSheet\', \'A1\', \'A5\', ">5")'),
    ('=CONCAT("Hello"," ","World")', 'concat("Hello", " ", "World")'),
    ('=LEN("Test")', 'len("Test")'),
    ("=ROUND(A1,2)", "round(get_cell(self.data, 'TestSheet', 'A1'), 2)"),
    ('=IF(ISERROR(A1/B1),"Error",A1/B1)',
     '("Error" if is_error(lambda: (get_cell(self.data, \'TestSheet\', \'A1\') / get_cell(self.data, \'TestSheet\', \'B1\'))) else (get_cell(self.data, \'TestSheet\', \'A1\') / get_cell(self.data, \'TestSheet\', \'B1\')))'),
    ("=INDEX(A1:C3,2,3)", "index(self.data, 'TestSheet', 'A1:C3', 2, 3)"),
]

@pytest.mark.parametrize("formula, expected", CONVERSION_TEST_CASES)
def test_formula_conversion(formula, expected):
    """Tests that individual formulas are converted correctly."""
    result = converter.analyze_formula(formula, "A1", "TestSheet", shared_data)
    cleaned_result = " ".join(result.python_expression.split())
    cleaned_expected = " ".join(expected.split())
    assert cleaned_result == cleaned_expected

def test_invalid_formula():
    """Tests that the parser handles invalid formulas gracefully."""
    with pytest.raises(Exception):
        converter.analyze_formula("=SUM(A1:)", "A1", "TestSheet", shared_data)

def test_dependency_ordering():
    """Tests that formulas with dependencies are processed in the correct order."""
    formulas_data = [
        {"formula": "=C1+A2", "cell": "D1", "sheet": "Sheet1"},
        {"formula": "=B1+A1", "cell": "C1", "sheet": "Sheet1"},
        {"formula": "=A1", "cell": "B1", "sheet": "Sheet1"},
    ]

    converted_formulas = [converter.analyze_formula(f["formula"], f["cell"], f["sheet"], shared_data) for f in formulas_data]
    
    dependency_graph = build_dependency_graph(converted_formulas)
    sorted_cells = topological_sort(dependency_graph)

    # Assert the relative ordering of dependent cells
    assert sorted_cells.index("Sheet1!B1") < sorted_cells.index("Sheet1!C1")
    assert sorted_cells.index("Sheet1!C1") < sorted_cells.index("Sheet1!D1")
    assert sorted_cells.index("Sheet1!A1") < sorted_cells.index("Sheet1!B1")
    assert sorted_cells.index("Sheet1!A2") < sorted_cells.index("Sheet1!D1")

def test_circular_dependency():
    """Tests that a circular dependency raises a ValueError."""
    formulas_data = [
        {"formula": "=B1", "cell": "A1", "sheet": "Sheet1"},
        {"formula": "=A1", "cell": "B1", "sheet": "Sheet1"},
    ]
    converted_formulas = [converter.analyze_formula(f["formula"], f["cell"], f["sheet"], shared_data) for f in formulas_data]
    dependency_graph = build_dependency_graph(converted_formulas)
    with pytest.raises(ValueError, match="Circular dependency detected!"):
        topological_sort(dependency_graph)

def test_self_dependency():
    """Tests that a formula referencing itself is detected."""
    formulas_data = [
        {"formula": "=A1", "cell": "A1", "sheet": "Sheet1"},
    ]
    converted_formulas = [converter.analyze_formula(f["formula"], f["cell"], f["sheet"], shared_data) for f in formulas_data]
    dependency_graph = build_dependency_graph(converted_formulas)
    with pytest.raises(ValueError, match="Circular dependency detected!"):
        topological_sort(dependency_graph)


# =RIGHT(@CELL("filename",A1),LEN(@CELL("filename",A1))-FIND("]",@CELL("filename",A1)))


# Converting Formulas:   0%|          | 0/5 [00:00<?, ?it/s]2025-07-31 14:28:38,579 - v Converted B1: =FIND("]", A1)
# line 1:0 token recognition error at: '''
# line 1:1 token recognition error at: '['
# line 1:3 token recognition error at: ']'
# 2025-07-31 14:28:38,580 - X Error converting B2: Invalid formula: ='[1]WS White Space'!A1
# line 1:0 token recognition error at: '$'
# line 1:2 token recognition error at: '$'
# 2025-07-31 14:28:38,580 - X Error converting B3: Invalid formula: =$B$1
# 2025-07-31 14:28:38,580 - X Error converting B4: Invalid formula: =YEARFRAC(E4, D4)
# line 1:9 token recognition error at: '&'
# 2025-07-31 14:28:38,581 - X Error converting B5: Invalid formula: ="HELLO " & D5
#  
# Converting Formulas: 100%|##########| 5/5 [00:00<00:00, 2001.67it/s]
#  
# Successfully converted 1 formulas
# v Formulas topologically sorted.
# Warning: Could not find formula for cell UnparsedFormulas!A1
# v Generated Python rules file: data/output\converted_rules.py
# v Generated conversion summary: data/output\conversion_summary.json
# Total execution time: 0.04 seconds