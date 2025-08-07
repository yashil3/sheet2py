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
    ("=A1", "get_cell(data, 'TestSheet', 'A1')"),
    ("=SUM(B2:B9)", "sum_range(data, 'TestSheet', 'B2', 'B9')"),
    ("=SUM(Source!C1:C10)", "sum_range(data, 'Source', 'C1', 'C10')"),
    ('=IF(A1>10,"High","Low")', '("High" if (get_cell(data, \'TestSheet\', \'A1\') > 10) else "Low")'),
    ('=AND(A1>0,B1<10)', '((get_cell(data, \'TestSheet\', \'A1\') > 0) and (get_cell(data, \'TestSheet\', \'B1\') < 10))'),
    ('=IFERROR(A1/B1,0)', "safe_execute(lambda: ((get_cell(data, 'TestSheet', 'A1') / get_cell(data, 'TestSheet', 'B1'))), 0)"),
    ('=IF(OR(A1="X",A1="Y"),1,0)', '((1 if ((get_cell(data, \'TestSheet\', \'A1\') == "X") or (get_cell(data, \'TestSheet\', \'A1\') == "Y")) else 0))'),
    ("=NOT(A1=0)", "not (get_cell(data, 'TestSheet', 'A1') == 0)"),
    ("=OR(A1>5,A1<2,B1=3)", "((get_cell(data, 'TestSheet', 'A1') > 5) or (get_cell(data, 'TestSheet', 'A1') < 2) or (get_cell(data, 'TestSheet', 'B1') == 3))"),
    ("=AVERAGE(A1:A3)", "average_range(data, 'TestSheet', 'A1', 'A3')"),
    ('=SUMIF(A1:A5,">10")', 'sum_if(data, \'TestSheet\', \'A1\', \'A5\', ">10")'),
    ('=COUNTIF(A1:A5,">5")', 'count_if(data, \'TestSheet\', \'A1\', \'A5\', ">5")'),
    ('=CONCAT("Hello"," ","World")', 'concat("Hello", " ", "World")'),
    ('=LEN("Test")', 'len("Test")'),
    ("=ROUND(A1,2)", "round(get_cell(data, 'TestSheet', 'A1'), 2)"),
    ('=IF(ISERROR(A1/B1),"Error",A1/B1)',
     '("Error" if is_error(lambda: (get_cell(data, \'TestSheet\', \'A1\') / get_cell(data, \'TestSheet\', \'B1\'))) else (get_cell(data, \'TestSheet\', \'A1\') / get_cell(data, \'TestSheet\', \'B1\')))'),
    ("=INDEX(A1:C3,2,3)", "index(data, 'TestSheet', 'A1:C3', 2, 3)"),
    # RIGHT function tests
    ('=RIGHT("Hello World",5)', 'right_text("Hello World", 5)'),
    ('=RIGHT(A1,3)', 'right_text(get_cell(data, \'TestSheet\', \'A1\'), 3)'),
    ('=RIGHT("Test",LEN("Test"))', 'right_text("Test", len("Test"))'),
    ('=RIGHT("Test",LEN("Test")-1)', 'right_text("Test", (len("Test") - 1))'),
    ('=RIGHT(A1,LEN(A1)-FIND("]",A1))', 'right_text(get_cell(data, \'TestSheet\', \'A1\'), (len(get_cell(data, \'TestSheet\', \'A1\')) - find_text("]", get_cell(data, \'TestSheet\', \'A1\'))))'),
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

def test_right_function_with_brackets():
    """Tests RIGHT function specifically with bracket-related scenarios."""
    test_data = {
        "TestSheet": {
            "A1": "[1]My Workbook.xlsx",
            "A2": "[2]Another File.xlsx", 
            "A3": "NoBrackets.xlsx",
            "A4": "[1]File with spaces.xlsx",
            "A5": "C:\\Users\\John\\Documents\\[1]My Workbook.xlsx"
        }
    }
 
    test_cases = [
        ("=RIGHT(A1,5)", "right_text(get_cell(data, 'TestSheet', 'A1'), 5)"),
        ("=RIGHT(A2,10)", "right_text(get_cell(data, 'TestSheet', 'A2'), 10)"),
        ("=RIGHT(A3,4)", "right_text(get_cell(data, 'TestSheet', 'A3'), 4)"),
        ("=RIGHT(A4,15)", "right_text(get_cell(data, 'TestSheet', 'A4'), 15)"),
        ("=RIGHT(A5,16)", "right_text(get_cell(data, 'TestSheet', 'A5'), 16)"),
       
        ("=RIGHT(A1,LEN(A1)-FIND(\"]\",A1))", 
         "right_text(get_cell(data, 'TestSheet', 'A1'), (len(get_cell(data, 'TestSheet', 'A1')) - find_text(\"]\", get_cell(data, 'TestSheet', 'A1'))))"),
        ("=RIGHT(A2,LEN(A2)-FIND(\"]\",A2))", 
         "right_text(get_cell(data, 'TestSheet', 'A2'), (len(get_cell(data, 'TestSheet', 'A2')) - find_text(\"]\", get_cell(data, 'TestSheet', 'A2'))))"),
    ]
    
    for formula, expected in test_cases:
        result = converter.analyze_formula(formula, "B1", "TestSheet", shared_data)
        cleaned_result = " ".join(result.python_expression.split())
        cleaned_expected = " ".join(expected.split())
        assert cleaned_result == cleaned_expected, f"Failed for {formula}"

def test_right_function_execution():
    """Tests that the RIGHT function actually executes correctly with bracket data."""
    from src.conversion.excel_functions import right_text, find_text
    
    
    test_strings = [
        ("[1]My Workbook.xlsx", "My Workbook.xlsx"),
        ("[2]Another File.xlsx", "Another File.xlsx"),
        ("NoBrackets.xlsx", "NoBrackets.xlsx"),
        ("[1]File with spaces.xlsx", "File with spaces.xlsx"),
        ("C:\\Users\\John\\Documents\\[1]My Workbook.xlsx", "My Workbook.xlsx")
    ]
    
    for input_str, expected in test_strings:
    
        bracket_pos = find_text("]", input_str)
        if bracket_pos > 0:
            chars_after = len(input_str) - bracket_pos
            result = right_text(input_str, chars_after)
        else:
            result = input_str  # No bracket found
            
        assert result == expected, f"Failed for '{input_str}': got '{result}', expected '{expected}'"