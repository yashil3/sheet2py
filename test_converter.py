import pytest
from formula_converter_improved import ExcelToPythonConverter
from excel_functions import *  # Import helper functions for testing

# Mock data for testing
mock_data = {
    "Formulas": {"A1": 15, "B1": 5, "A2": 20, "A3": 30, "A4": 5, "A5": 12},
    "Source": {"C1": 1, "C2": 2}
}

# Initialize the converter once for all tests
converter = ExcelToPythonConverter(mock_data)
shared_data = {}


CONVERSION_TEST_CASES = [
    ("=A1", "get_cell(self.data, 'Formulas', 'A1')"),
    ("=SUM(B2:B9)", "sum_range(self.data, 'Formulas', 'B2', 'B9')"),
    ("=SUM(Source!C1:C10)", "sum_range(self.data, 'Source', 'C1', 'C10')"),
    ('=IF(A1>10,"High","Low")', '("High" if (get_cell(self.data, \'Formulas\', \'A1\') > 10) else "Low")'),
    ('=AND(A1>0,B1<10)', '((get_cell(self.data, \'Formulas\', \'A1\') > 0) and (get_cell(self.data, \'Formulas\', \'B1\') < 10))'),
    ('=IFERROR(A1/B1,0)', "safe_execute(lambda: ((get_cell(self.data, 'Formulas', 'A1') / get_cell(self.data, 'Formulas', 'B1'))), 0)"),
    ('=IF(OR(A1="X",A1="Y"),1,0)', '((1 if ((get_cell(self.data, \'Formulas\', \'A1\') == "X") or (get_cell(self.data, \'Formulas\', \'A1\') == "Y")) else 0))'),
    
    
    ("=NOT(A1=0)", "not (get_cell(self.data, 'Formulas', 'A1') == 0)"),
    ("=OR(A1>5,A1<2,B1=3)", "((get_cell(self.data, 'Formulas', 'A1') > 5) or (get_cell(self.data, 'Formulas', 'A1') < 2) or (get_cell(self.data, 'Formulas', 'B1') == 3))"),
    ("=AVERAGE(A1:A3)", "average_range(self.data, 'Formulas', 'A1', 'A3')"),
    ("=SUMIF(A1:A5,\">10\")", "sum_if(self.data, 'Formulas', 'A1', 'A5', \">10\")"),
    ("=COUNTIF(A1:A5,\">5\")", "count_if(self.data, 'Formulas', 'A1', 'A5', \">5\")"),
    ("=CONCAT(\"Hello\",\" \",\"World\")", "concat(\"Hello\", \" \", \"World\")"),
    ("=LEN(\"Test\")", "len(\"Test\")"),
    ("=ROUND(A1,2)", "round(get_cell(self.data, 'Formulas', 'A1'), 2)"),
    ("=IF(ISERROR(A1/B1),\"Error\",A1/B1)",
     "('Error' if is_error(lambda: (get_cell(self.data, 'Formulas', 'A1') / get_cell(self.data, 'Formulas', 'B1'))) else (get_cell(self.data, 'Formulas', 'A1') / get_cell(self.data, 'Formulas', 'B1')))"),
    ("=INDEX(A1:C3,2,3)", "index(self.data, 'Formulas', 'A1:C3', 2, 3)"),
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
