import pytest
from src.conversion.converter import ExcelToPythonConverter
from src.evaluation.evaluator import evaluate_rules

# Shared mapping: cell -> semantic key
cell_to_key_map = {
    'Source': {
        'B5': 'Mileage',
        'B6': 'Year',
        'B12': 'Engine',
    }
}

# Data with both cell values and by_key mapping
mock_data = {
    'Source': {
        'B5': 57351,
        'B6': 2016,
        'B12': '6-cylinder',
        'by_key': {
            'Mileage': 57351,
            'Year': 2016,
            'Engine': '6-cylinder',
        }
    },
    'Formulas': {}
}


def _normalize(expr: str) -> str:
    return " ".join(expr.split())


def test_cell_reference_to_get_value():
    converter = ExcelToPythonConverter(mock_data)
    shared_data = {'cell_to_key_map': cell_to_key_map}
    result = converter.analyze_formula("=Source!B12", "A1", "Formulas", shared_data)
    assert _normalize(result.python_expression) == _normalize("get_value(data, 'Source', 'Engine')")
    assert 'Source:Engine' in result.input_keys


def test_sum_range_to_sum_keys():
    converter = ExcelToPythonConverter(mock_data)
    shared_data = {'cell_to_key_map': cell_to_key_map}
    result = converter.analyze_formula("=SUM(Source!B5:B6)", "A2", "Formulas", shared_data)
    assert _normalize(result.python_expression) == _normalize("sum_keys(data, 'Source', ['Mileage', 'Year'])")
    # Both keys should be captured as inputs
    assert set(result.input_keys) == set(['Source:Mileage', 'Source:Year'])


def test_countif_to_count_if_keys():
    converter = ExcelToPythonConverter(mock_data)
    shared_data = {'cell_to_key_map': cell_to_key_map}
    result = converter.analyze_formula('=COUNTIF(Source!B6:B6, ">2015")', "A3", "Formulas", shared_data)
    assert _normalize(result.python_expression) == _normalize('count_if_keys(data, \'Source\', [\'Year\'], ">2015")')


def test_average_range_to_average_keys():
    converter = ExcelToPythonConverter(mock_data)
    shared_data = {'cell_to_key_map': cell_to_key_map}
    result = converter.analyze_formula("=AVERAGE(Source!B5:B6)", "A4", "Formulas", shared_data)
    assert _normalize(result.python_expression) == _normalize("average_keys(data, 'Source', ['Mileage', 'Year'])")


def test_mixed_range_falls_back_to_cell_range():
    # B5 maps to Mileage, but B7 has no key mapping -> should fall back to sum_range
    converter = ExcelToPythonConverter({
        'Source': {
            'B5': 10,
            'B7': 5,
            'by_key': {'Mileage': 10}
        }
    })
    shared_data = {'cell_to_key_map': {'Source': {'B5': 'Mileage'}}}
    result = converter.analyze_formula("=SUM(Source!B5:B7)", "A5", "Formulas", shared_data)
    assert _normalize(result.python_expression) == _normalize("sum_range(data, 'Source', 'B5', 'B7')")


def test_execution_with_evaluator_sum_keys_and_get_value():
    # Build two rules and run through evaluator
    converter = ExcelToPythonConverter(mock_data)
    shared_data = {'cell_to_key_map': cell_to_key_map}
    r1 = converter.analyze_formula("=SUM(Source!B5:B6)", "C1", "Formulas", shared_data)
    r2 = converter.analyze_formula("=Source!B12", "C2", "Formulas", shared_data)

    rules = [
        {
            'cell': r1.cell_reference,
            'python_expression': r1.python_expression
        },
        {
            'cell': r2.cell_reference,
            'python_expression': r2.python_expression
        },
    ]

    results = evaluate_rules(rules, mock_data)
    assert results['C1'] == (57351 + 2016)
    assert results['C2'] == '6-cylinder'
