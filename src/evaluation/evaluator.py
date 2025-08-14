import logging
from src.conversion.excel_functions import * 

def evaluate_rules(rules, data):
    """
    Evaluates a list of rules against a dataset.

    Args:
        rules (list): A list of dictionaries, where each dictionary represents a rule.
        data (dict): A dictionary containing the data to evaluate the rules against.

    Returns:
        dict: A dictionary containing the evaluation results.
    """
    results = {}
    eval_globals = {
        'data': data,
        'get_cell': get_cell,
        'get_value': get_value,
        'sum_range': sum_range,
        'count_range': count_range,
        'average_range': average_range,
        'count_if_range': count_if_range,
        'count_if': count_if,
        'sum_if': sum_if,
        'sum_keys': sum_keys,
        'average_keys': average_keys,
        'count_if_keys': count_if_keys,
        'sum_if_keys': sum_if_keys,
        'countifs_keys': countifs_keys,
        'safe_execute': safe_execute,
        'is_error': is_error,
        'concat': concat,
        'round_down': round_down,
        'rows_count': rows_count,
        'rows_count_keys': rows_count_keys,
        'find_text': find_text,
        'vlookup': vlookup,
        'index': index,
        'indirect': indirect,
        'countifs': countifs,
        'eomonth': eomonth,
        'yearfrac': yearfrac,
    }

    for rule in rules:
        try:
            result = eval(rule['python_expression'], eval_globals)
            results[rule['cell']] = result
        except Exception as e:
            logging.error(f"Error evaluating rule for cell {rule['cell']}: {e}")
            results[rule['cell']] = None
    return results
