# Auto-generated Python rule functions
# Generated from Excel formulas

# Import Excel-like helper functions
from src.conversion.excel_functions import *

# Shared data (pre-calculated values)
shared_data = {}

# Generated rule functions
def rule_formulas_b10(data, shared_data):
    """
    General calculation or transformation
    Original Excel formula: =SUM(B2:B9)
    Cell: Formulas!B10
    Rule type: general_calculation
    Dependencies: Formulas!B2, Formulas!B3, Formulas!B4, Formulas!B5, Formulas!B6, Formulas!B7, Formulas!B8, Formulas!B9
    """
    try:
        result = sum_range(data, 'Formulas', 'B2', 'B9')
        return result
    except Exception as e:
        print(f"Error in rule_formulas_b10: {e}")
        return None

def rule_formulas_c2(data, shared_data):
    """
    Calculate relative weight as percentage of total
    Original Excel formula: =$B2/SUM($B$2:$B$9)
    Cell: Formulas!C2
    Rule type: weight_calculation
    Dependencies: Formulas!B2, Formulas!B3, Formulas!B4, Formulas!B5, Formulas!B6, Formulas!B7, Formulas!B8, Formulas!B9, Formulas!$B2
    """
    try:
        result = (get_cell(data, 'Formulas', '$B2') / sum_range(data, 'Formulas', '$B$2', '$B$9'))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_c2: {e}")
        return None

def rule_formulas_e2(data, shared_data):
    """
    Check if condition is met and return 1 or 0
    Original Excel formula: =IFERROR(IF(FIND(Source!G2,Source!B14)>0,1, 0),0)
    Cell: Formulas!E2
    Rule type: condition_check
    Dependencies: Source!G2, Source!B14
    """
    try:
        result = safe_execute(lambda: ((1 if (find_text(get_cell(data, 'Source', 'G2'), get_cell(data, 'Source', 'B14')) > 0) else 0)), 0)
        return result
    except Exception as e:
        print(f"Error in rule_formulas_e2: {e}")
        return None

def rule_formulas_c3(data, shared_data):
    """
    Calculate relative weight as percentage of total
    Original Excel formula: =$B3/SUM($B$2:$B$9)
    Cell: Formulas!C3
    Rule type: weight_calculation
    Dependencies: Formulas!B2, Formulas!B3, Formulas!B4, Formulas!B5, Formulas!B6, Formulas!B7, Formulas!B8, Formulas!B9, Formulas!$B3
    """
    try:
        result = (get_cell(data, 'Formulas', '$B3') / sum_range(data, 'Formulas', '$B$2', '$B$9'))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_c3: {e}")
        return None

def rule_formulas_e3(data, shared_data):
    """
    General calculation or transformation
    Original Excel formula: =IF(OR(Source!G3=Source!B20,Source!H3=Source!B20,Source!I3=Source!B20),1,0)
    Cell: Formulas!E3
    Rule type: general_calculation
    Dependencies: Source!I3, Source!H3, Source!B20, Source!G3
    """
    try:
        result = ((1 if ((get_cell(data, 'Source', 'G3') == get_cell(data, 'Source', 'B20')) or (get_cell(data, 'Source', 'H3') == get_cell(data, 'Source', 'B20')) or (get_cell(data, 'Source', 'I3') == get_cell(data, 'Source', 'B20'))) else 0))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_e3: {e}")
        return None

def rule_formulas_c4(data, shared_data):
    """
    Calculate relative weight as percentage of total
    Original Excel formula: =$B4/SUM($B$2:$B$9)
    Cell: Formulas!C4
    Rule type: weight_calculation
    Dependencies: Formulas!B2, Formulas!B3, Formulas!B4, Formulas!B5, Formulas!B6, Formulas!B7, Formulas!B8, Formulas!B9, Formulas!$B4
    """
    try:
        result = (get_cell(data, 'Formulas', '$B4') / sum_range(data, 'Formulas', '$B$2', '$B$9'))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_c4: {e}")
        return None

def rule_formulas_e4(data, shared_data):
    """
    Check if condition is met and return 1 or 0
    Original Excel formula: =IF(Source!B5<=60000,1,0)
    Cell: Formulas!E4
    Rule type: condition_check
    Dependencies: Source!B5
    """
    try:
        result = (1 if (get_cell(data, 'Source', 'B5') <= 60000) else 0)
        return result
    except Exception as e:
        print(f"Error in rule_formulas_e4: {e}")
        return None

def rule_formulas_c5(data, shared_data):
    """
    Calculate relative weight as percentage of total
    Original Excel formula: =$B5/SUM($B$2:$B$9)
    Cell: Formulas!C5
    Rule type: weight_calculation
    Dependencies: Formulas!B2, Formulas!B3, Formulas!B4, Formulas!B5, Formulas!B6, Formulas!B7, Formulas!B8, Formulas!B9, Formulas!$B5
    """
    try:
        result = (get_cell(data, 'Formulas', '$B5') / sum_range(data, 'Formulas', '$B$2', '$B$9'))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_c5: {e}")
        return None

def rule_formulas_e5(data, shared_data):
    """
    Check if condition is met and return 1 or 0
    Original Excel formula: =IF(Source!B6>=2015, 1,0)
    Cell: Formulas!E5
    Rule type: condition_check
    Dependencies: Source!B6
    """
    try:
        result = (1 if (get_cell(data, 'Source', 'B6') >= 2015) else 0)
        return result
    except Exception as e:
        print(f"Error in rule_formulas_e5: {e}")
        return None

def rule_formulas_c6(data, shared_data):
    """
    Calculate relative weight as percentage of total
    Original Excel formula: =$B6/SUM($B$2:$B$9)
    Cell: Formulas!C6
    Rule type: weight_calculation
    Dependencies: Formulas!B2, Formulas!B3, Formulas!B4, Formulas!B5, Formulas!B6, Formulas!B7, Formulas!B8, Formulas!B9, Formulas!$B6
    """
    try:
        result = (get_cell(data, 'Formulas', '$B6') / sum_range(data, 'Formulas', '$B$2', '$B$9'))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_c6: {e}")
        return None

def rule_formulas_e6(data, shared_data):
    """
    Count how many items match the criteria
    Original Excel formula: =(COUNTIF(Source!B8:B10, Source!G6)+COUNTIF(Source!B8:B10,Source!H6)+COUNTIF(Source!B8:B10,Source!I6)+COUNTIF(Source!B8:B10,Source!J6))/4
    Cell: Formulas!E6
    Rule type: count_matching
    Dependencies: Source!B8, Source!B9, Source!B10, Source!J6, Source!H6, Source!G6, Source!I6
    """
    try:
        result = (((((count_if(data, 'Source', 'B8', 'B10', get_cell(data, 'Source', 'G6')) + count_if(data, 'Source', 'B8', 'B10', get_cell(data, 'Source', 'H6'))) + count_if(data, 'Source', 'B8', 'B10', get_cell(data, 'Source', 'I6'))) + count_if(data, 'Source', 'B8', 'B10', get_cell(data, 'Source', 'J6')))) / 4)
        return result
    except Exception as e:
        print(f"Error in rule_formulas_e6: {e}")
        return None

def rule_formulas_c7(data, shared_data):
    """
    Calculate relative weight as percentage of total
    Original Excel formula: =$B7/SUM($B$2:$B$9)
    Cell: Formulas!C7
    Rule type: weight_calculation
    Dependencies: Formulas!B2, Formulas!B3, Formulas!B4, Formulas!B5, Formulas!B6, Formulas!B7, Formulas!B8, Formulas!B9, Formulas!$B7
    """
    try:
        result = (get_cell(data, 'Formulas', '$B7') / sum_range(data, 'Formulas', '$B$2', '$B$9'))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_c7: {e}")
        return None

def rule_formulas_e7(data, shared_data):
    """
    Assign score based on value ranges
    Original Excel formula: =IF(OR(Source!B12="6-cylinder", Source!B12="8-cylinder"),1,0)
    Cell: Formulas!E7
    Rule type: scoring_rule
    Dependencies: Source!B12
    """
    try:
        result = ((1 if ((get_cell(data, 'Source', 'B12') == "6-cylinder") or (get_cell(data, 'Source', 'B12') == "8-cylinder")) else 0))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_e7: {e}")
        return None

def rule_formulas_c8(data, shared_data):
    """
    Calculate relative weight as percentage of total
    Original Excel formula: =$B8/SUM($B$2:$B$9)
    Cell: Formulas!C8
    Rule type: weight_calculation
    Dependencies: Formulas!B2, Formulas!B3, Formulas!B4, Formulas!B5, Formulas!B6, Formulas!B7, Formulas!B8, Formulas!B9, Formulas!$B8
    """
    try:
        result = (get_cell(data, 'Formulas', '$B8') / sum_range(data, 'Formulas', '$B$2', '$B$9'))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_c8: {e}")
        return None

def rule_formulas_e8(data, shared_data):
    """
    General calculation or transformation
    Original Excel formula: =IF(Source!G8=Source!B13,1, 0)
    Cell: Formulas!E8
    Rule type: general_calculation
    Dependencies: Source!G8, Source!B13
    """
    try:
        result = (1 if (get_cell(data, 'Source', 'G8') == get_cell(data, 'Source', 'B13')) else 0)
        return result
    except Exception as e:
        print(f"Error in rule_formulas_e8: {e}")
        return None

def rule_formulas_c9(data, shared_data):
    """
    Calculate relative weight as percentage of total
    Original Excel formula: =$B9/SUM($B$2:$B$9)
    Cell: Formulas!C9
    Rule type: weight_calculation
    Dependencies: Formulas!B2, Formulas!B3, Formulas!B4, Formulas!B5, Formulas!B6, Formulas!B7, Formulas!B8, Formulas!B9, Formulas!$B9
    """
    try:
        result = (get_cell(data, 'Formulas', '$B9') / sum_range(data, 'Formulas', '$B$2', '$B$9'))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_c9: {e}")
        return None

def rule_formulas_e9(data, shared_data):
    """
    Assign score based on value ranges
    Original Excel formula: =(IF(AND(Source!G9="Auto-dimming Mirrors",Source!B16="YES"),1,0) + IF(AND(Source!G10="Heated seats",Source!B15="YES"),1,0) + IF(AND(Source!G11="Leather Seats",Source!B19="YES"),1,0) + IF(AND(Source!G12="Auto Wipers",Source!B17="YES"),1,0))/ROWS(Source!G9:G12)
    Cell: Formulas!E9
    Rule type: scoring_rule
    Dependencies: Source!G9, Source!G10, Source!G12, Source!B19, Source!B16, Source!G11, Source!B15, Source!B17
    """
    try:
        result = ((((((1 if ((get_cell(data, 'Source', 'G9') == "Auto-dimming Mirrors") and (get_cell(data, 'Source', 'B16') == "YES")) else 0) + (1 if ((get_cell(data, 'Source', 'G10') == "Heated seats") and (get_cell(data, 'Source', 'B15') == "YES")) else 0)) + (1 if ((get_cell(data, 'Source', 'G11') == "Leather Seats") and (get_cell(data, 'Source', 'B19') == "YES")) else 0)) + (1 if ((get_cell(data, 'Source', 'G12') == "Auto Wipers") and (get_cell(data, 'Source', 'B17') == "YES")) else 0))) / rows_count('G9', 'G12'))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_e9: {e}")
        return None

def rule_unparsedformulas_b1(data, shared_data):
    """
    Search for text within another text field
    Original Excel formula: =FIND("]", A1)
    Cell: UnparsedFormulas!B1
    Rule type: text_search
    Dependencies: UnparsedFormulas!A1
    """
    try:
        result = find_text("]", get_cell(data, 'UnparsedFormulas', 'A1'))
        return result
    except Exception as e:
        print(f"Error in rule_unparsedformulas_b1: {e}")
        return None

def rule_unparsedformulas_b2(data, shared_data):
    """
    General calculation or transformation
    Original Excel formula: ='[1]WS White Space'!A1
    Cell: UnparsedFormulas!B2
    Rule type: general_calculation
    Dependencies: [1]WS White Space!A1
    """
    try:
        result = get_cell(data, '[1]WS White Space', 'A1')
        return result
    except Exception as e:
        print(f"Error in rule_unparsedformulas_b2: {e}")
        return None

def rule_unparsedformulas_b3(data, shared_data):
    """
    General calculation or transformation
    Original Excel formula: =$B$1
    Cell: UnparsedFormulas!B3
    Rule type: general_calculation
    Dependencies: UnparsedFormulas!$B$1
    """
    try:
        result = get_cell(data, 'UnparsedFormulas', '$B$1')
        return result
    except Exception as e:
        print(f"Error in rule_unparsedformulas_b3: {e}")
        return None

def rule_unparsedformulas_b4(data, shared_data):
    """
    General calculation or transformation
    Original Excel formula: =YEARFRAC(E4, D4)
    Cell: UnparsedFormulas!B4
    Rule type: general_calculation
    Dependencies: UnparsedFormulas!D4, UnparsedFormulas!E4
    """
    try:
        result = yearfrac(get_cell(data, 'UnparsedFormulas', 'E4'), get_cell(data, 'UnparsedFormulas', 'D4'))
        return result
    except Exception as e:
        print(f"Error in rule_unparsedformulas_b4: {e}")
        return None

def rule_unparsedformulas_b5(data, shared_data):
    """
    General calculation or transformation
    Original Excel formula: ="HELLO " & D5
    Cell: UnparsedFormulas!B5
    Rule type: general_calculation
    Dependencies: UnparsedFormulas!D5
    """
    try:
        result = str("HELLO ") + str(get_cell(data, 'UnparsedFormulas', 'D5'))
        return result
    except Exception as e:
        print(f"Error in rule_unparsedformulas_b5: {e}")
        return None

def rule_formulas_c10(data, shared_data):
    """
    General calculation or transformation
    Original Excel formula: =SUM(C2:C9)
    Cell: Formulas!C10
    Rule type: general_calculation
    Dependencies: Formulas!C2, Formulas!C3, Formulas!C4, Formulas!C5, Formulas!C6, Formulas!C7, Formulas!C8, Formulas!C9
    """
    try:
        result = sum_range(data, 'Formulas', 'C2', 'C9')
        return result
    except Exception as e:
        print(f"Error in rule_formulas_c10: {e}")
        return None

def rule_formulas_e10(data, shared_data):
    """
    General calculation or transformation
    Original Excel formula: =E2*C2+E3*C3+E4*C4+E5*C5+E6*C6+E7*C7+E8*C8+E9*C9
    Cell: Formulas!E10
    Rule type: general_calculation
    Dependencies: Formulas!E5, Formulas!E2, Formulas!E8, Formulas!E4, Formulas!C6, Formulas!C2, Formulas!C8, Formulas!C9, Formulas!E6, Formulas!C3, Formulas!C7, Formulas!E9, Formulas!E7, Formulas!C5, Formulas!C4, Formulas!E3
    """
    try:
        result = (((((((((((((((get_cell(data, 'Formulas', 'E2') * get_cell(data, 'Formulas', 'C2')) + get_cell(data, 'Formulas', 'E3')) * get_cell(data, 'Formulas', 'C3')) + get_cell(data, 'Formulas', 'E4')) * get_cell(data, 'Formulas', 'C4')) + get_cell(data, 'Formulas', 'E5')) * get_cell(data, 'Formulas', 'C5')) + get_cell(data, 'Formulas', 'E6')) * get_cell(data, 'Formulas', 'C6')) + get_cell(data, 'Formulas', 'E7')) * get_cell(data, 'Formulas', 'C7')) + get_cell(data, 'Formulas', 'E8')) * get_cell(data, 'Formulas', 'C8')) + get_cell(data, 'Formulas', 'E9')) * get_cell(data, 'Formulas', 'C9'))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_e10: {e}")
        return None

def rule_formulas_a14(data, shared_data):
    """
    Check if condition is met and return 1 or 0
    Original Excel formula: ="This vehicle is a " & IF(E10<0.2, "Terrible match", IF(E10<0.4, "Bad match", IF(E10<0.6, "OK Match", IF(E10<0.8, "Good match", "Great match"))))
    Cell: Formulas!A14
    Rule type: condition_check
    Dependencies: Formulas!E10
    """
    try:
        result = str("This vehicle is a ") + str(("Terrible match" if (get_cell(data, 'Formulas', 'E10') < 0.2) else ("Bad match" if (get_cell(data, 'Formulas', 'E10') < 0.4) else ("OK Match" if (get_cell(data, 'Formulas', 'E10') < 0.6) else ("Good match" if (get_cell(data, 'Formulas', 'E10') < 0.8) else "Great match")))))
        return result
    except Exception as e:
        print(f"Error in rule_formulas_a14: {e}")
        return None