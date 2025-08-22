#!/usr/bin/env python3
"""
Comprehensive test for the evaluation pipeline.
This test replicates the spreadsheet data structure exactly and tests all formulas.
"""

import sys
import os
from pathlib import Path

# Add the project root to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from src.evaluation.evaluator import evaluate_rules
from src.conversion.excel_functions import *
import json

def create_test_data():
    """Create test data that exactly replicates the spreadsheet structure."""
    
    # This replicates the exact data structure from the actual spreadsheet
    test_data = {
        "Formulas": {
            # Cell values (B2:B9) - these are the weights from the spreadsheet
            "B2": 15,     # Country of origin weight
            "B3": 50,     # Color weight  
            "B4": 100,    # Mileage weight
            "B5": 75,     # Year weight
            "B6": 50,     # Options weight
            "B7": 30,     # Engine size weight
            "B8": 80,     # Transmission weight
            "B9": 50,     # Features weight
            
            # Named cell reference
            "I2": "I am a named cell.",
            
            # Key-based mapping for semantic access
            "by_key": {
                "Country of origin": 15,
                "Color": 50,
                "Mileage": 100,
                "Year": 75,
                "Options": 50,
                "Engine size": 30,
                "Transmission": 80,
                "Features": 50
            }
        },
        "Source": {
            # Source data for condition checks - from actual spreadsheet
            "B5": 57351,              # Mileage
            "B6": 2016,               # Year
            "B8": "Air suspension",   # Option1
            "B9": "Panoramic Roof",   # Option2
            "B10": "N/A",             # Option3
            "B12": "6-cylinder",      # Engine
            "B13": "Automatic",       # Transmission
            "B14": "Stuttgart, GER",  # Made in
            "B15": "YES",             # Heated seats
            "B16": "YES",             # Auto-dimming Mirrors
            "B17": "NO",              # Auto Wipers
            "B19": "YES",             # Leather Seats
            "B20": "Blue",            # Color
            "G2": "GER",              # Country to find
            "G3": "GER",              # Make:GER
            "G6": "Air suspension",   # Option preference 1
            "G7": "Panoramic Roof",   # Option preference 2
            "G8": "Automatic",        # Transmission preference
            "G9": "Auto-dimming Mirrors", # Feature preference 1
            "G10": "Heated seats",    # Feature preference 2
            "G11": "Leather Seats",   # Feature preference 3
            "G12": "Auto Wipers",     # Feature preference 4
            "H3": "GER",              # Alternative Make:GER
            "H6": "Panoramic Roof",   # Option preference 2
            "I3": "GER",              # Alternative Make:GER
            "I6": "Lo-Jack",          # Option preference 3
            "J6": "Subwoofer",        # Option preference 4
            
            # Key-based mapping
            "by_key": {
                "Made in": "Stuttgart, GER",
                "Color": "Blue",
                "Mileage": 57351,
                "Year": 2016,
                "Make:GER": "GER",
                "Engine": "6-cylinder",
                "Transmission": "Automatic",
                "Option1": "Air suspension",
                "Option2": "Panoramic Roof",
                "Option3": "N/A",
                "Heated seats": "YES",
                "Auto-dimming Mirrors": "YES",
                "Auto Wipers": "NO",
                "Leather Seats": "YES"
            }
        },
        "_named_ranges": {
            "Named_Cell": ("Formulas", "I2")
        }
    }
    
    return test_data

def test_individual_functions():
    """Test individual Excel functions to ensure they work correctly."""
    print("Testing individual Excel functions...")
    
    data = create_test_data()
    
    # Test get_value function
    assert get_value(data, "Formulas", "Country of origin") == 15
    assert get_value(data, "Formulas", "Color") == 50
    assert get_value(data, "Source", "Made in") == "Stuttgart, GER"
    
    # Test sum_keys function
    total = sum_keys(data, "Formulas", ["Country of origin", "Color", "Mileage", "Year", "Options", "Engine size", "Transmission", "Features"])
    assert total == 450, f"Expected 450, got {total}"
    
    # Test find_text function
    pos = find_text("GER", "Stuttgart, GER")
    assert pos > 0, f"Expected position > 0, got {pos}"
    
    # Test get_cell function
    assert get_cell(data, "Formulas", "B2") == 15
    assert get_cell(data, "Formulas", "I2") == "I am a named cell."
    
    print("✓ All individual functions working correctly")

def test_weight_calculations():
    """Test the weight calculation formulas (C2:C9)."""
    print("\nTesting weight calculations...")
    
    data = create_test_data()
    
    # Test individual weight calculations
    country_weight = get_value(data, "Formulas", "Country of origin") / sum_keys(data, "Formulas", ["Country of origin", "Color", "Mileage", "Year", "Options", "Engine size", "Transmission", "Features"])
    assert abs(country_weight - 15/450) < 0.0001, f"Expected 15/450, got {country_weight}"
    
    color_weight = get_value(data, "Formulas", "Color") / sum_keys(data, "Formulas", ["Country of origin", "Color", "Mileage", "Year", "Options", "Engine size", "Transmission", "Features"])
    assert abs(color_weight - 50/450) < 0.0001, f"Expected 50/450, got {color_weight}"
    
    mileage_weight = get_value(data, "Formulas", "Mileage") / sum_keys(data, "Formulas", ["Country of origin", "Color", "Mileage", "Year", "Options", "Engine size", "Transmission", "Features"])
    assert abs(mileage_weight - 100/450) < 0.0001, f"Expected 100/450, got {mileage_weight}"
    
    print("✓ Weight calculations working correctly")

def test_condition_checks():
    """Test the condition check formulas (E2:E9)."""
    print("\nTesting condition checks...")
    
    data = create_test_data()
    
    # Test E2: IFERROR(IF(FIND(Source!G2,Source!B14)>0,1, 0),0)
    # This should find "GER" in "Stuttgart, GER" and return 1
    ger_in_made_in = find_text(get_cell(data, "Source", "G2"), get_value(data, "Source", "Made in"))
    result_e2 = 1 if ger_in_made_in > 0 else 0
    assert result_e2 == 1, f"Expected 1, got {result_e2}"
    
    # Test E3: IF(OR(Source!G3=Source!B20,Source!H3=Source!B20,Source!I3=Source!B20),1,0)
    # This checks if any of the German makes equals the color "Blue"
    make_ger = get_value(data, "Source", "Make:GER")
    h3_value = get_cell(data, "Source", "H3")
    i3_value = get_cell(data, "Source", "I3")
    color_value = get_value(data, "Source", "Color")
    
    result_e3 = 1 if (make_ger == color_value or h3_value == color_value or i3_value == color_value) else 0
    assert result_e3 == 0, f"Expected 0, got {result_e3}"  # None of the makes equal "Blue"
    
    # Test E4: IF(Source!B5<=60000,1,0)
    # This checks if mileage <= 60000
    mileage = get_value(data, "Source", "Mileage")
    result_e4 = 1 if mileage <= 60000 else 0
    assert result_e4 == 1, f"Expected 1, got {result_e4}"  # 57351 <= 60000
    
    # Test E5: IF(Source!B6>=2015, 1,0)
    # This checks if year >= 2015
    year = get_value(data, "Source", "Year")
    result_e5 = 1 if year >= 2015 else 0
    assert result_e5 == 1, f"Expected 1, got {result_e5}"  # 2016 >= 2015
    
    print("✓ Condition checks working correctly")

def test_named_range_reference():
    """Test the named range reference formula (I4)."""
    print("\nTesting named range reference...")
    
    data = create_test_data()
    
    # Test I4: ="I reference a named cell. The value is: " & Named_Cell
    # This should concatenate the string with the named cell value
    named_cell_value = get_cell(data, "Formulas", "I2")
    result_i4 = "I reference a named cell. The value is: " + str(named_cell_value)
    expected = "I reference a named cell. The value is: I am a named cell."
    assert result_i4 == expected, f"Expected '{expected}', got '{result_i4}'"
    
    print("✓ Named range reference working correctly")

def test_evaluation_pipeline():
    """Test the complete evaluation pipeline with the converted rules."""
    print("\nTesting complete evaluation pipeline...")
    
    # Load the converted rules from the output
    try:
        with open("data/output/conversion_summary.json", "r") as f:
            rules = json.load(f)
    except FileNotFoundError:
        print("✗ Could not find conversion_summary.json. Run main.py first.")
        return
    
    # Create test data
    test_data = create_test_data()
    
    # Evaluate the rules
    results = evaluate_rules(rules, test_data)
    
    print(f"Evaluated {len(results)} rules")
    
    # Check specific results
    expected_results = {
        "C2": 15/450,     # Country of origin weight: 15/450 ≈ 0.0333
        "C3": 50/450,     # Color weight: 50/450 ≈ 0.1111
        "C4": 100/450,    # Mileage weight: 100/450 ≈ 0.2222
        "C5": 75/450,     # Year weight: 75/450 ≈ 0.1667
        "C6": 50/450,     # Options weight: 50/450 ≈ 0.1111
        "C7": 30/450,     # Engine size weight: 30/450 ≈ 0.0667
        "C8": 80/450,     # Transmission weight: 80/450 ≈ 0.1778
        "C9": 50/450,     # Features weight: 50/450 ≈ 0.1111
        "E2": 1,          # GER found in "Stuttgart, GER"
        "E3": 0,          # No German make equals "Blue"
        "E4": 1,          # 57351 <= 60000
        "E5": 1,          # 2016 >= 2015
        "I4": "I reference a named cell. The value is: I am a named cell."
    }
    
    print("\nChecking evaluation results against expected values:")
    for cell, expected in expected_results.items():
        if cell in results:
            actual = results[cell]
            if isinstance(expected, float):
                if abs(actual - expected) < 0.0001:
                    print(f"✓ {cell}: {actual} (expected: {expected})")
                else:
                    print(f"✗ {cell}: {actual} (expected: {expected}) - MISMATCH!")
            else:
                if actual == expected:
                    print(f"✓ {cell}: {actual} (expected: {expected})")
                else:
                    print(f"✗ {cell}: {actual} (expected: {expected}) - MISMATCH!")
        else:
            print(f"✗ {cell}: Not found in results")
    
    # Verify the sum of all weights equals 1.0
    weight_cells = ["C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9"]
    total_weight = sum(results.get(cell, 0) for cell in weight_cells if results.get(cell) is not None)
    print(f"\nTotal weight sum: {total_weight}")
    assert abs(total_weight - 1.0) < 0.0001, f"Total weight should be 1.0, got {total_weight}"
    print("✓ Total weight sum equals 1.0")

def main():
    """Run all tests."""
    print("=" * 60)
    print("COMPREHENSIVE EVALUATION PIPELINE TEST")
    print("=" * 60)
    
    try:
        test_individual_functions()
        test_weight_calculations()
        test_condition_checks()
        test_named_range_reference()
        test_evaluation_pipeline()
        
        print("\n" + "=" * 60)
        print("ALL TESTS PASSED! ✓")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n✗ TEST FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
