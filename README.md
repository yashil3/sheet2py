# sheet2py

**Convert Excel formulas into maintainable Python rules with advanced parsing, evaluation, and web interface.**

Sheet2Py is a comprehensive tool that transforms complex Excel spreadsheets into clean, auditable Python code. It features advanced formula parsing using ANTLR, dependency resolution, rule evaluation, and both CLI and web interfaces for seamless integration into your workflow.


## Project Structure

```
sheet2py/
├── README.md
├── LICENSE                     # MIT License
├── requirements.txt            # Python dependencies
├── main.py                     # CLI entry point
├── demo.py                     # Web UI launcher
├── extracted_data.json         # Cached extraction results
├── data/
│   ├── input/                 # Input Excel files (.xlsx, .xlsm)
│   ├── output/                # Generated Python rules and summaries
│   └── uploads/               # Web UI file uploads
├── src/
│   ├── antlr_files/           # ANTLR grammar and generated parsers
│   │   ├── ExcelFormula.g4    # Grammar definition
│   │   ├── ExcelFormulaLexer.py
│   │   ├── ExcelFormulaParser.py
│   │   └── FormulaConverterVisitor.py
│   ├── conversion/            # Core conversion logic
│   │   ├── converter.py       # Main converter with ANTLR integration
│   │   ├── excel_functions.py # Excel function implementations
│   │   ├── rules_generator.py # Python code generation
│   │   └── batch_process.py   # Batch processing utilities
│   ├── evaluation/            # Rule evaluation and validation
│   │   └── evaluator.py       # Formula accuracy testing
│   └── utils/                 # Utilities and interfaces
│       ├── scrape.py          # Excel data extraction
│       └── web_ui.py          # Flask web interface
├── static/                    # Web UI assets
│   ├── styles.css
│   └── script.js
├── templates/                 # HTML templates
│   └── index.html
└── tests/                     # Test suite
    ├── test_converter.py      # Conversion tests
    └── test_key_mapping.py    # Mapping validation tests
```

## Quick Start

### Option 1: Command Line Interface

1. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Add Excel Files:**
   Place your Excel files (`.xlsx`, `.xlsm`) in the `data/input/` directory.

3. **Run Conversion:**
   ```bash
   python main.py
   ```

4. **View Results:**
   - **Python rules:** `data/output/converted_rules.py`
   - **Conversion summary:** `data/output/conversion_summary.json`
   - **Evaluation results:** included in the summary with accuracy metrics

### Option 2: Web Interface

1. **Start the Web Server:**
   ```bash
   python demo.py
   ```

2. **Open Your Browser:**
   Navigate to `http://localhost:5000`

3. **Upload & Convert:**
   - Drag and drop Excel files
   - View real-time conversion progress
   - Explore data interactively
   - Download generated Python rules

### Option 3: API Integration

Use the REST API for programmatic access:

```bash
# Convert a file via API
curl -X POST http://localhost:5000/api/convert \
  -F "file=@your_spreadsheet.xlsx" \
  -F "include_code=true" \
  -F "strict=false"
```

## Configuration

### Environment Variables

- `STRICT_NO_CELLS=1`: Enable strict mode (no direct cell references in output)

### Advanced Options

The converter supports various configuration options through the shared data parameter:

```python
shared_data = {
    'cell_to_key_map': {},      # Custom cell-to-variable mappings
    'strict_no_cells': False,   # Strict mode flag
}
```

## Supported Excel Functions

### Mathematical Functions
- `SUM`, `SUMIF`, `SUMIFS`
- `COUNT`, `COUNTIF`, `COUNTIFS`
- `AVERAGE`, `MIN`, `MAX`
- `ROUND`, `ABS`, `SQRT`

### Logical Functions
- `IF`, `AND`, `OR`, `NOT`
- `IFERROR`, `ISBLANK`, `ISNUMBER`

### Lookup Functions
- `VLOOKUP`, `INDEX`, `MATCH`
- `XLOOKUP` (partial support)

### Text Functions
- `CONCATENATE`, `LEFT`, `RIGHT`, `MID`
- `LEN`, `TRIM`, `UPPER`, `LOWER`

### Date Functions
- `TODAY`, `NOW`, `DATE`
- `YEAR`, `MONTH`, `DAY`

## Testing

Run the comprehensive test suite:

```bash
# Run all tests
pytest tests/

# Run specific test files
pytest tests/test_converter.py
pytest tests/test_key_mapping.py

# Run with verbose output
pytest tests/ -v

# Run with coverage
pytest tests/ --cov=src/
```

## Output Examples

### Generated Python Rules

```python
# Auto-generated from Excel formulas
def calculate_total_score():
    return sum_keys(['product_quality', 'service_rating', 'value_score'])

def determine_recommendation():
    total = calculate_total_score()
    if total >= 80:
        return "Highly Recommended"
    elif total >= 60:
        return "Recommended"
    else:
        return "Not Recommended"
```

### Conversion Summary

```json
{
  "cell": "D5",
  "sheet": "Analysis",
  "rule_type": "calculation",
  "description": "Sum of quality metrics",
  "original_formula": "=SUM(A5:C5)",
  "python_expression": "sum_keys(['product_quality', 'service_rating', 'value_score'])",
  "dependencies": ["A5", "B5", "C5"],
  "evaluation_result": {
    "accuracy": 100.0,
    "excel_result": 85,
    "python_result": 85
  }
}
```

### Adding New Excel Functions

1. **Implement the function** in `src/conversion/excel_functions.py`:
   ```python
   def my_excel_function(arg1, arg2):
       """Implementation of MYEXCELFUNCTION"""
       return result
   ```

2. **Add to the visitor** in `src/antlr_files/FormulaConverterVisitor.py`:
   ```python
   def visit_function_call(self, ctx):
       if function_name.upper() == 'MYEXCELFUNCTION':
           return f"my_excel_function({args})"
   ```

3. **Write tests** in `tests/test_converter.py`





## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
