# BAH RuleBuilder

This project converts Excel formulas into Python rules, providing a more maintainable and auditable way to manage complex business logic.

## Project Structure

```
BAH_RuleBuilder/
├── .gitignore
├── README.md
├── requirements.txt
├── main.py             # Main entry point for the application
├── config.yaml         # (Optional) Configuration file for future enhancements
├── data/
│   ├── input/          # Place your Excel files here
│   └── output/         # Converted Python rules and summary will be saved here
├── src/
│   ├── __init__.py
│   ├── antlr_files/    # ANTLR grammar and generated parser/lexer files
│   │   ├── __init__.py
│   │   ├── ExcelFormula.g4
│   │   └── ... (ANTLR generated files)
│   ├── conversion/     # Core logic for formula conversion
│   │   ├── __init__.py
│   │   ├── converter.py
│   │   ├── excel_functions.py
│   │   └── rules_generator.py
│   └── utils/          # Utility functions (e.g., Excel scraping)
│       ├── __init__.py
│       └── scrape.py
└── tests/
    ├── __init__.py
    └── test_converter.py
```

## Features


### Formula Conversion
The system converts Excel formulas to Python expressions while maintaining:
- Mathematical operations and functions
- Logical operations (IF, AND, OR, etc.)
- Excel-specific functions (VLOOKUP, INDEX, etc.)
- Cell references and ranges
- Named range references

## Workflow

1.  **Place Input Files:** Put your Excel files (e.g., `.xlsx`, `.xlsm`) into the `data/input/` directory.

2.  **Install Dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

3.  **Run Conversion:** Execute the `main.py` script from your terminal:

    ```bash
    python main.py
    ```

4.  **Get Output:**
    *   The converted Python rules will be saved in `data/output/converted_rules.py`.
    *   A summary of the conversion process will be available in `data/output/conversion_summary.json`.

## Testing

To run the tests, navigate to the project root and execute:

```bash
pytest tests/
```
