from flask import Flask, render_template, jsonify, request
import json
import os
from pathlib import Path
from werkzeug.utils import secure_filename

from src.utils.scrape import extract_data_and_formulas_from_excel
from src.conversion.converter import ExcelToPythonConverter, build_dependency_graph, topological_sort
from src.conversion.rules_generator import generate_python_rules_file
from src.evaluation.evaluator import evaluate_rules

# Resolve project root (two levels up from this file: src/utils/web_ui.py -> project root)
CURRENT_FILE = Path(__file__).resolve()
PROJECT_ROOT = CURRENT_FILE.parents[2]
TEMPLATES_DIR = PROJECT_ROOT / 'templates'
STATIC_DIR = PROJECT_ROOT / 'static'

# Explicitly set template & static folders so running from any CWD works
app = Flask(__name__, template_folder=str(TEMPLATES_DIR), static_folder=str(STATIC_DIR))


def load_extracted_data():
    """Load the extracted data from JSON file."""
    # Use project root to avoid CWD issues
    data_file = PROJECT_ROOT / "extracted_data.json"
    if data_file.exists():
        with open(data_file, 'r') as f:
            return json.load(f)
    return {}


@app.route('/')
def index():
    """Main page showing the dataset viewer."""
    data = load_extracted_data()
    return render_template('index.html', data=data)


@app.route('/api/data')
def get_data():
    """API endpoint to get the data as JSON."""
    data = load_extracted_data()
    return jsonify(data)


def process_excel_file(file_path: str, include_code: bool = False, strict: bool = False):
    """Process a single Excel file path through conversion and evaluation."""
    extracted = extract_data_and_formulas_from_excel(file_path)

    all_formulas = []
    all_data = {}
    sheet_cell_to_key = {}

    for sheet_name, sheet_data in extracted.items():
        # formulas
        for formula_data in sheet_data.get("formulas", []):
            all_formulas.append({
                "formula": formula_data["formula"],
                "cell": formula_data["cell"],
                "sheet": sheet_name
            })
        # data model
        cell_values = sheet_data.get("data", {})
        key_values = sheet_data.get("key_values", {})
        all_data.setdefault(sheet_name, {})
        all_data[sheet_name].update(cell_values)
        all_data[sheet_name]['by_key'] = key_values
        # semantic map
        sheet_cell_to_key.setdefault(sheet_name, {})
        sheet_cell_to_key[sheet_name].update(sheet_data.get("cell_to_key", {}))

    converter = ExcelToPythonConverter({})
    shared_data = { 'cell_to_key_map': sheet_cell_to_key, 'strict_no_cells': strict }

    converted = []
    errors = []
    for f in all_formulas:
        try:
            conv = converter.analyze_formula(f["formula"], f["cell"], f["sheet"], shared_data)
            converted.append(conv)
        except Exception as e:
            errors.append({"cell": f["cell"], "sheet": f["sheet"], "error": str(e)})

    response = {
        "file": os.path.basename(file_path),
        "converted_count": len(converted),
        "errors": errors,
    }

    if converted:
        graph = build_dependency_graph(converted)
        try:
            order = topological_sort(graph)
        except Exception as e:
            order = []
            errors.append({"stage": "toposort", "error": str(e)})

        # rules summary
        summary = []
        for conv in converted:
            node_id = f"{conv.sheet}!{conv.cell_reference}"
            deps = list(conv.dependencies.predecessors(node_id))
            summary.append({
                "cell": conv.cell_reference,
                "sheet": conv.sheet,
                "rule_type": conv.rule_type,
                "description": conv.description,
                "original_formula": conv.original_formula,
                "python_expression": conv.python_expression,
                "dependencies": deps,
                "inputs": conv.input_keys,
                "unresolved_inputs": conv.unresolved_inputs,
            })

        
        eval_results = evaluate_rules(summary, all_data)
        for item in summary:
            item['evaluation_result'] = eval_results.get(item['cell'])

        response.update({
            "sorted_cells": order,
            "summary": summary,
        })

        if include_code:
            # Keep order alignment
            formula_map = {f"{f.sheet}!{f.cell_reference}": f for f in converted}
            sorted_formulas = [formula_map[c] for c in order if c in formula_map]
            code = generate_python_rules_file(converter, sorted_formulas, shared_data, order)
            response["generated_code"] = code

    return response


@app.route('/api/convert', methods=['POST'])
def convert_endpoint():
    """Accept an Excel upload, process it, and return conversion results as JSON."""
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    filename = secure_filename(file.filename)
    upload_dir = PROJECT_ROOT / 'data' / 'uploads'
    upload_dir.mkdir(parents=True, exist_ok=True)
    file_path = str(upload_dir / filename)
    file.save(file_path)

    include_code = request.args.get('include_code') in ('1', 'true', 'True')
    strict = request.args.get('strict') in ('1', 'true', 'True')

    result = process_excel_file(file_path, include_code=include_code, strict=strict)
    return jsonify(result)


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)