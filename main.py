import json
import time
import logging
from tqdm import tqdm
import os

from src.conversion.converter import ExcelToPythonConverter, build_dependency_graph, topological_sort, ConvertedFormula
from src.conversion.rules_generator import generate_python_rules_file
from src.evaluation.evaluator import evaluate_rules
from src.utils.scrape import extract_data_and_formulas_from_excel

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')


def main():
    """Main conversion function."""
    start_time = time.time()

    input_dir = 'data/input'
    output_dir = 'data/output'
    os.makedirs(output_dir, exist_ok=True)

    all_formulas = []
    all_data = {}
    # For semantic mapping
    sheet_cell_to_key = {}

    for filename in os.listdir(input_dir):
        if filename.endswith(('.xlsx', '.xlsm')):
            file_path = os.path.join(input_dir, filename)
            logging.info(f"Extracting data and formulas from {filename}...")
            extracted_data = extract_data_and_formulas_from_excel(file_path)
            for sheet_name, sheet_data in extracted_data.items():
                for formula_data in sheet_data["formulas"]:
                    all_formulas.append({
                        "formula": formula_data["formula"],
                        "cell": formula_data["cell"],
                        "sheet": sheet_name
                    })
                # Build normalized data structure
                cell_values = sheet_data.get("data", {})
                key_values = sheet_data.get("key_values", {})
                all_data.setdefault(sheet_name, {})
                # Put flat cell values at top-level for backward compatibility
                all_data[sheet_name].update(cell_values)
                # Also attach by_key mapping used by get_value
                all_data[sheet_name]['by_key'] = key_values

                # Build cell->key map for this sheet
                sheet_cell_to_key.setdefault(sheet_name, {})
                sheet_cell_to_key[sheet_name].update(sheet_data.get("cell_to_key", {}))

    converter = ExcelToPythonConverter({})
    # Provide shared mappings to the visitor
    strict_flag = os.getenv('STRICT_NO_CELLS', '0') in ('1', 'true', 'True')
    shared_data = {
        'cell_to_key_map': sheet_cell_to_key,
        'strict_no_cells': strict_flag,
    }
    converted_formulas = []

    with tqdm(total=len(all_formulas), desc="Converting Formulas") as pbar:
        for formula_data in all_formulas:
            try:
                converted = converter.analyze_formula(
                    formula_data["formula"],
                    formula_data["cell"],
                    formula_data["sheet"],
                    shared_data
                )
                converted_formulas.append(converted)
                logging.info(f"✓ Converted {formula_data['cell']}: {formula_data['formula']}")
            except Exception as e:
                logging.error(f"✗ Error converting {formula_data['cell']}: {e}")
            pbar.update(1)

    print(f"\nSuccessfully converted {len(converted_formulas)} formulas")

    if converted_formulas:
        dependency_graph = build_dependency_graph(converted_formulas)

        try:
            sorted_cells = topological_sort(dependency_graph)
            print("✓ Formulas topologically sorted.")
        except ValueError as e:
            print(f"✗ Error: {e}")
            return

        formula_map = {f"{f.sheet}!{f.cell_reference}": f for f in converted_formulas}
        sorted_formulas = [formula_map[cell] for cell in sorted_cells if cell in formula_map]

        python_code = generate_python_rules_file(converter, sorted_formulas, shared_data, sorted_cells)

        output_file = os.path.join(output_dir, "converted_rules.py")
        with open(output_file, 'w') as f:
            f.write(python_code)
        print(f"✓ Generated Python rules file: {output_file}")

        summary = []
        for conv in converted_formulas:
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

        with open(os.path.join(output_dir, "conversion_summary.json"), 'w') as f:
            json.dump(summary, f, indent=2)
        print(f"✓ Generated conversion summary: {os.path.join(output_dir, "conversion_summary.json")}")

        # Evaluate the rules
        evaluation_results = evaluate_rules(summary, all_data)

        # Append evaluation results to the summary
        for item in summary:
            item['evaluation_result'] = evaluation_results.get(item['cell'])

        with open(os.path.join(output_dir, "conversion_summary.json"), 'w') as f:
            json.dump(summary, f, indent=2)
        print(f"✓ Appended evaluation results to conversion summary.")

    end_time = time.time()
    print(f"Total execution time: {end_time - start_time:.2f} seconds")


if __name__ == "__main__":
    main()