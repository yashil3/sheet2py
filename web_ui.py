from flask import Flask, render_template, jsonify
import json
import os
from pathlib import Path

app = Flask(__name__)

def load_extracted_data():
    """Load the extracted data from JSON file."""
    data_file = Path("extracted_data.json")
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

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)