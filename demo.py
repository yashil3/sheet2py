#!/usr/bin/env python3
"""
Demo launcher for BAH RuleBuilder
This script properly sets up the Python path and starts the Flask demo server.
"""
import os
import sys
from pathlib import Path

# Add the project root to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# Import and run the Flask app
from src.utils.web_ui import app

if __name__ == '__main__':
    print("Starting BAH RuleBuilder Demo Server...")
    print("=" * 50)
    print("Web Interface: http://localhost:5000")
    print("API Endpoint: http://localhost:5000/api/convert")
    print("Upload Excel files to see real-time conversion!")
    print("=" * 50)
    print("\nPress Ctrl+C to stop the server")
    
    app.run(debug=True, host='0.0.0.0', port=5000) 