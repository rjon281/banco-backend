import os
import pdfplumber
import pandas as pd
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import tempfile
import re

app = Flask(__name__)
CORS(app)

@app.route('/')
def home():
    return "The Kitchen is Open - Advanced PNC Edition!"

@app.route('/convert', methods=['POST'])
def convert_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    temp_dir = tempfile.gettempdir()
    pdf_path = os.path.join(temp_dir, file.filename)
    excel_filename = file.filename.replace('.pdf', '.xlsx')
    excel_path = os.path.join(temp_dir, excel_filename)
    
    file.save(pdf_path)

    try:
        all_data = []
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                # STRATEGY 1: Look for explicit grid lines (Standard)
                tables = page.extract_tables({
                    "vertical_strategy": "lines", 
                    "horizontal_strategy": "lines"
                })
                
                # STRATEGY 2: If that fails, look for "Text Columns" (PNC Style)
                if not tables:
                    tables = page.extract_tables({
                        "vertical_strategy": "text", 
                        "horizontal_strategy": "text",
                        "snap_tolerance": 4,
                    })

                # Process whatever tables we found
                if tables:
                    for table in tables:
                        for row in table:
                            # Clean up the row data
                            clean_row = [str(cell).strip() if cell else '' for cell in row]
                            # Only keep rows that look like transactions (have a date or amount)
                            if any(clean_row): 
                                all_data.append(clean_row)
                
                # STRATEGY 3: Emergency Text Extraction (If tables fail completely)
                else:
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            # Simple logic: split by large spaces
                            parts = re.split(r'\s{3,}', line) 
                            if len(parts) > 1:
                                all_data.append(parts)

        if not all_data:
            return jsonify({'error': 'Could not extract any data. The PDF might be encrypted or have a unique layout.'}), 400

        # Create DataFrame and clean it
        df = pd.DataFrame(all_data)
        
        # Save to Excel
        df.to_excel(excel_path, index=False, header=False)

        return send_file(excel_path, as_attachment=True, download_name=excel_filename)

    except Exception as e:
        print(f"ERROR: {str(e)}") # Print to Render logs for debugging
        return jsonify({'error': f'Server Error: {str(e)}'}), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)
