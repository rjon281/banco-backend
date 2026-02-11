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
    return "The Kitchen is Open - Date Filter Edition!"

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
        all_transactions = []
        
        # Regex pattern for dates like 12/18, 01/05, 12/18/2025
        # It looks for: 1 or 2 digits, a slash, 1 or 2 digits
        date_pattern = re.compile(r'^\d{1,2}/\d{1,2}')

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Extract text using a layout-preserving method
                text = page.extract_text()
                
                if text:
                    for line in text.split('\n'):
                        # Check if the line STARTS with a date
                        if date_pattern.match(line):
                            # It's a transaction! 
                            # Split the line by large spaces (2 or more spaces) to separate columns
                            parts = re.split(r'\s{2,}', line)
                            all_transactions.append(parts)

        if not all_transactions:
            return jsonify({'error': 'No transactions found. Does your statement have dates like MM/DD?'}), 400

        # Create DataFrame
        # We don't know the exact column names, so we let pandas guess or leave them blank
        df = pd.DataFrame(all_transactions)
        
        # Save to Excel
        df.to_excel(excel_path, index=False, header=["Date", "Description/Data", "Amount", "Balance", "Misc"][:len(df.columns)])

        return send_file(excel_path, as_attachment=True, download_name=excel_filename)

    except Exception as e:
        print(f"ERROR: {str(e)}")
        return jsonify({'error': f'Server Error: {str(e)}'}), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)
