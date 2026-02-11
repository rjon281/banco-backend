import os
import pdfplumber
import pandas as pd
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import tempfile
import re
from datetime import datetime

app = Flask(__name__)
CORS(app)

@app.route('/')
def home():
    return "The Kitchen is Open - Smart Column Edition!"

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
        found_year = None
        
        # Regex to find the Year (2025, 2026, etc.)
        year_pattern = re.compile(r'\b(202[0-9])\b')
        
        # Regex to find Date at start of line (e.g., 12/18 or 01/05)
        date_start_pattern = re.compile(r'^(\d{1,2}/\d{1,2})')
        
        # Regex to find ANY money amount in the line (e.g., 1,200.50 or -50.00)
        # Handles commas, negatives at start/end
        money_pattern = re.compile(r'(-?[\d,]+\.\d{2}[-]?)')

        with pdfplumber.open(pdf_path) as pdf:
            # Step 1: Find the Year
            if len(pdf.pages) > 0:
                first_page_text = pdf.pages[0].extract_text()
                year_match = year_pattern.search(first_page_text)
                if year_match:
                    found_year = year_match.group(1)
            
            if not found_year:
                found_year = str(datetime.now().year)

            # Step 2: Extract Data
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        # Check if line looks like a transaction row
                        date_match = date_start_pattern.search(line)
                        
                        if date_match:
                            raw_date = date_match.group(1)
                            
                            # Find ALL money numbers in this line
                            money_matches = list(money_pattern.finditer(line))
                            
                            if money_matches:
                                # LOGIC: 
                                # If we have 2+ numbers, the LAST one is usually the 'Balance'.
                                # The one BEFORE the last is the 'Transaction Amount'.
                                if len(money_matches) >= 2:
                                    target_match = money_matches[-2] 
                                else:
                                    # If only 1 number, that must be the amount
                                    target_match = money_matches[0]
                                
                                raw_amount = target_match.group(1)
                                
                                # Extract Description: Text between Date and the Chosen Amount
                                description_start = date_match.end()
                                description_end = target_match.start()
                                description = line[description_start:description_end].strip()
                                
                                # Clean Amount
                                clean_amount = raw_amount.replace(',', '')
                                if clean_amount.endswith('-'):
                                    clean_amount = '-' + clean_amount[:-1]
                                
                                # Format Date
                                full_date = f"{raw_date}/{found_year}"

                                all_transactions.append({
                                    "Date": full_date,
                                    "Description": description,
                                    "Amount": float(clean_amount)
                                })

        if not all_transactions:
            return jsonify({'error': 'No transactions found. Please ensure the PDF is a text-based bank statement.'}), 400

        # Create Clean DataFrame
        df = pd.DataFrame(all_transactions)
        
        # Save to Excel
        df.to_excel(excel_path, index=False)

        return send_file(excel_path, as_attachment=True, download_name=excel_filename)

    except Exception as e:
        print(f"ERROR: {str(e)}")
        return jsonify({'error': f'Server Error: {str(e)}'}), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)
