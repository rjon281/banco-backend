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
    return "The Kitchen is Open - Universal Import Edition!"

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
        
        # Regex to find a year (looks for 2024, 2025, 2026, etc.)
        year_pattern = re.compile(r'\b(202[0-9])\b')
        
        # Regex to find the Date at the START of a line (MM/DD)
        date_start_pattern = re.compile(r'^(\d{1,2}/\d{1,2})')
        
        # Regex to find the Amount at the END of a line
        # Looks for numbers with a decimal, possibly a minus sign or 'CR' at the end
        # Examples: 25.17, 1,000.00, -50.00, 25.17-
        amount_end_pattern = re.compile(r'(-?[\d,]+\.\d{2}[-]?)\s*$')

        with pdfplumber.open(pdf_path) as pdf:
            # Step 1: Find the Year from the first page
            if len(pdf.pages) > 0:
                first_page_text = pdf.pages[0].extract_text()
                year_match = year_pattern.search(first_page_text)
                if year_match:
                    found_year = year_match.group(1)
            
            # Use current year if we couldn't find one
            if not found_year:
                found_year = str(datetime.now().year)

            # Step 2: Extract Transactions
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        # Does line start with a date?
                        date_match = date_start_pattern.search(line)
                        
                        if date_match:
                            raw_date = date_match.group(1)
                            
                            # Does line end with an amount?
                            amount_match = amount_end_pattern.search(line)
                            
                            if amount_match:
                                raw_amount = amount_match.group(1)
                                
                                # The Description is everything IN BETWEEN the Date and the Amount
                                # We slice the string using the lengths of the matches
                                description = line[date_match.end():amount_match.start()].strip()
                                
                                # Clean up the amount (remove commas, handle trailing negatives)
                                clean_amount = raw_amount.replace(',', '')
                                if clean_amount.endswith('-'):
                                    clean_amount = '-' + clean_amount[:-1]
                                
                                # Format Date with Year
                                full_date = f"{raw_date}/{found_year}"

                                all_transactions.append({
                                    "Date": full_date,
                                    "Description": description,
                                    "Amount": float(clean_amount)
                                })

        if not all_transactions:
            return jsonify({'error': 'No transactions found. Try downloading a fresh PDF from the bank.'}), 400

        # Create DataFrame with the 3 Universal Columns
        df = pd.DataFrame(all_transactions)
        
        # Save to Excel
        df.to_excel(excel_path, index=False)

        return send_file(excel_path, as_attachment=True, download_name=excel_filename)

    except Exception as e:
        print(f"ERROR: {str(e)}")
        return jsonify({'error': f'Server Error: {str(e)}'}), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)
