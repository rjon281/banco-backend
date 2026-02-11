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
    return "The Kitchen is Open - Hybrid Layout Edition!"

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
        
        # Regex to find the Year
        year_pattern = re.compile(r'\b(202[0-9])\b')
        
        # Regex to find Date at start (e.g., 12/18)
        date_start_pattern = re.compile(r'^(\d{1,2}/\d{1,2})')
        
        # Regex to find Money (e.g., 1,200.50 or -50.00)
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

            # Step 2: Extract Transactions
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        date_match = date_start_pattern.search(line)
                        
                        if date_match:
                            raw_date = date_match.group(1)
                            money_matches = list(money_pattern.finditer(line))
                            
                            if money_matches:
                                # HYBRID LOGIC: Check where the FIRST money amount is
                                first_money = money_matches[0]
                                date_end_index = date_match.end()
                                money_start_index = first_money.start()
                                
                                # If money starts within 10 characters of the date, 
                                # Assume Format: Date -> Amount -> Description
                                if money_start_index - date_end_index < 10:
                                    target_match = first_money
                                    raw_amount = target_match.group(1)
                                    # Description is everything AFTER the amount
                                    description = line[target_match.end():].strip()
                                
                                # Otherwise, Assume Format: Date -> Description -> Amount
                                else:
                                    # Pick the amount (handling the Balance column logic from before)
                                    if len(money_matches) >= 2:
                                        target_match = money_matches[-2]
                                    else:
                                        target_match = money_matches[0]
                                    
                                    raw_amount = target_match.group(1)
                                    # Description is everything BETWEEN Date and Amount
                                    description = line[date_end_index:target_match.start()].strip()
                                
                                # Clean & Save
                                if description: # Filter out empty lines (like Balance summaries)
                                    clean_amount = raw_amount.replace(',', '')
                                    if clean_amount.endswith('-'):
                                        clean_amount = '-' + clean_amount[:-1]
                                    
                                    full_date = f"{raw_date}/{found_year}"

                                    all_transactions.append({
                                        "Date": full_date,
                                        "Description": description,
                                        "Amount": float(clean_amount)
                                    })

        if not all_transactions:
            return jsonify({'error': 'No transactions found. Layout might be unique.'}), 400

        df = pd.DataFrame(all_transactions)
        df.to_excel(excel_path, index=False)

        return send_file(excel_path, as_attachment=True, download_name=excel_filename)

    except Exception as e:
        print(f"ERROR: {str(e)}")
        return jsonify({'error': f'Server Error: {str(e)}'}), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)
