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
    return "The Kitchen is Open - Final Perfection Edition!"

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
        
        # Regex patterns
        year_pattern = re.compile(r'\b(202[0-9])\b')
        date_start_pattern = re.compile(r'^(\d{1,2}/\d{1,2})')
        money_pattern = re.compile(r'(-?[\d,]+\.\d{2}[-]?)')

        with pdfplumber.open(pdf_path) as pdf:
            # Step 1: Find the Start Year
            if len(pdf.pages) > 0:
                first_page_text = pdf.pages[0].extract_text()
                year_match = year_pattern.search(first_page_text)
                if year_match:
                    found_year = int(year_match.group(1))
            
            if not found_year:
                found_year = datetime.now().year

            # Step 2: Extract Data
            current_year = found_year
            last_month = 0

            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        
                        # Stop processing if we hit the "Daily Balance" section
                        if "Daily Balance" in line or "Balance Detail" in line:
                            continue

                        date_match = date_start_pattern.search(line)
                        
                        if date_match:
                            raw_date = date_match.group(1)
                            month = int(raw_date.split('/')[0])

                            # YEAR ROLLOVER LOGIC:
                            # If we go from Month 12 to Month 01, increment the year
                            if last_month == 12 and month == 1:
                                current_year += 1
                            
                            # Update last_month seen
                            if month > 0: 
                                last_month = month

                            money_matches = list(money_pattern.finditer(line))
                            
                            if money_matches:
                                # HYBRID LOGIC (Same as before)
                                first_money = money_matches[0]
                                date_end_index = date_match.end()
                                money_start_index = first_money.start()
                                
                                if money_start_index - date_end_index < 10:
                                    target_match = first_money
                                    raw_amount = target_match.group(1)
                                    description = line[target_match.end():].strip()
                                else:
                                    if len(money_matches) >= 2:
                                        target_match = money_matches[-2]
                                    else:
                                        target_match = money_matches[0]
                                    
                                    raw_amount = target_match.group(1)
                                    description = line[date_end_index:target_match.start()].strip()
                                
                                # FILTER: Ignore lines that look like Balance Summaries
                                # (Balance lines usually have date+amount but NO description, or just numbers)
                                is_balance_row = re.search(r'\d{2}/\d{2}', description) # If desc contains a date, it's likely a balance row
                                
                                if description and not is_balance_row:
                                    clean_amount = raw_amount.replace(',', '')
                                    if clean_amount.endswith('-'):
                                        clean_amount = '-' + clean_amount[:-1]
                                    
                                    full_date = f"{raw_date}/{current_year}"

                                    all_transactions.append({
                                        "Date": full_date,
                                        "Description": description,
                                        "Amount": float(clean_amount)
                                    })

        if not all_transactions:
            return jsonify({'error': 'No transactions found.'}), 400

        df = pd.DataFrame(all_transactions)
        df.to_excel(excel_path, index=False)

        return send_file(excel_path, as_attachment=True, download_name=excel_filename)

    except Exception as e:
        print(f"ERROR: {str(e)}")
        return jsonify({'error': f'Server Error: {str(e)}'}), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)
