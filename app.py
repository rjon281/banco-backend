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
    return "The Kitchen is Open - Period Aware Edition!"

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
        
        # Default years (fallback)
        start_year = datetime.now().year - 1
        end_year = datetime.now().year
        
        # Regex to find the precise period in the header
        # Looks for "For the period 12/16/2025 to 01/16/2026"
        period_pattern = re.compile(r'period\s+(\d{1,2}/\d{1,2}/(\d{4}))\s+to\s+(\d{1,2}/\d{1,2}/(\d{4}))', re.IGNORECASE)
        
        # Regex for transaction data
        date_start_pattern = re.compile(r'^(\d{1,2}/\d{1,2})')
        money_pattern = re.compile(r'(-?[\d,]+\.\d{2}[-]?)')

        with pdfplumber.open(pdf_path) as pdf:
            # Step 1: Detect the Statement Period from Page 1
            if len(pdf.pages) > 0:
                first_page_text = pdf.pages[0].extract_text()
                period_match = period_pattern.search(first_page_text)
                
                if period_match:
                    # We found the exact dates!
                    # Group 2 is Start Year, Group 4 is End Year
                    start_year = int(period_match.group(2))
                    end_year = int(period_match.group(4))
                    start_month = int(period_match.group(1).split('/')[0])
                else:
                    # Fallback: Find just a single year (e.g. 2025)
                    year_match = re.search(r'\b(202[0-9])\b', first_page_text)
                    if year_match:
                        start_year = int(year_match.group(1))
                        end_year = start_year + 1 # Assume max 1 year rollover
                        start_month = 1 # Fallback
            
            # Step 2: Extract Data
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        # Skip Balance Summary lines
                        if "Daily Balance" in line or "Balance Detail" in line:
                            continue

                        date_match = date_start_pattern.search(line)
                        
                        if date_match:
                            raw_date = date_match.group(1)
                            month = int(raw_date.split('/')[0])
                            
                            # SMART YEAR LOGIC:
                            # If the transaction month is >= the statement start month, it's the Start Year.
                            # Otherwise (e.g., Month 01 < Month 12), it's the End Year.
                            # Example: Period is Dec (12) to Jan (01).
                            # Row is 12/17 -> 12 >= 12 -> 2025
                            # Row is 01/05 -> 01 < 12 -> 2026
                            
                            # Note: This simple logic handles 99% of statements. 
                            # For safety, we check if start_month is late in the year (>= 10) 
                            # to assume a rollover is possible.
                            if start_month >= 10 and month < 6:
                                assigned_year = end_year
                            else:
                                assigned_year = start_year
                                
                            money_matches = list(money_pattern.finditer(line))
                            
                            if money_matches:
                                # HYBRID COLUMN DETECTION
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
                                
                                # Filter Balance Rows
                                is_balance_row = re.search(r'\d{2}/\d{2}', description)
                                
                                if description and not is_balance_row:
                                    clean_amount = raw_amount.replace(',', '')
                                    if clean_amount.endswith('-'):
                                        clean_amount = '-' + clean_amount[:-1]
                                    
                                    full_date = f"{raw_date}/{assigned_year}"

                                    all_transactions.append({
                                        "Date": full_date,
                                        "Description": description,
                                        "Amount": float(clean_amount)
                                    })

        if not all_transactions:
            return jsonify({'error': 'No transactions found.'}), 400

        df = pd.DataFrame(all_transactions)
        # Sort by date to fix the "jumping" order in the final Excel
        df['DateObj'] = pd.to_datetime(df['Date'], format='%m/%d/%Y')
        df = df.sort_values(by='DateObj')
        df = df.drop(columns=['DateObj'])
        
        df.to_excel(excel_path, index=False)

        return send_file(excel_path, as_attachment=True, download_name=excel_filename)

    except Exception as e:
        print(f"ERROR: {str(e)}")
        return jsonify({'error': f'Server Error: {str(e)}'}), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)
