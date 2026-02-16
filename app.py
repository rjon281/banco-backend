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

# --- CONFIGURATION ---
FREE_PAGE_LIMIT = 10 
# ---------------------

# --- BACKEND TRACKER ---
def log_successful_conversion(filename):
    try:
        with open("conversions_log.txt", "a", encoding="utf-8") as f:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] SUCCESS: Converted '{filename}' to Excel.\n")
    except Exception as e:
        print(f"Logging failed: {e}")
# -----------------------

@app.route('/')
def home():
    return "The Kitchen is Open - LATAM Edition!"

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
        # --- THE BOUNCER (Page Limit Check) ---
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            if total_pages > FREE_PAGE_LIMIT:
                return jsonify({
                    'error': f'⚠️ El archivo tiene {total_pages} páginas. El plan gratuito permite hasta {FREE_PAGE_LIMIT} páginas. Por favor, sube un archivo más pequeño.'
                }), 400
        # --------------------------------------

        all_transactions = []
        
        start_year = datetime.now().year - 1
        end_year = datetime.now().year
        
        period_pattern = re.compile(r'period\s+(\d{1,2}/\d{1,2}/(\d{4}))\s+to\s+(\d{1,2}/\d{1,2}/(\d{4}))', re.IGNORECASE)
        date_start_pattern = re.compile(r'^(\d{1,2}/\d{1,2})')
        
        # UPGRADE 1: Broadened regex to catch spaces, commas, and periods in LATAM numbers
        money_pattern = re.compile(r'(-?(?:\d{1,3}(?:[.,\s]\d{3})*|\d+)[.,]\d{2}[-]?)')

        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) > 0:
                first_page_text = pdf.pages[0].extract_text()
                period_match = period_pattern.search(first_page_text)
                
                if period_match:
                    start_year = int(period_match.group(2))
                    end_year = int(period_match.group(4))
                    start_month = int(period_match.group(1).split('/')[0])
                else:
                    year_match = re.search(r'\b(202[0-9])\b', first_page_text)
                    if year_match:
                        start_year = int(year_match.group(1))
                        end_year = start_year + 1
                        start_month = 1
            
            original_index = 0
            
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        if "Daily Balance" in line or "Balance Detail" in line or "Saldo" in line:
                            continue

                        date_match = date_start_pattern.search(line)
                        
                        if date_match:
                            raw_date = date_match.group(1)
                            
                            # UPGRADE 2: Smart Date Handling (Assuming DD/MM for LATAM by default)
                            date_parts = raw_date.split('/')
                            # Failsafe: if the first part is > 12, it must be the day (DD/MM)
                            if int(date_parts[0]) > 12:
                                month = int(date_parts[1]) 
                            else:
                                month = int(date_parts[0]) # Fallback
                            
                            if start_month >= 10 and month < 6:
                                assigned_year = end_year
                            else:
                                assigned_year = start_year
                                
                            money_matches = list(money_pattern.finditer(line))
                            
                            if money_matches:
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
                                
                                is_balance_row = re.search(r'\d{2}/\d{2}', description)
                                
                                if description and not is_balance_row:
                                    # UPGRADE 3: The LATAM Float Cleaner
                                    clean_amount = raw_amount.replace(' ', '')
                                    if clean_amount.endswith('-'):
                                        clean_amount = '-' + clean_amount[:-1]
                                    
                                    # Detect if comma is the decimal separator (e.g., 1.000,50)
                                    if clean_amount[-3] == ',':
                                        clean_amount = clean_amount.replace('.', '').replace(',', '.')
                                    else:
                                        clean_amount = clean_amount.replace(',', '')
                                    
                                    amount_val = float(clean_amount)
                                    
                                    # UPGRADE 4: Split into Cargo and Abono
                                    cargo = abs(amount_val) if amount_val < 0 else ""
                                    abono = amount_val if amount_val > 0 else ""
                                    
                                    full_date = f"{raw_date}/{assigned_year}"

                                    all_transactions.append({
                                        "Fecha": full_date,
                                        "Concepto": description,
                                        "Cargo": cargo,
                                        "Abono": abono,
                                        "OriginalOrder": original_index,
                                        "DateForSorting": full_date
                                    })
                                    original_index += 1

        if not all_transactions:
            return jsonify({'error': 'No transactions found.'}), 400

        df = pd.DataFrame(all_transactions)
        
        # Use dayfirst=True to properly sort LATAM DD/MM/YYYY dates
        df['DateObj'] = pd.to_datetime(df['DateForSorting'], dayfirst=True, errors='coerce')
        df = df.sort_values(by=['DateObj', 'OriginalOrder'], ascending=True)
        
        # Clean up backend columns before exporting
        df = df.drop(columns=['DateObj', 'OriginalOrder', 'DateForSorting'])
        
        df.to_excel(excel_path, index=False)

        log_successful_conversion(file.filename)

        return send_file(excel_path, as_attachment=True, download_name=excel_filename)

    except Exception as e:
        print(f"ERROR: {str(e)}")
        return jsonify({'error': f'Server Error: {str(e)}'}), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)
