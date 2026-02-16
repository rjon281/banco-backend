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
FREE_PAGE_LIMIT = 15 
ALLOWED_EXTENSIONS = {'pdf', 'csv', 'xls', 'xlsx'}
# ---------------------

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def log_successful_conversion(filename):
    try:
        with open("conversions_log.txt", "a", encoding="utf-8") as f:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] SUCCESS: Converted '{filename}' to Excel.\n")
    except Exception as e:
        print(f"Logging failed: {e}")

# --- SMART CLEAN ENGINE ---
def smart_clean_vendor(text):
    """Strips bank jargon, card numbers, and transaction IDs to leave a clean vendor name."""
    if not isinstance(text, str):
        return text
    
    desc = text.upper()
    # Remove common bank filler words
    prefixes = [r'\bCOMPRA\b', r'\bTARJETA\b', r'\bDEBITO\b', r'\bSPEI\b', r'\bTRANSFERENCIA\b', 
                r'\bPAGO\b', r'\bCAJERO\b', r'\bRETIRO\b', r'\bDEPOSITO\b', r'\bMOVIMIENTO\b']
    for p in prefixes:
        desc = re.sub(p, '', desc)
    
    # Remove card numbers (e.g., *1234, xxxx-1234)
    desc = re.sub(r'[*#xX-]*\d{4,}\b', '', desc)
    # Remove long reference numbers (5 or more digits)
    desc = re.sub(r'\b\d{5,}\b', '', desc)
    # Clean up extra spaces
    desc = re.sub(r'\s+', ' ', desc).strip()
    
    return desc if desc else "DESCONOCIDO / UNKNOWN"

@app.route('/')
def home():
    return "The Kitchen is Open - Global Edition with CSV Support!"

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '' or not allowed_file(file.filename):
        return jsonify({'error': 'No selected file or unsupported format. Please upload PDF, CSV, or Excel.'}), 400

    temp_dir = tempfile.gettempdir()
    file_ext = file.filename.rsplit('.', 1)[1].lower()
    
    # Save incoming file
    input_path = os.path.join(temp_dir, file.filename)
    file.save(input_path)
    
    # Output file path
    excel_filename = file.filename.rsplit('.', 1)[0] + '_BancoAXLS.xlsx'
    excel_path = os.path.join(temp_dir, excel_filename)

    try:
        # ==========================================
        # PIPELINE 1: CSV / EXCEL PROCESSING
        # ==========================================
        if file_ext in ['csv', 'xls', 'xlsx']:
            if file_ext == 'csv':
                # Try UTF-8 first, fallback to Latin-1 (Very common in LATAM/Spain)
                try:
                    df = pd.read_csv(input_path, encoding='utf-8')
                except UnicodeDecodeError:
                    df = pd.read_csv(input_path, encoding='latin1')
            else:
                df = pd.read_excel(input_path)

            # Auto-detect Description column to run the Smart Cleaner
            desc_cols = [col for col in df.columns if any(keyword in str(col).lower() for keyword in ['concepto', 'description', 'descripción', 'detalle', 'motivo'])]
            
            if desc_cols:
                target_col = desc_cols[0]
                df['Clean_Vendor (Auto-Cleaned)'] = df[target_col].apply(smart_clean_vendor)

            df.to_excel(excel_path, index=False)
            log_successful_conversion(file.filename)
            return send_file(excel_path, as_attachment=True, download_name=excel_filename)

        # ==========================================
        # PIPELINE 2: PDF PROCESSING
        # ==========================================
        elif file_ext == 'pdf':
            with pdfplumber.open(input_path) as pdf:
                total_pages = len(pdf.pages)
                if total_pages > FREE_PAGE_LIMIT:
                    return jsonify({
                        'error': f'⚠️ El archivo tiene {total_pages} páginas. El plan gratuito permite hasta {FREE_PAGE_LIMIT}. / File exceeds free limit.'
                    }), 400

                # Extract text for Auto-Detection
                full_text = ""
                for page in pdf.pages[:min(3, total_pages)]:
                    extracted = page.extract_text()
                    if extracted:
                        full_text += extracted + "\n"
                
                # --- AUTO-DETECT LANGUAGE ---
                text_lower = full_text.lower()
                es_words = ['fecha', 'concepto', 'saldo', 'cargo', 'abono', 'retiro', 'depósito', 'movimiento']
                en_words = ['date', 'description', 'balance', 'amount', 'withdrawal', 'deposit', 'summary']
                
                is_spanish = sum(text_lower.count(w) for w in es_words) > sum(text_lower.count(w) for w in en_words)
                
                # --- AUTO-DETECT DECIMAL FORMAT ---
                dot_matches = len(re.findall(r'\.\d{2}(?:\s|$|-)', full_text))
                comma_matches = len(re.findall(r',\d{2}(?:\s|$|-)', full_text))
                is_comma_decimal = comma_matches > dot_matches
                
                if is_comma_decimal:
                    money_pattern = re.compile(r'(-?(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}[-]?)')
                else:
                    money_pattern = re.compile(r'(-?(?:\d{1,3}(?:,\d{3})*|\d+)\.\d{2}[-]?)')

                period_pattern = re.compile(r'period\s+(\d{1,2}/\d{1,2}/(\d{4}))\s+to\s+(\d{1,2}/\d{1,2}/(\d{4}))', re.IGNORECASE)
                date_start_pattern = re.compile(r'^(\d{1,2}/\d{1,2})')
                
                start_year = datetime.now().year - 1
                end_year = datetime.now().year
                start_month = 1

                if len(pdf.pages) > 0:
                    first_page_text = full_text
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

                all_transactions = []
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
                                
                                date_parts = raw_date.split('/')
                                month = int(date_parts[1]) if int(date_parts[0]) > 12 else int(date_parts[0])
                                
                                assigned_year = end_year if (start_month >= 10 and month < 6) else start_year
                                    
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
                                        target_match = money_matches[-2] if len(money_matches) >= 2 else money_matches[0]
                                        raw_amount = target_match.group(1)
                                        description = line[date_end_index:target_match.start()].strip()
                                    
                                    is_balance_row = re.search(r'\d{2}/\d{2}', description)
                                    
                                    if description and not is_balance_row:
                                        clean_amount = raw_amount.replace(' ', '')
                                        if clean_amount.endswith('-'):
                                            clean_amount = '-' + clean_amount[:-1]
                                        
                                        if is_comma_decimal:
                                            clean_amount = clean_amount.replace('.', '').replace(',', '.')
                                        else:
                                            clean_amount = clean_amount.replace(',', '')
                                        
                                        try:
                                            amount_val = float(clean_amount)
                                        except ValueError:
                                            continue 
                                        
                                        full_date = f"{raw_date}/{assigned_year}"
                                        cleaned_vendor = smart_clean_vendor(description)
                                        
                                        if is_spanish:
                                            cargo = abs(amount_val) if amount_val < 0 else ""
                                            abono = amount_val if amount_val > 0 else ""
                                            all_transactions.append({
                                                "Fecha": full_date,
                                                "Concepto": description,
                                                "Comercio (Limpio)": cleaned_vendor,
                                                "Cargo": cargo,
                                                "Abono": abono,
                                                "OriginalOrder": original_index,
                                                "DateForSorting": full_date
                                            })
                                        else:
                                            all_transactions.append({
                                                "Date": full_date,
                                                "Description": description,
                                                "Clean Vendor": cleaned_vendor,
                                                "Amount": amount_val,
                                                "OriginalOrder": original_index,
                                                "DateForSorting": full_date
                                            })
                                            
                                        original_index += 1

                if not all_transactions:
                    return jsonify({'error': 'No transactions found.'}), 400

                df = pd.DataFrame(all_transactions)
                df['DateObj'] = pd.to_datetime(df['DateForSorting'], dayfirst=is_spanish, errors='coerce')
                df = df.sort_values(by=['DateObj', 'OriginalOrder'], ascending=True)
                df = df.drop(columns=['DateObj', 'OriginalOrder', 'DateForSorting'])
                
                df.to_excel(excel_path, index=False)
                log_successful_conversion(file.filename)
                return send_file(excel_path, as_attachment=True, download_name=excel_filename)

    except Exception as e:
        print(f"ERROR: {str(e)}")
        return jsonify({'error': f'Server Error: {str(e)}'}), 500
    finally:
        if os.path.exists(input_path): os.remove(input_path)
