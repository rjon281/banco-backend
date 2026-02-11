import os
import pdfplumber
import pandas as pd
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import tempfile

app = Flask(__name__)
CORS(app) 

@app.route('/')
def home():
    return "The Kitchen is Open and Resilient!"

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
            for page in pdf.pages:
                try:
                    # Improved table extraction settings
                    tables = page.extract_tables({
                        "vertical_strategy": "lines", 
                        "horizontal_strategy": "lines",
                        "snap_tolerance": 3,
                    })
                    
                    # If that fails, try a more aggressive text-based strategy
                    if not tables:
                        tables = page.extract_tables()

                    if tables:
                        for table in tables:
                            for row in table:
                                if any(row): # Only add if row isn't empty
                                    clean_row = [str(cell).strip() if cell is not None else '' for cell in row]
                                    all_data.append(clean_row)
                except Exception as page_err:
                    print(f"Skipping a messy page: {page_err}")
                    continue

        if not all_data:
            return jsonify({'error': 'Could not find any data tables in this PDF. Is it a scanned image?'}), 400

        # Convert to Excel
        df = pd.DataFrame(all_data)
        df.to_excel(excel_path, index=False, header=False)

        return send_file(excel_path, as_attachment=True, download_name=excel_filename)

    except Exception as e:
        return jsonify({'error': f'Server Error: {str(e)}'}), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)
