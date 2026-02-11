import os
import pdfplumber
import pandas as pd
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import tempfile

app = Flask(__name__)
# This allows your Hostinger website to talk to this server
CORS(app)

@app.route('/')
def home():
    return "The Kitchen is Open! (Backend is running)"

@app.route('/convert', methods=['POST'])
def convert_pdf():
    # 1. Basic Validation
    if 'file' not in request.files:
        return jsonify({'error': 'No file part found'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    # 2. Save PDF temporarily
    temp_dir = tempfile.gettempdir()
    pdf_path = os.path.join(temp_dir, file.filename)
    excel_filename = file.filename.replace('.pdf', '.xlsx')
    excel_path = os.path.join(temp_dir, excel_filename)

    file.save(pdf_path)

    try:
        # 3. The Conversion Magic
        all_data = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            # Clean data (remove 'None' values)
                            clean_row = [str(cell) if cell is not None else '' for cell in row]
                            all_data.append(clean_row)

        if not all_data:
            return jsonify({'error': 'No readable tables found in this PDF.'}), 400

        # 4. Create Excel
        df = pd.DataFrame(all_data)
        df.to_excel(excel_path, index=False, header=False)

        # 5. Send back to User
        return send_file(excel_path, as_attachment=True, download_name=excel_filename)

    except Exception as e:
        return jsonify({'error': f'Internal Error: {str(e)}'}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
