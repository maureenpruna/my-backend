from flask import Flask, request, jsonify
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Allow only Excel file extensions
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/process_excel', methods=['POST'])
def process_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']

    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Invalid file type. Only .xls or .xlsx allowed"}), 400

    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join("/tmp", filename)  # temp save
        file.save(filepath)

        excel_data = pd.ExcelFile(filepath)

        result = {}
        for sheet in excel_data.sheet_names:
            df = pd.read_excel(filepath, sheet_name=sheet)
            result[sheet] = df.to_dict(orient='records')

        os.remove(filepath)  # cleanup temp file

        return jsonify({"sheets": result})

    except Exception as e:
        return jsonify({"error": f"Failed to process Excel file: {str(e)}"}), 500


@app.route('/')
def home():
    return "Excel Processing API is running!"


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
