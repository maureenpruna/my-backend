from flask import Flask, request, jsonify
import os
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {"xlsx", "xls"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/upload', methods=['POST'])
def upload_file():
    # Get dealId from URL query string
    deal_id = request.args.get('dealId')
    if not deal_id:
        return jsonify({"error": "Missing dealId", "success": False}), 400

    # Check if a file is included
    if 'file' not in request.files:
        return jsonify({"error": "No file part", "success": False}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({"error": "No selected file", "success": False}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Invalid file type, only xls/xlsx allowed", "success": False}), 400

    # Save the file temporarily
    filename = secure_filename(file.filename)
    file_path = os.path.join(UPLOAD_FOLDER, f"{deal_id}_{filename}")
    file.save(file_path)

    try:
        # Read Excel file and get sheet names
        workbook = load_workbook(file_path, read_only=True)
        sheet_names = workbook.sheetnames

        return jsonify({
            "success": True,
            "dealId": deal_id,
            "filename": filename,
            "sheets": sheet_names,
            "message": "File uploaded and processed successfully"
        })

    except Exception as e:
        return jsonify({"error": f"Failed to process Excel file: {str(e)}", "success": False}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
