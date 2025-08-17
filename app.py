from flask import Flask, request, jsonify
from openpyxl import load_workbook
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/upload", methods=["POST"])
def upload_file():
    try:
        # Get dealId from URL
        deal_id = request.args.get("dealId")
        if not deal_id:
            return jsonify({"error": "Missing dealId", "success": False}), 400

        # Collect all uploaded files regardless of key
        files = list(request.files.values())
        if not files:
            return jsonify({"error": "No file received", "success": False}), 400

        response_files = []

        for file in files:
            if file and allowed_file(file.filename):
                # Save the file locally
                filename = secure_filename(file.filename)
                save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(save_path)

                # Read Excel and get sheet names
                workbook = load_workbook(save_path, data_only=True)
                sheet_names = workbook.sheetnames

                file_response = {
                    "filename": filename,
                    "sheets": sheet_names
                }

                # Look specifically for "RB Label"
                if "RB Label" in sheet_names:
                    sheet = workbook["RB Label"]
                    header = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]

                    try:
                        job_code_idx = header.index("Job Code")
                        fabric_label_idx = header.index("Fabric Label")
                    except ValueError:
                        # Column missing
                        fabric_values = []
                    else:
                        # Collect Fabric Label values where Job Code is not empty
                        fabric_values = [
                            row[fabric_label_idx]
                            for row in sheet.iter_rows(min_row=2, values_only=True)
                            if row[job_code_idx] not in (None, "")
                        ]

                    file_response["fabric_values"] = fabric_values
                    file_response["fabric_count"] = len(fabric_values)

                response_files.append(file_response)

            else:
                response_files.append({
                    "filename": file.filename,
                    "error": "Invalid file type, only xls/xlsx allowed"
                })

        return jsonify({
            "success": True,
            "dealId": deal_id,
            "files": response_files,
            "message": "Files uploaded and processed successfully"
        })

    except Exception as e:
        return jsonify({"error": f"Failed to process Excel file: {str(e)}", "success": False}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
