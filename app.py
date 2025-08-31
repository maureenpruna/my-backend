from flask import Flask, request, jsonify
from openpyxl import load_workbook
from werkzeug.utils import secure_filename
import os
import requests

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

        result_data = {}

        for file in files:
            if file and allowed_file(file.filename):
                # Save the file locally
                filename = secure_filename(file.filename)
                save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(save_path)

                # Open workbook
                workbook = load_workbook(save_path, data_only=True, read_only=True)

                # --- Process Label sheet only ---
                label_sheet = workbook["Label"] if "Label" in workbook.sheetnames else None
                label_data = {"Label_T": []}
                if label_sheet:
                    row_index = 1
                    for row in label_sheet.iter_rows(min_row=2):
                        if row[0].value:  # Column A must have a value
                            entry = {
                                "Row": row_index,
                                "T1": row[0].value,
                                "T2": row[1].value if len(row) > 1 else None
                            }
                            label_data["Label_T"].append(entry)
                            row_index += 1

                result_data["Label"] = label_data

                # Send data to Zoho CRM
                crm_endpoint = "https://www.zohoapis.com.au/crm/v7/functions/parsedexceldata/actions/execute?auth_type=apikey&zapikey=1003.7aca215a9c900fecfbe589e436532a6a.560d726649a7d347aa1025a35d19c914"
                payload = {
                    "dealId": deal_id,
                    "data": result_data
                }
                headers = {"Content-Type": "application/json"}
                crm_response = requests.post(crm_endpoint, json=payload, headers=headers)

                # Delete the uploaded file
                os.remove(save_path)

            else:
                return jsonify({"error": f"Invalid file type for {file.filename}", "success": False}), 400

        return jsonify({
            "success": True,
            "dealId": deal_id,
            "message": "Files uploaded, processed, sent to CRM, and removed successfully",
            "crm_response": crm_response.json()
        })

    except Exception as e:
        return jsonify({"error": f"Failed to process Excel file: {str(e)}", "success": False}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
