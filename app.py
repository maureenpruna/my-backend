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

                # --- Process RB Label ---
                rb_sheet = workbook["RB Label"] if "RB Label" in workbook.sheetnames else None
                rb_data = {"Fabric Label": [], "Fabric Label count": 0,
                           "Bottom Label": [], "Bottom Label count": 0}
                if rb_sheet:
                    for row in rb_sheet.iter_rows(min_row=2):
                        if row[0].value:  # Column A is Job Code
                            # Column W -> Fabric Label (index 22)
                            if len(row) > 22:
                                rb_data["Fabric Label"].append(row[22].value)
                            # Column X -> Bottom Label (index 23)
                            if len(row) > 23:
                                rb_data["Bottom Label"].append(row[23].value)
                    rb_data["Fabric Label count"] = len(rb_data["Fabric Label"])
                    rb_data["Bottom Label count"] = len(rb_data["Bottom Label"])
                result_data["RB"] = rb_data

                # --- Process CTN Label ---
                ctn_sheet = workbook["CTN Label"] if "CTN Label" in workbook.sheetnames else None
                ctn_data = {"Fabric Label": [], "Fabric Label count": 0,
                            "Bottom Label": [], "Bottom Label count": 0}
                if ctn_sheet:
                    for row in ctn_sheet.iter_rows(min_row=2):
                        if row[0].value:  # Column A is Job Code
                            if len(row) > 22:
                                ctn_data["Fabric Label"].append(row[22].value)
                            if len(row) > 23:
                                ctn_data["Bottom Label"].append(row[23].value)
                    ctn_data["Fabric Label count"] = len(ctn_data["Fabric Label"])
                    ctn_data["Bottom Label count"] = len(ctn_data["Bottom Label"])
                result_data["CTN"] = ctn_data

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
