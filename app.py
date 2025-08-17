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

        # Get uploaded file
        files = list(request.files.values())
        if not files:
            return jsonify({"error": "No file received", "success": False}), 400

        file = files[0]
        if not file or not allowed_file(file.filename):
            return jsonify({"error": "Invalid file type, only xls/xlsx allowed", "success": False}), 400

        # Save the file
        filename = secure_filename(file.filename)
        save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(save_path)

        # Read Excel workbook
        workbook = load_workbook(save_path, read_only=True)
        if "RB Label" not in workbook.sheetnames:
            return jsonify({"error": "'RB Label' sheet not found", "success": False}), 400

        sheet = workbook["RB Label"]
        fabric_values = []
        for row in sheet.iter_rows(min_row=2, values_only=True):  # skip header
            job_code = row[0]  # Column A
            fabric_label = row[22] if len(row) > 22 else None  # Column W
            if job_code not in (None, "") and fabric_label not in (None, ""):
                fabric_values.append(fabric_label)

        return jsonify({
            "success": True,
            "dealId": deal_id,
            "fabric_labels_count": len(fabric_values),
            "fabric_labels": fabric_values,
            "message": "Fabric labels extracted successfully"
        })

    except Exception as e:
        return jsonify({"error": f"Failed to process Excel file: {str(e)}", "success": False}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
