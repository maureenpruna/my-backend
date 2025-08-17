from flask import Flask, request, jsonify
from openpyxl import load_workbook
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

@app.route("/upload", methods=["POST"])
def upload_file():
    try:
        # Get dealId from URL
        deal_id = request.args.get("dealId")
        if not deal_id:
            return jsonify({"error": "Missing dealId", "success": False}), 400

        if "file" not in request.files:
            return jsonify({"error": "No file part", "success": False}), 400

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "No file selected", "success": False}), 400

        filename = secure_filename(file.filename)
        save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(save_path)

        # Load workbook
        workbook = load_workbook(save_path, data_only=True)
        
        if "RB Label" not in workbook.sheetnames:
            return jsonify({"error": "Sheet 'RB Label' not found", "success": False}), 400

        sheet = workbook["RB Label"]

        # Identify columns
        header = [cell.value for cell in sheet[1]]  # first row as header
        try:
            job_code_idx = header.index("Job Code")
            fabric_label_idx = header.index("Fabric Label")
        except ValueError as e:
            return jsonify({"error": f"Required column missing: {str(e)}", "success": False}), 400

        # Collect values
        fabric_values = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[job_code_idx] not in (None, ""):
                fabric_values.append(row[fabric_label_idx])

        return jsonify({
            "success": True,
            "dealId": deal_id,
            "filename": filename,
            "sheet": "RB Label",
            "fabric_values": fabric_values,
            "count": len(fabric_values),
            "message": "File uploaded and processed successfully"
        })

    except Exception as e:
        return jsonify({"error": f"Failed to process Excel file: {str(e)}", "success": False}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
