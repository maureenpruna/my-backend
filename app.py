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

def extract_labels(sheet):
    """Extract Fabric and Bottom labels from the given sheet (Columns W & X), only rows with Column A not empty."""
    fabric_labels = []
    bottom_labels = []
    for row in sheet.iter_rows(min_row=2):
        job_code = row[0].value  # Column A
        if job_code:
            fabric_labels.append(row[22].value)  # Column W
            bottom_labels.append(row[23].value)  # Column X
    return fabric_labels, bottom_labels

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
                filename = secure_filename(file.filename)
                save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(save_path)

                workbook = load_workbook(save_path, data_only=True)

                # RB Labels
                if "RB Label" in workbook.sheetnames:
                    rb_sheet = workbook["RB Label"]
                    rb_fabric, rb_bottom = extract_labels(rb_sheet)
                else:
                    rb_fabric, rb_bottom = [], []

                # CTN Labels
                if "CTN Label" in workbook.sheetnames:
                    ctn_sheet = workbook["CTN Label"]
                    ctn_fabric, ctn_bottom = extract_labels(ctn_sheet)
                else:
                    ctn_fabric, ctn_bottom = [], []

                response_files.append({
                    "filename": filename,
                    "RB": {
                        "Fabric Label": rb_fabric,
                        "Fabric Label count": len(rb_fabric),
                        "Bottom Label": rb_bottom,
                        "Bottom Label count": len(rb_bottom)
                    },
                    "CTN": {
                        "Fabric Label": ctn_fabric,
                        "Fabric Label count": len(ctn_fabric),
                        "Bottom Label": ctn_bottom,
                        "Bottom Label count": len(ctn_bottom)
                    }
                })
            else:
                response_files.append({
                    "filename": file.filename,
                    "error": "Invalid file type, only xls/xlsx allowed"
                })

        return jsonify({
            "success": True,
            "dealId": deal_id,
            "files": response_files,
            "message": "Labels extracted successfully"
        })

    except Exception as e:
        return jsonify({"error": f"Failed to process Excel file: {str(e)}", "success": False}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
