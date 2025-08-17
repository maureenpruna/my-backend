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

        fabric_labels_all = []

        for file in files:
            if file and allowed_file(file.filename):
                # Save the file locally
                filename = secure_filename(file.filename)
                save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(save_path)

                # Open workbook with data_only=True to get calculated values
                workbook = load_workbook(save_path, data_only=True)
                
                if "RB Label" in workbook.sheetnames:
                    sheet = workbook["RB Label"]
                    fabric_labels = []

                    for row in sheet.iter_rows(min_row=2):  # start from row 2
                        job_code = row[0].value  # Column A
                        if job_code:  # only if column A is not empty
                            cell_w = row[22].value  # Column W (0-indexed)
                            fabric_labels.append(cell_w)

                    fabric_labels_all.extend(fabric_labels)

            else:
                continue

        return jsonify({
            "success": True,
            "dealId": deal_id,
            "fabric_labels": fabric_labels_all,
            "fabric_labels_count": len(fabric_labels_all),
            "message": "Fabric labels extracted successfully"
        })

    except Exception as e:
        return jsonify({"error": f"Failed to process Excel file: {str(e)}", "success": False}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
