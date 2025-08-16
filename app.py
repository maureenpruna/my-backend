from flask import Flask, request, jsonify
import os
from werkzeug.utils import secure_filename
from openpyxl import load_workbook  # ✅ to read Excel files

app = Flask(__name__)

# where to temporarily save uploads
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/upload", methods=["POST"])
def upload_file():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file part in request"}), 400

        file = request.files["file"]
        deal_id = request.form.get("dealId")

        if not deal_id:
            return jsonify({"error": "Missing dealId"}), 400

        if file.filename == "":
            return jsonify({"error": "No file selected"}), 400

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(save_path)

            # ✅ Open the Excel workbook and extract sheet names
            workbook = load_workbook(filename=save_path, read_only=True)
            sheet_names = workbook.sheetnames

            return jsonify({
                "message": "File uploaded and processed successfully",
                "dealId": deal_id,
                "filename": filename,
                "sheets": sheet_names,
                "size_bytes": os.path.getsize(save_path)
            }), 200
        else:
            return jsonify({"error": "Invalid file type, only xls/xlsx allowed"}), 400

    except Exception as e:
        return jsonify({"error": f"Failed to process Excel file: {str(e)}"}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
