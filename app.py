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

        # Check if any files are uploaded
        if not request.files:
            return jsonify({"error": "No files received", "success": False}), 400

        files = request.files.getlist("file")  # get all uploaded files
        result = []

        for file in files:
            if file.filename == "":
                continue  # skip empty filenames

            filename = secure_filename(file.filename)
            save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(save_path)

            # Read Excel and get sheet names
            try:
                workbook = load_workbook(save_path, read_only=True)
                sheet_names = workbook.sheetnames
            except Exception as e:
                sheet_names = []
                return jsonify({"error": f"Failed to process Excel file {filename}: {str(e)}", "success": False}), 500

            result.append({
                "filename": filename,
                "sheets": sheet_names
            })

        return jsonify({
            "success": True,
            "dealId": deal_id,
            "files": result,
            "message": "Files uploaded and processed successfully"
        })

    except Exception as e:
        return jsonify({"error": f"Server error: {str(e)}", "success": False}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
