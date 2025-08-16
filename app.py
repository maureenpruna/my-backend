from flask import Flask, request, jsonify
import os
from werkzeug.utils import secure_filename

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
        # check if the request has the file part
        if "file" not in request.files:
            return jsonify({"error": "No file part in request"}), 400

        file = request.files["file"]
        deal_id = request.form.get("dealId")  # string param

        if not deal_id:
            return jsonify({"error": "Missing dealId"}), 400

        if file.filename == "":
            return jsonify({"error": "No file selected"}), 400

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(save_path)

            # ✅ Here you can process the Excel file (e.g., pandas openpyxl)
            # Example: read with pandas
            # import pandas as pd
            # df = pd.read_excel(save_path)

            return jsonify({
                "message": "File uploaded successfully",
                "dealId": deal_id,
                "filename": filename,
                "path": save_path
            }), 200
        else:
            return jsonify({"error": "Invalid file type, only xls/xlsx allowed"}), 400

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
