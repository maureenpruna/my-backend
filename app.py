from flask import Flask, request, jsonify
from openpyxl import load_workbook
from io import BytesIO
import base64

app = Flask(__name__)

@app.route("/upload", methods=["POST"])
def upload_file():
    data = request.get_json()
    deal_id = data.get("deal_id")
    filename = data.get("filename")
    file_base64 = data.get("file_base64")

    if not file_base64:
        return jsonify({"error": "No file received", "success": False}), 400

    try:
        file_bytes = base64.b64decode(file_base64)
        in_memory_file = BytesIO(file_bytes)

        workbook = load_workbook(filename=in_memory_file, read_only=True)
        sheet_names = workbook.sheetnames

        return jsonify({
            "message": f"Received {filename} for Deal {deal_id}",
            "sheets": sheet_names,
            "success": True
        })
    except Exception as e:
        return jsonify({"error": f"Failed to read Excel file: {str(e)}", "success": False}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
