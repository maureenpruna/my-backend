from flask import Flask, request
from openpyxl import load_workbook
import io
import base64

app = Flask(__name__)

@app.route("/upload", methods=["POST"])
def upload_file():
    data = request.get_json()

    if not data or "file_base64" not in data:
        return "No file data received", 400

    file_base64 = data["file_base64"]
    file_bytes = base64.b64decode(file_base64)

    deal_id = data.get("deal_id", "Unknown")
    filename = data.get("filename", "Unknown.xlsx")

    # Read Excel in memory
    in_memory_file = io.BytesIO(file_bytes)

    try:
        workbook = load_workbook(filename=in_memory_file, read_only=True)
        sheet_names = workbook.sheetnames
    except Exception as e:
        return f"Failed to read Excel file: {str(e)}", 400

    return f"Received {filename} with {len(file_bytes)} bytes for Deal {deal_id}. Sheets: {sheet_names}"
