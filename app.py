from flask import Flask, request
from openpyxl import load_workbook
import io

app = Flask(__name__)

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return "No file part", 400

    file = request.files["file"]
    deal_id = request.form.get("deal_id")

    if file.filename == "":
        return "No selected file", 400

    # Get file size
    file.stream.seek(0, 2)  # move to end
    file_size = file.stream.tell()
    file.stream.seek(0)  # reset pointer

    # Read the Excel file
    in_memory_file = io.BytesIO(file.read())
    workbook = load_workbook(filename=in_memory_file, read_only=True)
    sheet_names = workbook.sheetnames  # List of tab names

    return f"Received {file.filename} with size {file_size} bytes for Deal {deal_id}. Tabs: {', '.join(sheet_names)}"

if __name__ == "__main__":
    app.run(debug=True)
