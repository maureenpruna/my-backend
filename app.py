from flask import Flask, request, jsonify
import tempfile
import os
from openpyxl import load_workbook

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Ensure both dealId and file are present
        deal_id = request.form.get("dealId")
        if not deal_id:
            return jsonify({"success": False, "error": "Missing dealId"}), 400

        if "file" not in request.files:
            return jsonify({"success": False, "error": "Missing file"}), 400

        uploaded_file = request.files["file"]

        # Save file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            uploaded_file.save(tmp.name)
            tmp_path = tmp.name

        # Load workbook
        wb = load_workbook(tmp_path, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()

        # Cleanup temp file
        os.remove(tmp_path)

        return jsonify({
            "success": True,
            "dealId": deal_id,
            "sheets": sheet_names
        })

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
