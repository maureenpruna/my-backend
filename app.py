from flask import Flask, request, jsonify
import os
import tempfile
import pandas as pd

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Case 1: dealId in JSON (optional, if you want to support JSON POST)
        if request.is_json:
            data = request.get_json()
            deal_id = data.get("dealId")
            if deal_id:
                return jsonify({"message": f"DealId {deal_id} received and processed."})

        # Case 2: File upload with dealId as form-data
        if "file" in request.files:
            file = request.files["file"]
            deal_id = request.form.get("dealId")

            if file.filename == "":
                return jsonify({"error": "No file selected"}), 400

            # Save to temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                file.save(tmp.name)
                tmp_path = tmp.name

            try:
                # Load Excel file and get sheet names
                excel_file = pd.ExcelFile(tmp_path, engine="openpyxl")
                sheet_names = excel_file.sheet_names

                os.remove(tmp_path)  # clean up

                return jsonify({
                    "message": "Excel file processed successfully",
                    "dealId": deal_id,
                    "filename": file.filename,
                    "sheets": sheet_names
                })

            except Exception as e:
                return jsonify({"error": f"Failed to process Excel file: {str(e)}"}), 500

        return jsonify({"error": "No dealId or file provided"}), 400

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
