from flask import Flask, request, jsonify
import os
import tempfile
import pandas as pd

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process_request():
    try:
        # Case 1: If dealId is sent in JSON (existing working part)
        if request.is_json:
            data = request.get_json()
            deal_id = data.get("dealId")
            if deal_id:
                # keep your existing logic for dealId
                return jsonify({"message": f"DealId {deal_id} received and processed."})

        # Case 2: If file is uploaded
        if "file" in request.files:
            file = request.files["file"]

            if file.filename == "":
                return jsonify({"error": "No file selected"}), 400

            # Save to temp file
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                file.save(tmp.name)
                tmp_path = tmp.name

            try:
                # Read Excel file
                df = pd.read_excel(tmp_path, engine="openpyxl")
                os.remove(tmp_path)  # clean up temp file

                # Example: return first 5 rows as JSON
                return jsonify({
                    "message": "Excel file processed successfully",
                    "data_preview": df.head().to_dict(orient="records")
                })

            except Exception as e:
                return jsonify({"error": f"Failed to process Excel file: {str(e)}"}), 500

        return jsonify({"error": "No dealId or file provided"}), 400

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True)
