from flask import Flask, request, jsonify
from openpyxl import load_workbook
from werkzeug.utils import secure_filename
import os
import requests

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

        result_data = {}

        for file in files:
            if file and allowed_file(file.filename):
                # Save the file locally
                filename = secure_filename(file.filename)
                save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(save_path)

                # Open workbook
                workbook = load_workbook(save_path, data_only=True, read_only=True)

                # --- Process Label sheet only ---
                label_sheet = workbook["Label"] if "Label" in workbook.sheetnames else None
                # Use lists of maps as requested
                label_data = {
                    "Label_T": [],   # [{T1, T2}, ...]  (A,B) include if A has value; B may be null
                    "Label_TP": [],  # [{TP}, ...]      (C)   include if C has value
                    "Label_C": [],   # [{C1, C2}, ...]  (D,E) include if D or E has value
                    "Label_L": [],   # [{L1, L2}, ...]  (F,G) include if F or G has value
                    "Label_CP": [],  # [{CP}, ...]      (H)   include if H has value
                    "Label_RT": [],  # [{RT1, RT2}, ...](I,J) include if I or J has value
                    "Label_RB": [],  # [{RB1, RB2}, ...](K,L) include if K or L has value
                    "Label_RP": []   # [{RP}, ...]      (M)   include if M has value
                }

                if label_sheet:
                    for row in label_sheet.iter_rows(min_row=2):
                        # Safely pull cell values by index
                        val = lambda idx: (row[idx].value if len(row) > idx else None)

                        A = val(0); B = val(1); C = val(2); D = val(3); E = val(4)
                        F = val(5); G = val(6); H = val(7); I = val(8); J = val(9)
                        K = val(10); L = val(11); M = val(12)

                        # Label_T: include only if A has value; B may be None
                        if A is not None and A != "":
                            label_data["Label_T"].append({"T1": A, "T2": B})

                        # Label_TP: include only if C has value
                        if C is not None and C != "":
                            label_data["Label_TP"].append({"TP": C})

                        # Label_C: include if D or E has value
                        if (D is not None and D != "") or (E is not None and E != ""):
                            label_data["Label_C"].append({"C1": D, "C2": E})

                        # Label_L: include if F or G has value
                        if (F is not None and F != "") or (G is not None and G != ""):
                            label_data["Label_L"].append({"L1": F, "L2": G})

                        # Label_CP: include if H has value
                        if H is not None and H != "":
                            label_data["Label_CP"].append({"CP": H})

                        # Label_RT: include if I or J has value
                        if (I is not None and I != "") or (J is not None and J != ""):
                            label_data["Label_RT"].append({"RT1": I, "RT2": J})

                        # Label_RB: include if K or L has value
                        if (K is not None and K != "") or (L is not None and L != ""):
                            label_data["Label_RB"].append({"RB1": K, "RB2": L})

                        # Label_RP: include if M has value
                        if M is not None and M != "":
                            label_data["Label_RP"].append({"RP": M})

                # Always include the Label key with all groups (lists may be empty)
                result_data["Label"] = label_data

                # Send data to Zoho CRM
                crm_endpoint = "https://www.zohoapis.com.au/crm/v7/functions/parsedexceldata/actions/execute?auth_type=apikey&zapikey=1003.7aca215a9c900fecfbe589e436532a6a.560d726649a7d347aa1025a35d19c914"
                payload = {
                    "dealId": deal_id,
                    "data": result_data
                }
                headers = {"Content-Type": "application/json"}
                crm_response = requests.post(crm_endpoint, json=payload, headers=headers)

                # Delete the uploaded file
                os.remove(save_path)

            else:
                return jsonify({"error": f"Invalid file type for {file.filename}", "success": False}), 400

        return jsonify({
            "success": True,
            "dealId": deal_id,
            "message": "Files uploaded, processed, sent to CRM, and removed successfully",
            "crm_response": crm_response.json()
        })

    except Exception as e:
        return jsonify({"error": f"Failed to process Excel file: {str(e)}", "success": False}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
