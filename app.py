from flask import Flask, request, jsonify
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"  # Folder on the server where files are saved
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "No file received"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if file and allowed_file(file.filename):
        # Save file locally with fixed name
        ext = file.filename.rsplit(".", 1)[1].lower()
        filename = f"localfile.{ext}"
        save_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(save_path)

        return jsonify({"success": True, "message": f"File saved as {filename}", "path": save_path})
    else:
        return jsonify({"error": "Invalid file type"}), 400

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
