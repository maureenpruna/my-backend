from flask import Flask, request

app = Flask(__name__)

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return "No file part", 400

    file = request.files["file"]
    deal_id = request.form.get("deal_id")

    if file.filename == "":
        return "No selected file", 400

    # For now, just return confirmation
    #return f"Received {file.filename} for Deal {deal_id}"
        # Get file size in bytes
    file_size = len(file.read())
    file.seek(0)  # Reset pointer if you need to process the file later

    return f"Received {file.filename} with size {file_size} bytes for Deal {deal_id}"
