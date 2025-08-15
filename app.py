from flask import Flask, request

app = Flask(__name__)

@app.route("/", methods=["GET"])
def home():
    return "Hello, your backend is running!"

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return "No file part", 400

    file = request.files["file"]
    if file.filename == "":
        return "No selected file", 400

    # Log confirmation to the server logs
    print(f"Received file: {file.filename}, size: {len(file.read())} bytes")
    file.seek(0)  # Reset file pointer if you want to process it later

    return f"We have received {file.filename}"
