from flask import Flask, render_template, request, jsonify, send_file
import os
import threading
from app.processing import generate_excel

app = Flask(__name__)

# Ensure output directory exists
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Background processing function
def process_data_in_background(stations):
    generate_excel(stations)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/allocate_slots_endpoint", methods=["POST"])
def allocate_slots_endpoint():
    try:
        stations = request.json
        
        # Start processing without threading for now
        generate_excel(stations)  # Call directly in /allocate_slots_endpoint

        return jsonify({"message": "Processing started, check back in a few seconds", "fileUrl": "/download"})
    except Exception as e:
        print(f"Error in allocation: {e}")  # Debug log
        return jsonify({"error": str(e)}), 500

@app.route("/download")
def download_file():
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], "output_kavach_slots_colored.xlsx")
    print(f"Checking if file exists: {file_path} -> {os.path.exists(file_path)}")

    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return jsonify({"message": "Final output file not yet available. Try again later."}), 202


@app.route("/upload_excel", methods=["POST"])
def upload_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not file.filename.endswith((".xls", ".xlsx")):
        return jsonify({"error": "Invalid file format"}), 400

    file_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(file_path)

    try:
        df = pd.read_excel(file_path)
        # Assuming your Excel file has columns that match your data structure,
        # e.g., "name", "stationSlots", "onboardSlots". You can rename columns as needed.
        station_data = df.to_dict(orient="records")
        return jsonify(station_data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)
