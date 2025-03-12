from flask import Flask, render_template, request, jsonify, send_file
import os
import threading
from app.processing import generate_excel

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)  # Ensure the folder exists
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
        
        # Start processing in a new thread
        '''
        thread = threading.Thread(target=process_data_in_background, args=(stations,), daemon=True)
        thread.start()'''
        generate_excel(stations)  # Call directly in /allocate_slots_endpoint

        return jsonify({"message": "Processing started, check back in a few seconds", "fileUrl": "/download"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/download")
def download_file():
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], "output_kavach_slots_colored.xlsx")
    print(f"Checking if file exists: {file_path} -> {os.path.exists(file_path)}")

    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return jsonify({"message": "Final output file not yet available. Try again later."}), 202

if __name__ == "__main__":
    app.run(debug=True)
