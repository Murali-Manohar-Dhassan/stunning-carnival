from flask import Flask, render_template, request, jsonify, send_file
import os
import threading
from app.processing import generate_excel

app = Flask(__name__)

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
        thread = threading.Thread(target=process_data_in_background, args=(stations,))
        thread.start()

        return jsonify({"message": "Processing started, check back in a few seconds", "fileUrl": "/download"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/download")
def download_file():
    file_path = os.path.join(os.getcwd(), "output_kavach_slots_colored.xlsx")

    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return jsonify({"error": "Final output file not yet available. Try again later."}), 404

if __name__ == "__main__":
    app.run(debug=True)
