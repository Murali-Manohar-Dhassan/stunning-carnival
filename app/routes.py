from flask import Flask, render_template, request, jsonify, send_file
import os
from app.processing import allocate_slots, generate_excel

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/allocate_slots_endpoint", methods=["POST"])
def allocate_slots_endpoint():
    try:
        stations = request.json
        allocations = allocate_slots(stations)
        file_path = generate_excel(allocations)  # Ensure this returns the correct path
        # Correct file URL
        return jsonify({"fileUrl": "/download"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/download")
def download_file():
    file_path = os.path.join(os.getcwd(), "slot_allocation.xlsx")  # Correct file path

    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return jsonify({"error": "File not found"}), 404

if __name__ == "__main__":
    app.run(debug=True)
