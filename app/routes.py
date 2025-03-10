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
        file_path = generate_excel(allocations)
        return jsonify({"fileUrl": f"/download/{file_path}"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/download/<path:filename>")
def download_file(filename):
    return send_file(filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)