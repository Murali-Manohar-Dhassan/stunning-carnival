from flask import Flask, render_template, request, jsonify, send_file
import os
from app.processing import generate_excel

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/allocate_slots_endpoint", methods=["POST"])
def allocate_slots_endpoint():
    try:
        stations = request.json
        final_excel_path = generate_excel(stations)  # Now generates final output file

        return jsonify({"fileUrl": "/download"})  # Return correct download path
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/download")
def download_file():
    file_path = os.path.join(os.getcwd(), "output_kavach_slots_colored.xlsx")  # Ensure correct file path

    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return jsonify({"error": "Final output file not found"}), 404

if __name__ == "__main__":
    app.run(debug=True)
