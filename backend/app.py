import os
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  

# YOUR SPECIFIC VM FOLDER PATH
UPLOAD_FOLDER = "/home/user/Weekly-Dashboard/dashboard-data"
FILE_NAME = "latest_dashboard.xlsx"
FILE_PATH = os.path.join(UPLOAD_FOLDER, FILE_NAME)

# ==========================================
# WRITING LOGIC: Saves the uploaded file here
# ==========================================
@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file payload chunk received"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "Empty filename selection error"}), 400

    
    file.save(FILE_PATH)
    return jsonify({"message": f"Workbook saved globally inside VM directory {UPLOAD_FOLDER}!"}), 200


# ==========================================
# READING LOGIC: Fetches and streams data on refresh
# ==========================================
@app.route('/api/file', methods=['GET'])
def get_file():
    if os.path.exists(FILE_PATH):
        
        return send_from_directory(UPLOAD_FOLDER, FILE_NAME, as_attachment=True)
    else:
        return jsonify({"error": "No operational worksheet template uploaded to server memory yet"}), 404


if __name__ == '__main__':
   
    app.run(host='0.0.0.0', port=5000)