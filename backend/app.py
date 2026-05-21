import os
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  

# YOUR SPECIFIC VM FOLDER PATH
UPLOAD_FOLDER = "/home/fnc_user_4/Weekly-Dashboard/dashboard-data"
FILE_NAME = "latest_dashboard.xlsx"
FILE_PATH = os.path.join(UPLOAD_FOLDER, FILE_NAME)

@app.route('/api/upload', methods=['POST'])
def upload_file():
    print("Files tracking log:", request.files) # Prints to your terminal to help debug
    
    if not request.files:
        return jsonify({"error": "No file payload received by Flask backend"}), 400

    # Dynamically grab the first file key sent by the frontend
    file_key = list(request.files.keys())[0]
    file = request.files[file_key]
    
    if file.filename == '':
        return jsonify({"error": "Empty file name"}), 400

    try:
        # Overwrite latest_dashboard.xlsx in your folder
        file.save(FILE_PATH)
        print(f"File successfully written to: {FILE_PATH}")
        return jsonify({"message": f"Workbook saved globally inside VM directory {UPLOAD_FOLDER}!"}), 200
    except Exception as e:
        print(f"WRITE ERROR DETECTED: {str(e)}")
        return jsonify({"error": f"Failed to save to disk: {str(e)}"}), 500


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