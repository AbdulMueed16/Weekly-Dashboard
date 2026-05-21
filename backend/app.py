import os
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

app = Flask(__name__)
# Enable CORS so your React frontend can cleanly communicate with this API
CORS(app)

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
FILE_NAME = 'latest_dashboard.xlsx'
FILE_PATH = os.path.join(UPLOAD_FOLDER, FILE_NAME)

# Ensure the uploads directory exists on the VM
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Endpoint 1: Stream and overwrite the file on the VM
@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file chunk found"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "Empty filename selection"}), 400

    # Save the file under a static name to overwrite older metrics
    file.save(FILE_PATH)
    return jsonify({"message": "Spreadsheet updated globally on Nginx Server VM!"}), 200

# Endpoint 2: Fetch the active server spreadsheet array buffer
@app.route('/api/file', methods=['GET'])
def get_file():
    if os.path.exists(FILE_PATH):
        return send_from_directory(UPLOAD_FOLDER, FILE_NAME, as_attachment=True)
    else:
        return jsonify({"error": "No spreadsheet uploaded to server storage yet"}), 404

if __name__ == '__main__':
    # Running backend app on port 5000
    app.run(host='0.0.0.0', port=5000)