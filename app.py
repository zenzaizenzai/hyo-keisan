from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import json
import os

app = Flask(__name__, static_folder='static')
CORS(app)

DATA_FILE = 'data.json'

@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/<path:path>')
def static_proxy(path):
    return send_from_directory(app.static_folder, path)

@app.route('/api/load', methods=['GET'])
def load_data():
    if not os.path.exists(DATA_FILE):
        # Initial empty data: 20 rows x 10 columns
        initial_data = [["" for _ in range(10)] for _ in range(20)]
        return jsonify(initial_data)
    
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        return jsonify(json.load(f))

@app.route('/api/save', methods=['POST'])
def save_data():
    data = request.json
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return jsonify({"status": "success"})

if __name__ == '__main__':
    app.run(port=5000, debug=True)
