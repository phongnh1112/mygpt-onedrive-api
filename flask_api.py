from flask import Flask, request, jsonify
import os
from excel_user_analysis import download_excel_graph_api, analyze_user_learning

app = Flask(__name__)

@app.route("/analyze-user", methods=["POST"])
def analyze_user():
    try:
        user_code = request.json.get("user_code")
        token = os.getenv("ACCESS_TOKEN")
        if not token:
            return jsonify({"error": "ACCESS_TOKEN not found in environment"}), 401
        file_path = download_excel_graph_api(token)
        summary, _ = analyze_user_learning(file_path, user_code)
        return jsonify({"result": summary})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
