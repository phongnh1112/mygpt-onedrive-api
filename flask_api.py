from flask import Flask, request, jsonify
from excel_user_analysis import get_access_token, download_excel_graph_api, analyze_user_learning

app = Flask(__name__)

@app.route("/analyze-user", methods=["POST"])
def analyze_user():
    user_code = request.json.get("user_code")
    token = get_access_token()
    file_path = download_excel_graph_api(token)
    summary, _ = analyze_user_learning(file_path, user_code)
    return jsonify({"result": summary})
