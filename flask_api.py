from flask import Flask, request, jsonify
from excel_user_analysis import get_access_token, download_excel_graph_api, analyze_user_learning

app = Flask(__name__)

@app.route("/")
def index():
    return "✅ API đang hoạt động. Dùng POST /analyze-user với user_code để phân tích."

@app.route("/analyze-user", methods=["POST"])
def analyze_user():
    try:
        user_code = request.json.get("user_code")
        print("🔍 User code nhận được:", user_code)

        access_token = get_access_token()
        file_path = download_excel_graph_api(access_token)
        print("📁 File tải về:", file_path)

        summary, _ = analyze_user_learning(file_path, user_code)
        return jsonify({"result": summary})
    except Exception as e:
        print("❌ Lỗi trong analyze_user():", str(e))
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
