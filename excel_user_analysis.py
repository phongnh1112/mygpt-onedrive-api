import os
import json
import requests
import pandas as pd
from flask import Flask, request, jsonify
from dotenv import load_dotenv
from typing import Tuple

load_dotenv()

app = Flask(__name__)

EXCEL_PATH_ON_ONEDRIVE = "/0.App/KẾT_QUẢ_LUYỆN_TẬP_AI.xlsx"
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
REFRESH_TOKEN = os.getenv("REFRESH_TOKEN")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")

def refresh_access_token() -> str:
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'grant_type': 'refresh_token',
        'refresh_token': REFRESH_TOKEN,
        'scope': 'https://graph.microsoft.com/.default'
    }
    response = requests.post(url, data=data, headers=headers)
    response.raise_for_status()
    new_token = response.json()["access_token"]
    os.environ["ACCESS_TOKEN"] = new_token
    return new_token

def download_excel_graph_api(access_token: str, save_path: str = "data.xlsx") -> str:
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:{EXCEL_PATH_ON_ONEDRIVE}:/content"
    response = requests.get(url, headers=headers)

    if response.status_code == 401:
        access_token = refresh_access_token()
        headers["Authorization"] = f"Bearer {access_token}"
        response = requests.get(url, headers=headers)

    response.raise_for_status()
    with open(save_path, "wb") as f:
        f.write(response.content)
    return save_path

def analyze_user_learning(path: str, user_code: str) -> Tuple[str, pd.DataFrame]:
    df = pd.read_excel(path, engine='openpyxl')
    df.columns = df.columns.str.strip()

    user_code = user_code.strip().lower()
    col_user = "User của bạn là?"

    if col_user not in df.columns:
        return f"Không tìm thấy cột '{col_user}' trong dữ liệu.", pd.DataFrame()

    df[col_user] = df[col_user].astype(str).str.strip().str.lower()
    df = df.dropna(subset=[col_user])

    user_df = df[df[col_user] == user_code]

    if user_df.empty:
        return f"Không tìm thấy dữ liệu cho user '{user_code}'", pd.DataFrame()

    user_df = user_df[["Completion time", "Bài luyện tập hôm nay của bạn là?", "Kết quả bài luyện tập là?"]]
    user_df.columns = ['Completion Time', 'Practice Today', 'Result']

    analysis = [f"🔎 Phân tích cho user: {user_code}\n"]
    for _, row in user_df.iterrows():
        time_fmt = row['Completion Time'].strftime('%d/%m/%Y %H:%M') if not pd.isnull(row['Completion Time']) else "Không rõ thời gian"
        analysis.append(f"- 📅 {time_fmt}: luyện \"{row['Practice Today']}\" → kết quả: \"{row['Result']}\"")

    if any('không' in str(x).lower() or 'chưa' in str(x).lower() or 'chịu' in str(x).lower() for x in user_df['Result']):
        suggestion = "📌 Gợi ý: Nên ôn lại bài tập cũ hoặc bắt đầu từ kiến thức nền tảng."
    else:
        suggestion = "📌 Gợi ý: Bạn đang học tốt, hãy chuyển sang bài học AI nâng cao tiếp theo."

    analysis.append(suggestion)
    return "\n".join(analysis), user_df

@app.route("/")
def home():
    return "✅ API đang hoạt động. Dùng POST /analyze-user với user_code để phân tích."

@app.route("/analyze-user", methods=["POST"])
def analyze_user():
    try:
        data = request.get_json(silent=True) or {}
        user_code = data.get("user_code")

        if not user_code or not isinstance(user_code, str) or not user_code.strip():
            return jsonify({"error": "Thiếu hoặc sai định dạng user_code. Vui lòng gửi user_code hợp lệ."}), 400

        user_code = user_code.strip().lower()
        print("🔍 User code nhận được:", user_code)

        file_path = download_excel_graph_api(os.getenv("ACCESS_TOKEN"))
        print("📁 File tải về:", file_path)

        summary, _ = analyze_user_learning(file_path, user_code)
        return jsonify({"result": summary})
    except Exception as e:
        print("❌ Lỗi trong analyze_user():", str(e))
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
