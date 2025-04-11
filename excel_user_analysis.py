import os
import json
import requests
import pandas as pd
from flask import Flask, request, jsonify
from dotenv import load_dotenv
from typing import Tuple
import msal

load_dotenv()

app = Flask(__name__)

CLIENT_ID = "2bdb3693-4837-4cc6-9f60-ea3858985b16"
TENANT_ID = "e5039572-eed3-431f-92a3-6c3dd04c34fb"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["Files.Read", "offline_access"]
EXCEL_PATH_ON_ONEDRIVE = "/Documents/0.App/KẾT_QUẢ_LUYỆN_TẬP_AI.xlsx"

# === Tự động lấy access_token, ưu tiên làm mới từ refresh_token ===
def get_access_token():
    token = os.getenv("ACCESS_TOKEN")
    refresh_token = os.getenv("REFRESH_TOKEN")

    if token:
        test = requests.get(
            "https://graph.microsoft.com/v1.0/me",
            headers={"Authorization": f"Bearer {token}"}
        )
        if test.status_code == 200:
            return token

    if refresh_token:
        app_auth = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
        result = app_auth.acquire_token_by_refresh_token(refresh_token, scopes=SCOPE)
        if "access_token" in result:
            # Ghi lại token mới ra file .env (nếu cần)
            os.environ["ACCESS_TOKEN"] = result["access_token"]
            return result["access_token"]

    raise EnvironmentError("⚠️ Token không hợp lệ và không làm mới được từ REFRESH_TOKEN.")

# === Tải file Excel từ OneDrive ===
def download_excel_graph_api(access_token: str, save_path: str = "data.xlsx") -> str:
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:{EXCEL_PATH_ON_ONEDRIVE}:/content"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    with open(save_path, "wb") as f:
        f.write(response.content)
    return save_path

# === Phân tích dữ liệu người dùng ===
def analyze_user_learning(path: str, user_code: str) -> Tuple[str, pd.DataFrame]:
    df = pd.read_excel(path, engine='openpyxl')
    df.columns = df.columns.str.strip()
    user_df = df[df[df.columns[5]] == user_code]  # Cột F
    if user_df.empty:
        return f"Không tìm thấy dữ liệu cho user '{user_code}'", pd.DataFrame()

    user_df = user_df[[df.columns[2], df.columns[6], df.columns[7]]]  # C, G, H
    user_df.columns = ['Completion Time', 'Practice Today', 'Result']

    analysis = [f"🔎 Phân tích cho user: {user_code}\n"]
    for _, row in user_df.iterrows():
        analysis.append(f"- 📅 {row['Completion Time'].strftime('%d/%m/%Y %H:%M')}: luyện \"{row['Practice Today']}\" → kết quả: \"{row['Result']}\"")

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
    user_code = request.json.get("user_code")
    token = get_access_token()
    file_path = download_excel_graph_api(token)
    summary, _ = analyze_user_learning(file_path, user_code)
    return jsonify({"result": summary})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
