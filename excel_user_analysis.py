import os
import requests
import pandas as pd
from flask import Flask, request, jsonify
from dotenv import load_dotenv
from typing import Tuple

load_dotenv()

app = Flask(__name__)

EXCEL_PATH_ON_ONEDRIVE = "/Documents/0.App/KẾT_QUẢ_LUYỆN_TẬP_AI.xlsx"

# === Tải access token từ biến môi trường ===
def get_access_token():
    token = os.getenv("ACCESS_TOKEN")
    if not token:
        raise EnvironmentError("⚠️ ACCESS_TOKEN chưa được thiết lập trong biến môi trường.")
    return token

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
