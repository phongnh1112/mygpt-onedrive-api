import os
import msal
import requests
import pandas as pd
from typing import Tuple

# === Thiết lập thông tin truy cập nội bộ ===
CLIENT_ID = "04f0c124-f2bc-4f3a-83f7-1e29a3b8c6a4"  # Microsoft public client ID
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPE = ["Files.Read"]
EXCEL_PATH_ON_ONEDRIVE = "/Documents/0.App/KẾT_QUẢ_LUYỆN_TẬP_AI.xlsx"  # Đường dẫn file gốc trong OneDrive cá nhân của bạn

# === Lấy access token bằng device code flow (phù hợp môi trường server) ===
def get_access_token():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    flow = app.initiate_device_flow(scopes=SCOPE)
    if "user_code" not in flow:
        raise Exception("Không khởi tạo được device code flow.")
    print(f"🔑 Vui lòng truy cập {flow['verification_uri']} và nhập mã: {flow['user_code']}")
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Không lấy được access token.")

# === Tải file Excel từ OneDrive cá nhân (Microsoft Graph API) ===
def download_excel_graph_api(access_token: str, save_path: str = "data.xlsx") -> str:
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:{EXCEL_PATH_ON_ONEDRIVE}:/content"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    with open(save_path, "wb") as f:
        f.write(response.content)
    return save_path

# === Phân tích dữ liệu user từ file Excel ===
def analyze_user_learning(path: str, user_code: str) -> Tuple[str, pd.DataFrame]:
    df = pd.read_excel(path, engine='openpyxl')
    df.columns = df.columns.str.strip()

    user_df = df[df[df.columns[5]] == user_code]  # Cột F: User
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

# === Chạy thử ===
if __name__ == "__main__":
    token = get_access_token()
    file_path = download_excel_graph_api(token)

    user_input = "phongnh9"
    summary, user_data = analyze_user_learning(file_path, user_input)

    print(summary)
    print(user_data)
