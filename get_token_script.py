import msal
import json
from datetime import datetime

CLIENT_ID = "2bdb3693-4837-4cc6-9f60-ea3858985b16"
TENANT_ID = "e5039572-eed3-431f-92a3-6c3dd04c34fb"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/Files.Read"]

def get_device_code_token():
    app_auth = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    flow = app_auth.initiate_device_flow(scopes=SCOPE)

    if "user_code" not in flow:
        raise RuntimeError("❌ Không thể tạo device flow.")

    print("🔐 Truy cập:", flow["verification_uri"])
    print("🧾 Và nhập mã:", flow["user_code"])

    result = app_auth.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        print("\n✅ Access Token:")
        print(result["access_token"])
        print("\n✅ Refresh Token:")
        print(result["refresh_token"])
        with open(".env", "w") as f:
            f.write(f"ACCESS_TOKEN={result['access_token']}\n")
            f.write(f"REFRESH_TOKEN={result['refresh_token']}\n")
            f.write(f"# Last updated: {datetime.now().isoformat()}\n")
        print("📁 Token đã được lưu vào file .env")
    else:
        print("❌ Không thể lấy token:", result.get("error_description"))

if __name__ == "__main__":
    get_device_code_token()
