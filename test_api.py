import requests

res = requests.post("https://mygpt-onedrive-api.onrender.com/analyze-user", json={"user_code": "phongnh9"})
print("Status:", res.status_code)
print("Raw text:\n", res.text)
