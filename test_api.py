import requests

res = requests.post("http://127.0.0.1:5000/analyze-user", json={"user_code": "phongnh9"})
print("Status:", res.status_code)
print("Raw text:\n", res.text)
