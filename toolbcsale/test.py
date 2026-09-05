import getpass
import requests

URL = "https://api.hpwcargo.com.vn/api/auth"

# Python tự hỏi token khi chạy
TOKEN = getpass.getpass("Nhập Bearer token: ").strip()

if not TOKEN:
    raise RuntimeError("Bạn chưa nhập token.")

headers = {
    "Accept": "application/json, text/plain, */*",
    "Authorization": f"Bearer {TOKEN}",
    "Origin": "https://admin.hpwcargo.com.vn",
    "Referer": "https://admin.hpwcargo.com.vn/",
    "User-Agent": "Mozilla/5.0",
    "X-Client-Id": "x-provider",
}

print("\nĐang gọi API...")

try:
    response = requests.get(
        URL,
        headers=headers,
        timeout=15
    )

    print("=" * 60)
    print("STATUS :", response.status_code)
    print("ETAG   :", response.headers.get("ETag"))
    print("LIMIT  :", response.headers.get("X-RateLimit-Limit"))
    print("LEFT   :", response.headers.get("X-RateLimit-Remaining"))
    print("=" * 60)

    try:
        print(response.json())
    except ValueError:
        print(response.text)

except requests.RequestException as e:
    print("Request error:", e)



