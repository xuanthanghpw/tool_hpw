import requests
import time
import json
import random
import sys
from datetime import datetime

API_URL = "https://us-street.api.smarty.com/street-address"
DEFAULT_KEY = "21102174564513388"

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:124.0) Gecko/20100101 Firefox/124.0",
]

ACCEPT_LANGUAGES = [
    "en-US,en;q=0.9",
    "en-US,en;q=0.8,vi;q=0.5",
    "en-GB,en;q=0.9",
    "en-US,en;q=0.9,fr;q=0.6",
    "en-CA,en;q=0.9",
]

HEADERS_FULL_BROWSER = {
    "accept": "*/*",
    "accept-encoding": "gzip, deflate, br, zstd",
    "accept-language": "en-US,en;q=0.9",
    "origin": "https://www.smarty.com",
    "referer": "https://www.smarty.com/",
    "sec-ch-ua": '"Not=A?Brand";v="99", "Opera";v="135", "Chromium";v="151"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "user-agent": USER_AGENTS[0],
}

PROXIES = []
try:
    with open("proxies.txt", "r", encoding="utf-8") as f:
        PROXIES = [line.strip() for line in f if line.strip()]
except FileNotFoundError:
    pass

KEYS = []
try:
    with open("smarty_keys.txt", "r", encoding="utf-8") as f:
        KEYS = [line.strip() for line in f if line.strip()]
except FileNotFoundError:
    pass

if not KEYS:
    KEYS = [DEFAULT_KEY]

COOKIE = ""
try:
    with open("cookie.txt", "r", encoding="utf-8") as f:
        COOKIE = f.read().strip()
except FileNotFoundError:
    pass

class SmartBlockAudit:
    def __init__(self):
        self.report = []

    def log(self, msg):
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
        print(line)
        self.report.append(line)

    def request_with_config(self, proxy=None, key=DEFAULT_KEY, headers=None, cookie=COOKIE, street="1600 Amphitheatre Pkwy, Mountain View, CA 94043"):
        if headers is None:
            headers = HEADERS_FULL_BROWSER.copy()
        if cookie:
            headers["cookie"] = cookie
        params = {
            "key": key,
            "agent": "smarty (website:demo/single-address@latest)",
            "match": "enhanced",
            "candidates": "5",
            "geocode": "true",
            "license": "us-core-cloud",
            "street": street,
        }
        proxies = {"http": proxy, "https": proxy} if proxy else None
        session = requests.Session()
        resp = session.get(API_URL, params=params, headers=headers, proxies=proxies, timeout=20)
        return {
            "status_code": resp.status_code,
            "x_request_id": resp.headers.get("x-request-id", ""),
            "x_license_name": resp.headers.get("x-license-name", ""),
            "body": resp.text[:300],
        }

    def audit_proxy_pool(self):
        self.log("\n=== KIỂM TRA TOÀN BỘ PROXY POOL ===")
        if not PROXIES:
            self.log("KHÔNG CÓ PROXY NÀO. Đây là trạng thái TỐT NHẤT để tránh bị quét.")
            return

        bad_count = 0
        for i, proxy in enumerate(PROXIES):
            self.log(f"Test proxy #{i}: {proxy}")
            try:
                result = self.request_with_config(proxy=proxy)
                if result["status_code"] == 200:
                    self.log(f"  -> OK 200 | license={result['x_license_name']} | req_id={result['x_request_id'][:12]}")
                elif result["status_code"] == 429:
                    self.log("  -> 429 Too Many Requests: PROXY NÀY ĐÃ BỊ SMARTY ĐƯA VÀO DANH SÁCH CHẶN")
                    bad_count += 1
                else:
                    self.log(f"  -> HTTP {result['status_code']}: {result['body'][:100]}")
                    bad_count += 1
            except Exception as e:
                self.log(f"  -> LỖI KẾT NỐI/TIMEOUT: {str(e)[:120]}. PROXY CHẾT HOẶC QUÁ CHẬM.")
                bad_count += 1
            time.sleep(0.5)

        self.log(f"\nTổng kết proxy: {len(PROXIES)} proxy, {bad_count} proxy xấu/bị chặn.")
        if bad_count > 0:
            self.log("KHUYẾN NGHỊ: XÓA proxy xấu khỏi proxies.txt HOẶC bỏ proxy hoàn toàn để chạy bằng IP nhà.")

    def audit_key_list(self):
        self.log("\n=== KIỂM TRA DANH SÁCH API KEY ===")
        if len(KEYS) == 1 and KEYS[0] == DEFAULT_KEY:
            self.log("Chỉ có 1 key public mặc định. Không thể xoay vòng key.")
            self.log("Key public vẫn 200 khi đi IP sạch, nên không cần thay key ngay.")
        else:
            self.log(f"Có {len(KEYS)} key. Mỗi key test 1 lần để xem quota:")
            for i, key in enumerate(KEYS):
                self.log(f"Test key #{i}: ...{key[-4:]}")
                try:
                    result = self.request_with_config(key=key)
                    self.log(f"  -> HTTP {result['status_code']} | {result['body'][:100]}")
                except Exception as e:
                    self.log(f"  -> LỖI: {str(e)[:100]}")
                time.sleep(1)

    def audit_header_integrity(self):
        self.log("\n=== KIỂM TRA HEADER TRÌNH DUYỆT THẬT ===")
        missing = []
        for required in ["accept", "accept-encoding", "accept-language", "origin", "referer", "sec-ch-ua", "sec-fetch-dest", "sec-fetch-mode", "sec-fetch-site", "user-agent"]:
            if required not in HEADERS_FULL_BROWSER:
                missing.append(required)
        if missing:
            self.log(f"THIẾU HEADER QUAN TRỌNG: {', '.join(missing)}")
        else:
            self.log("Đầy đủ header trình duyệt thật.")

        self.log("So sánh header tool cũ vs header thật:")
        old_headers = ["Referer", "Origin", "User-Agent", "Accept-Language"]
        self.log(f"  Header tool cũ chỉ có: {old_headers}")
        self.log("  Thiếu các header bảo mật như sec-ch-ua, sec-fetch-* sẽ khiến request dễ bị gắn cờ bot.")

    def audit_cookie(self):
        self.log("\n=== KIỂM TRA COOKIE ===")
        if not COOKIE:
            self.log("KHÔNG CÓ COOKIE. Vẫn có thể 200 nếu IP sạch và tần suất thấp.")
            self.log("Cookie không bắt buộc, nhưng có cookie trình duyệt thật sẽ giúp giả mạo giống người hơn.")
        else:
            self.log(f"Cookie dài {len(COOKIE)} ký tự. Test request có cookie:")
            try:
                result = self.request_with_config(cookie=COOKIE)
                self.log(f"  -> HTTP {result['status_code']}")
            except Exception as e:
                self.log(f"  -> LỖI: {str(e)[:100]}")

    def audit_request_frequency(self):
        self.log("\n=== KIỂM TRA TẦN SUẤT REQUEST ===")
        self.log("Gửi 6 request liên tiếp không delay, mỗi request 1 lần để tìm ngưỡng bị chặn:")
        block_at = None
        for i in range(6):
            try:
                result = self.request_with_config()
                status = result["status_code"]
                self.log(f"  Request #{i+1}: HTTP {status}")
                if status == 429:
                    block_at = i + 1
                    break
            except Exception as e:
                self.log(f"  Request #{i+1}: LỖI {str(e)[:80]}")
            time.sleep(0.7)
        if block_at:
            self.log(f"Bị chặn sau {block_at} request. Cần tăng delay lên tối thiểu {3 * block_at} giây giữa các request.")
        else:
            self.log("Không bị chặn sau 6 request. Tần suất hiện tại an toàn.")

    def audit_agent_and_license(self):
        self.log("\n=== KIỂM TRA AGENT VÀ LICENSE ===")
        self.log("Trình duyệt thật dùng:")
        self.log("  agent = smarty (website:demo/single-address@latest)")
        self.log("  license = us-core-cloud")
        self.log("Tool cũ dùng:")
        self.log("  agent = smarty (website:demo)")
        self.log("  license = us-rooftop-geocoding-cloud")
        self.log("Sự khác biệt này có thể khiến Smarty phân loại request từ tool khác với web thật.")
        self.log("KHUYẾN NGHỊ: Sửa agent và license trong tool chính cho giống hệt trình duyệt.")

    def run_full_audit(self):
        self.log("BẮT ĐẦU AUDIT TOÀN DIỆN TRƯỚC KHI ĐƯA RA ĐIỂM CẦN CẢI THIỆN")
        self.audit_proxy_pool()
        time.sleep(1)
        self.audit_key_list()
        time.sleep(1)
        self.audit_header_integrity()
        time.sleep(1)
        self.audit_cookie()
        time.sleep(1)
        self.audit_request_frequency()
        time.sleep(1)
        self.audit_agent_and_license()
        self.log("\n=== ĐIỂM CẦN CẢI THIỆN ĐỂ TRÁNH BỊ QUÉT ===")
        self.log("1. BỎ proxy xấu/bị chặn hoặc bỏ proxy hoàn toàn.")
        self.log("2. Dùng agent và license giống hệt trình duyệt thật.")
        self.log("3. Thêm đầy đủ header sec-ch-ua, sec-fetch-* vào mọi request.")
        self.log("4. Dùng cookie trình duyệt thật và không xóa cookie sau mỗi request.")
        self.log("5. Giữ delay 3-5 giây giữa các request, không retry liên tục khi dính 429.")
        self.log("6. Không xoay vòng User-Agent quá nhanh, giữ cố định một User-Agent cho cả phiên.")
        self.log("7. Nếu vẫn bị quét, chuyển sang Selenium/Playwright mở trình duyệt thật.")
        self.log("8. Chạy bằng IP nhà sạch thay vì proxy datacenter bẩn.")

        with open("full_audit_result.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(self.report))
        print("\nĐã lưu full_audit_result.txt")

if __name__ == "__main__":
    audit = SmartBlockAudit()
    audit.run_full_audit()