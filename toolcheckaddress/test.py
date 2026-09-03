import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import requests
import json
import re
import pandas as pd
import threading
import time
import os
import random
import tempfile
import concurrent.futures
from openpyxl.styles import PatternFill, Font

try:
    from google import genai
    from google.genai import types as genai_types
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

API_URL = "https://us-street.api.smarty.com/street-address"
DEFAULT_SMARTY_KEY = "21102174564513388"

# Pool User-Agent để xoay vòng theo từng request/attempt, tránh bị nhận diện
# là 1 "client" cố định gửi liên tục -> giảm khả năng dính "Too many requests".
USER_AGENT_POOL = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edg/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:124.0) Gecko/20100101 Firefox/124.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
]

ACCEPT_LANGUAGE_POOL = [
    "en-US,en;q=0.9",
    "en-US,en;q=0.8,vi;q=0.5",
    "en-GB,en;q=0.9",
    "en-US,en;q=0.9,fr;q=0.6",
    "en-CA,en;q=0.9",
]

GROQ_MODELS = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant"]
OPENROUTER_MODELS = [
    "meta-llama/llama-3.1-8b-instruct:free",
    "google/gemini-2.0-flash-exp:free",
    "mistralai/mistral-7b-instruct:free",
]
GEMINI_FALLBACK_MODELS = ['gemini-flash-latest', 'gemini-2.5-flash', 'gemini-2.5-flash-lite', 'gemini-2.0-flash']

class SmartyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Smarty API & Google Gemini Auto Check (Xoay vòng Proxy)")
        self.root.geometry("820x780")

        self.excel_path = ""
        self.output_dir = ""
        self.stop_requested = False

        self.session = requests.Session()
        self.gemini_client = None
        self.proxy_list = self._load_proxies()
        self.api_key_list = self._load_api_keys()
        self._identity_counter = 0

        self.setup_gui()

    def _load_proxies(self):
        if os.path.exists("proxies.txt"):
            with open("proxies.txt", "r", encoding="utf-8") as f:
                return [line.strip() for line in f if line.strip()]
        return []

    def _load_api_keys(self):
        """Tuỳ chọn: nếu có file smarty_keys.txt (mỗi dòng 1 key Smarty),
        tool sẽ XOAY VÒNG nhiều key cùng với Proxy/User-Agent để né rate-limit
        triệt để hơn. Nếu không có file, dùng key mặc định."""
        if os.path.exists("smarty_keys.txt"):
            with open("smarty_keys.txt", "r", encoding="utf-8") as f:
                keys = [line.strip() for line in f if line.strip()]
            if keys:
                return keys
        return [DEFAULT_SMARTY_KEY]

    def _next_identity(self):
        """Trả về 1 'danh tính' request mới: Proxy + Smarty Key + User-Agent +
        Accept-Language, xoay vòng độc lập với nhau theo từng lần gọi (không
        chỉ khi bị chặn) để rải request ra nhiều 'vỏ bọc' khác nhau ngay từ đầu,
        giảm khả năng dính 'Too many requests' thay vì chỉ phản ứng sau khi bị chặn."""
        i = self._identity_counter
        self._identity_counter += 1

        proxy_url = self.proxy_list[i % len(self.proxy_list)] if self.proxy_list else None
        api_key = self.api_key_list[i % len(self.api_key_list)]
        user_agent = USER_AGENT_POOL[i % len(USER_AGENT_POOL)]
        accept_language = ACCEPT_LANGUAGE_POOL[i % len(ACCEPT_LANGUAGE_POOL)]

        headers = {
            "Referer": "https://www.smarty.com/",
            "Origin": "https://www.smarty.com",
            "User-Agent": user_agent,
            "Accept-Language": accept_language,
        }
        return {
            "proxy_url": proxy_url,
            "api_key": api_key,
            "headers": headers,
            "index": i,
        }

    def _warmup_proxies(self):
        """'Đánh thức' từng proxy TRƯỚC khi bắt đầu xử lý dòng thật, thay vì để dòng
        đầu tiên lãnh đủ hàng loạt lỗi timeout. Nhiều proxy (đặc biệt proxy free/dùng
        chung) có kết nối 'nguội': lần chạm đầu tiên tới 1 proxy mới phải bắt tay
        TCP/TLS lại từ đầu (và nếu đó là proxy xoay IP/dạng gateway thì backend có
        thể cần thời gian để cấp phiên mới) nên hay timeout, nhưng các lần sau đó
        thường ổn định. Việc này chạy song song, timeout ngắn, và CHỈ in 1 dòng
        tổng kết cuối cùng - không in từng lỗi riêng lẻ để tránh gây cảm giác
        "app bị lỗi liên tục" ngay khi vừa bắt đầu."""
        if not self.proxy_list:
            return

        self.log(f"Đang khởi động {len(self.proxy_list)} proxy trước khi xử lý (có thể mất vài giây)...")

        def _check(proxy_url):
            try:
                requests.get(
                    "https://www.smarty.com/",
                    proxies={"http": proxy_url, "https": proxy_url},
                    timeout=8,
                )
                return True
            except requests.exceptions.RequestException:
                return False

        results = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(10, len(self.proxy_list))) as executor:
            futures = [executor.submit(_check, p) for p in self.proxy_list]
            for fut in concurrent.futures.as_completed(futures, timeout=20):
                try:
                    results.append(fut.result())
                except Exception:
                    results.append(False)

        ok_count = sum(1 for r in results if r)
        self.log(
            f"Đã khởi động xong proxy: {ok_count}/{len(self.proxy_list)} phản hồi tốt "
            f"(số còn lại vẫn sẽ được thử lại tự động trong lúc chạy nếu cần)."
        )

    def setup_gui(self):
        frame_top = tk.Frame(self.root)
        frame_top.pack(pady=(15, 5), padx=10, fill="x")

        self.btn_select_excel = tk.Button(frame_top, text="1. Chọn file Excel gốc", command=self.select_excel, width=20, bg="#28a745", fg="white", font=("Arial", 10, "bold"))
        self.btn_select_excel.pack(side="left", padx=(0, 10))

        self.lbl_excel_path = tk.Label(frame_top, text="Chưa chọn file...", fg="gray")
        self.lbl_excel_path.pack(side="left")

        frame_delay = tk.Frame(self.root)
        frame_delay.pack(pady=5, padx=10, fill="x")

        lbl_delay = tk.Label(frame_delay, text="Độ trễ API Smarty (giây):", font=("Arial", 10, "bold"))
        lbl_delay.pack(side="left")

        self.delay_var = tk.StringVar(value="2.0")
        self.entry_delay = tk.Entry(frame_delay, textvariable=self.delay_var, width=8, font=("Consolas", 11), justify="center")
        self.entry_delay.pack(side="left", padx=10)

        frame_match = tk.LabelFrame(self.root, text=" Chế độ khớp địa chỉ Smarty (match) ", font=("Arial", 10, "bold"), fg="#0078D7")
        frame_match.pack(pady=10, padx=10, fill="x")

        # Mặc định "strict" - đúng như mặc định của Smarty khi nhập tay trên web, chỉ trả
        # kết quả dựa hoàn toàn trên dữ liệu gốc, không tự suy đoán/vá thêm bớt.
        self.match_mode_var = tk.StringVar(value="strict")

        self.rb_match_strict = tk.Radiobutton(
            frame_match, variable=self.match_mode_var, value="strict",
            text="Strict (Khuyến khích - mặc định): chỉ khớp dựa hoàn toàn trên dữ liệu gốc, giống hệt khi nhập tay vào Smarty, không tự vá/suy đoán",
            font=("Arial", 9, "bold"), justify="left", fg="#1da462"
        )
        self.rb_match_strict.pack(anchor="w", padx=10, pady=(5, 0))

        self.rb_match_enhanced = tk.Radiobutton(
            frame_match, variable=self.match_mode_var, value="enhanced",
            text="Enhanced: Smarty tự khớp mạnh tay hơn, có thể tự suy đoán/sửa dữ liệu mơ hồ để cố trả ra kết quả",
            font=("Arial", 9), justify="left"
        )
        self.rb_match_enhanced.pack(anchor="w", padx=10, pady=(0, 2))

        self.lbl_match_note = tk.Label(
            frame_match,
            text="* Khuyến khích dùng Strict để kết quả trung thực với dữ liệu gốc trong Excel. Chỉ chọn Enhanced nếu bạn "
                 "hiểu rõ và chấp nhận việc Smarty có thể tự suy đoán/sửa đổi địa chỉ để cố khớp.",
            fg="gray", font=("Arial", 8), justify="left", wraplength=760
        )
        self.lbl_match_note.pack(anchor="w", padx=10, pady=(0, 5))

        frame_ai = tk.LabelFrame(self.root, text=" Kiểm tra kết quả đáng ngờ bằng AI (Miễn phí) ", font=("Arial", 10, "bold"), fg="#1da462")
        frame_ai.pack(pady=10, padx=10, fill="x")

        self.use_ai_var = tk.BooleanVar(value=False)
        self.chk_use_ai = tk.Checkbutton(frame_ai, text="Bật tự động kiểm tra trạng thái bằng AI", variable=self.use_ai_var, font=("Arial", 10), command=self.toggle_ai_input)
        self.chk_use_ai.pack(anchor="w", padx=10, pady=5)

        frame_mode = tk.Frame(frame_ai)
        frame_mode.pack(anchor="w", padx=20, pady=(0, 5), fill="x")

        self.ai_mode_var = tk.StringVar(value="auto")
        self.rb_mode_auto = tk.Radiobutton(
            frame_mode, variable=self.ai_mode_var, value="auto",
            text="Tự động: hệ thống tự dò & gọi các AI miễn phí (Gemini / Groq / OpenRouter)",
            font=("Arial", 9), justify="left", command=self.toggle_ai_input
        )
        self.rb_mode_auto.pack(anchor="w")

        self.rb_mode_manual = tk.Radiobutton(
            frame_mode, variable=self.ai_mode_var, value="manual",
            text="Thủ công: Tool tạo sẵn Prompt để bạn dán vào AI ngoài (ChatGPT/Gemini web...), rồi dán JSON kết quả về",
            font=("Arial", 9), justify="left", command=self.toggle_ai_input
        )
        self.rb_mode_manual.pack(anchor="w")

        frame_ai_keys = tk.Frame(frame_ai)
        frame_ai_keys.pack(fill="x", padx=20, pady=5)

        lbl_gemini = tk.Label(frame_ai_keys, text="Gemini API Key:", width=18, anchor="w")
        lbl_gemini.grid(row=0, column=0, sticky="w", pady=2)
        self.entry_ai_key = tk.Entry(frame_ai_keys, width=45, font=("Consolas", 10), show="*")
        self.entry_ai_key.grid(row=0, column=1, padx=5, pady=2)

        lbl_groq = tk.Label(frame_ai_keys, text="Groq API Key (tuỳ chọn):", width=18, anchor="w")
        lbl_groq.grid(row=1, column=0, sticky="w", pady=2)
        self.entry_groq_key = tk.Entry(frame_ai_keys, width=45, font=("Consolas", 10), show="*")
        self.entry_groq_key.grid(row=1, column=1, padx=5, pady=2)

        lbl_openrouter = tk.Label(frame_ai_keys, text="OpenRouter API Key (tuỳ chọn):", width=18, anchor="w")
        lbl_openrouter.grid(row=2, column=0, sticky="w", pady=2)
        self.entry_openrouter_key = tk.Entry(frame_ai_keys, width=45, font=("Consolas", 10), show="*")
        self.entry_openrouter_key.grid(row=2, column=1, padx=5, pady=2)

        self.lbl_ai_note = tk.Label(
            frame_ai,
            text="* Chế độ Thủ công không cần API Key. Chế độ Tự động cần ít nhất 1 trong 3 API Key trên.",
            fg="gray", font=("Arial", 8)
        )
        self.lbl_ai_note.pack(anchor="w", padx=20, pady=(0, 5))

        self.lbl_analysis_note = tk.Label(
            frame_ai,
            text=("* Dù BẬT hay TẮT AI, tool luôn tự kiểm tra thêm phần \"analysis\" (DPV) của Smarty "
                  "(vacant / CMRA / no_stat / inactive / dpv_match_code). Kết quả cuối cùng có 3 mức: "
                  "TRUE (chắc chắn đúng) / SUSPECT (nghi ngờ - cần xem lại) / FALSE (lỗi). "
                  "AI không thể tự hạ một dòng đã bị đánh dấu nghi ngờ xuống thành TRUE."),
            fg="#1da462", font=("Arial", 8), justify="left", wraplength=760
        )
        self.lbl_analysis_note.pack(anchor="w", padx=20, pady=(0, 5))

        self._set_ai_inputs_state(tk.DISABLED)

        if not HAS_GEMINI:
            lbl_gemini.config(text="Gemini (chưa cài 'google-genai'):", fg="red")

        frame_btns = tk.Frame(self.root)
        frame_btns.pack(pady=10)

        self.btn_start = tk.Button(frame_btns, text="2. Bắt đầu xử lý & Xuất file (TXT + Excel)", command=self.start_processing_thread, width=35, bg="#0078D7", fg="white", font=("Arial", 11, "bold"), state=tk.DISABLED)
        self.btn_start.pack(side="left", padx=5)

        self.btn_stop = tk.Button(frame_btns, text="Dừng lại", command=self.stop_processing, width=10, bg="#dc3545", fg="white", font=("Arial", 11, "bold"), state=tk.DISABLED)
        self.btn_stop.pack(side="left", padx=5)

        lbl_log = tk.Label(self.root, text="Tiến trình xử lý:", font=("Arial", 10, "bold"))
        lbl_log.pack(anchor="w", padx=10)

        self.log_text = scrolledtext.ScrolledText(self.root, font=("Consolas", 10), width=95, height=18, bg="#1E1E1E", fg="#D4D4D4")
        self.log_text.pack(pady=5, padx=10)

    def _set_ai_inputs_state(self, state):
        self.entry_ai_key.config(state=state)
        self.entry_groq_key.config(state=state)
        self.entry_openrouter_key.config(state=state)
        self.rb_mode_auto.config(state=state)
        self.rb_mode_manual.config(state=state)

    def toggle_ai_input(self):
        if self.use_ai_var.get():
            self._set_ai_inputs_state(tk.NORMAL)
        else:
            self._set_ai_inputs_state(tk.DISABLED)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def select_excel(self):
        file_path = filedialog.askopenfilename(title="Chọn file Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.excel_path = file_path
            self.lbl_excel_path.config(text=self.excel_path, fg="black")
            self.btn_start.config(state=tk.NORMAL)
            self.log(f"Đã chọn file: {self.excel_path}")
            if self.proxy_list:
                self.log(f"Đã nạp {len(self.proxy_list)} proxy từ file proxies.txt (xoay vòng).")
            else:
                self.log("Không có proxies.txt -> chỉ xoay vòng User-Agent/Key, khi bị chặn sẽ phải chờ (backoff).")
            if len(self.api_key_list) > 1:
                self.log(f"Đã nạp {len(self.api_key_list)} Smarty Key từ file smarty_keys.txt (xoay vòng).")
            self.log(f"Đang xoay vòng {len(USER_AGENT_POOL)} User-Agent khác nhau cho mỗi request.")

    def stop_processing(self):
        self.stop_requested = True
        self.log("\n[HỆ THỐNG] Đang yêu cầu dừng tiến trình...")
        self.btn_stop.config(state=tk.DISABLED)

    def start_processing_thread(self):
        try:
            delay_val = float(self.delay_var.get())
            if delay_val < 0: raise ValueError
        except ValueError:
            messagebox.showwarning("Cảnh báo", "Độ trễ không hợp lệ!")
            return

        if self.use_ai_var.get() and self.ai_mode_var.get() == "auto":
            keys = [self.entry_ai_key.get().strip(), self.entry_groq_key.get().strip(), self.entry_openrouter_key.get().strip()]
            if not any(keys):
                messagebox.showwarning("Cảnh báo", "Chế độ Tự động cần ít nhất 1 API Key!\nHoặc chuyển sang chế độ Thủ công.")
                return

        save_path = filedialog.asksaveasfilename(
            title="Lưu kết quả",
            defaultextension="",
            initialfile="KetQua_Smarty",
        )
        if not save_path:
            return

        base_path = save_path.replace('.txt', '').replace('.xlsx', '')
        self.txt_output_path = base_path + ".txt"
        self.excel_output_path = base_path + ".xlsx"

        self.stop_requested = False
        self.btn_start.config(state=tk.DISABLED, text="Đang xử lý...")
        self.btn_select_excel.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.entry_delay.config(state=tk.DISABLED)
        self.rb_match_strict.config(state=tk.DISABLED)
        self.rb_match_enhanced.config(state=tk.DISABLED)
        self.chk_use_ai.config(state=tk.DISABLED)
        self._set_ai_inputs_state(tk.DISABLED)

        threading.Thread(target=self.process_data, daemon=True).start()

    def process_data(self):
        collected_data = []

        try:
            delay_seconds = float(self.delay_var.get())
            self.log("Đang đọc dữ liệu từ file Excel...")
            df = pd.read_excel(self.excel_path)

            address_cols = ['Shipping Address1', 'Shipping Address2', 'Shipping City', 'Shipping State', 'Shipping PostalCode', 'Shipping Country']
            ref_col = 'Shipment Reference Id'
            required_cols = address_cols + [ref_col]

            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                messagebox.showerror("Lỗi dữ liệu", f"File thiếu các cột:\n{', '.join(missing_cols)}")
                self.reset_ui()
                return

            total_rows = len(df)
            match_mode = self.match_mode_var.get()
            self.log(f"Chế độ khớp Smarty: {match_mode.upper()}" + (" (khuyến khích)" if match_mode == "strict" else ""))

            self._warmup_proxies()

            self.log(f"Bắt đầu giai đoạn 1: Gọi API Smarty ({total_rows} dòng)...")

            f_handle = self._open_output_txt_with_retry(self.txt_output_path)
            self.txt_output_path = f_handle.name
            with f_handle as f_out:
                for index, row in df.iterrows():
                    if self.stop_requested:
                        self.log("\n[HỆ THỐNG] Tiến trình dừng bởi người dùng.")
                        break

                    ref_id = str(row[ref_col]).strip() if pd.notna(row[ref_col]) else "NO_REF"
                    # QUAN TRỌNG: KHÔNG được biến đổi/giải mã dữ liệu đầu vào từ Excel dưới bất kỳ
                    # hình thức nào (kể cả HTML entity như "&#39;", "&amp;"...). Dữ liệu gốc phải
                    # được giữ nguyên 100% và gửi y hệt như vậy cho API, để đảm bảo tính trung thực
                    # và nhất quán với đúng những gì có trong file Excel gốc của người dùng.
                    parts = [str(row[col]).strip() for col in address_cols if pd.notna(row[col]) and str(row[col]).strip() != ""]
                    input_string = " ".join(parts)

                    if not input_string:
                        continue

                    outcome = self.call_smarty_api_with_retry(input_string, 10, index+1)
                    result_string = outcome.get("result", "")
                    analysis = outcome.get("analysis", {}) or {}

                    dpv_match_code = analysis.get("dpv_match_code", "")
                    dpv_vacant = analysis.get("dpv_vacant", "")
                    dpv_cmra = analysis.get("dpv_cmra", "")
                    dpv_no_stat = analysis.get("dpv_no_stat", "")
                    active = analysis.get("active", "")
                    # Trước đây 2 trường này bị GỘP LẪN vào 1 biến "footnotes" duy nhất (và không
                    # hề được xuất ra Excel) -> không ai kiểm chứng được vì sao 1 dòng bị/không bị
                    # gắn cờ. Giờ tách riêng, xuất cả 2 raw fields ra Excel để minh bạch/debug.
                    dpv_footnotes_raw = analysis.get("dpv_footnotes", "")
                    footnotes_raw = analysis.get("footnotes", "")

                    analysis_reasons = self._compute_analysis_flags(analysis)
                    # Kiểm tra ĐỘC LẬP trên chính chuỗi input gốc (chưa qua Smarty). Cộng dồn với
                    # analysis_reasons: dù Smarty (đặc biệt match=enhanced) có tự "vá" và trả về
                    # dpv_match_code=Y cho input bẩn, dòng này vẫn phải bị đưa vào diện nghi ngờ vì
                    # bản thân dữ liệu nguồn có vấn đề (vd còn dính "&#39;" - mã HTML entity).
                    input_anomaly_reasons = self._detect_input_anomalies(input_string)
                    all_reasons = analysis_reasons + input_anomaly_reasons

                    analysis_summary = "; ".join(analysis_reasons) if analysis_reasons else (
                        "Không có cảnh báo" if analysis else "Không có dữ liệu analysis"
                    )
                    if input_anomaly_reasons:
                        analysis_summary += " | " + "; ".join(input_anomaly_reasons)

                    f_out.write(f"({input_string} : {result_string} | Analysis: {analysis_summary} [{ref_id}])\n")
                    f_out.flush()

                    item = {
                        "Reference Id": ref_id,
                        "Chuỗi đầu vào": input_string,
                        "Chuỗi đầu ra": result_string,
                        "DPV Match Code": dpv_match_code,
                        "DPV Vacant": dpv_vacant,
                        "DPV CMRA": dpv_cmra,
                        "DPV No Stat": dpv_no_stat,
                        "Active": active,
                        "DPV Footnotes": dpv_footnotes_raw,
                        "Footnotes": footnotes_raw,
                        # Tự động đánh dấu "suspect" ngay từ bước gọi API (kể cả khi TẮT AI),
                        # dựa trên dữ liệu "analysis" của Smarty VÀ chất lượng của chính input gốc,
                        # không chỉ dựa vào so khớp chuỗi hay việc Smarty có match được hay không.
                        "Trạng thái xử lý": "suspect" if all_reasons else "",
                        "Lý do đáng ngờ": ("Cảnh báo: " + "; ".join(all_reasons)) if all_reasons else "",
                    }
                    item["_analysis_reasons"] = all_reasons
                    collected_data.append(item)

                    if outcome.get("success"):
                        self.log(f"Smarty API - Dòng {index + 1}/{total_rows} -> OK")
                    else:
                        self.log(f"Smarty API - Dòng {index + 1}/{total_rows} -> LỖI: {result_string}")

                    # Chỉ delay nếu không dùng proxy hoặc đang chạy luồng bình thường
                    if delay_seconds > 0:
                        time.sleep(delay_seconds)

                if not self.stop_requested:
                    gpt_prompt = """
=======================================================================================================================
[SYSTEM PROMPT - DÀNH CHO AI/CHATGPT/GEMINI]
Nhiệm vụ của bạn là đóng vai trò chuyên gia kiểm duyệt địa chỉ quốc tế khắt khe. Dựa vào danh sách đối chiếu bên trên (định dạng: (chuỗi gốc : chuỗi chuẩn hóa từ API | Analysis: <cảnh báo DPV nếu có> [Mã đơn hàng])), hãy thực hiện:

1. QUÉT TOÀN BỘ danh sách và LỌC RA những đơn hàng CÓ DẤU HIỆU ĐÁNG NGỜ hoặc LỖI.
2. Tiêu chí bắt LỖI (out không hợp lệ):
   - Trả về "Không tìm thấy kết quả từ Server", "Lỗi HTTP...", "Lỗi API...".
   - Trả về "Không tìm thấy delivery_line_1 hoặc last_line".
   - Chuỗi trả về bị thiếu quá nhiều thành phần trọng yếu so với chuỗi gốc (mất hẳn số nhà, tên đường khác hoàn toàn, sai khác mã ZIP nghiêm trọng).
3. Tiêu chí bắt NGHI NGỜ (out khớp chuỗi nhưng vẫn cần xem lại) — QUAN TRỌNG, đừng chỉ so khớp chuỗi:
   - Phần "Analysis" đi kèm có cảnh báo, ví dụ: địa chỉ đang bị BỎ TRỐNG (vacant), là hộp thư CMRA, không có dữ liệu xác thực (no_stat), địa chỉ KHÔNG hoạt động (active=N), hoặc mã khớp DPV chưa xác nhận đầy đủ.
   - Những trường hợp này KHÔNG được coi là "hợp lệ" chỉ vì chuỗi địa chỉ khớp — phải liệt kê vào diện đáng ngờ.
4. ĐẦU RA BẮT BUỘC TRÌNH BÀY DƯỚI DẠNG BẢNG (Markdown Table) với các cột sau:
   | Mã đơn hàng (ID) | Chuỗi đầu vào (Gốc) | Chuỗi đầu ra (API trả về) | Loại (Lỗi/Nghi ngờ) | Lý do chi tiết |

TUYỆT ĐỐI KHÔNG đưa các đơn hàng hoàn toàn hợp lệ (không lỗi, không cảnh báo Analysis) vào bảng này. Nếu tất cả đều hợp lệ, hãy trả lời ngắn gọn: "Tất cả địa chỉ đều hợp lệ."
=======================================================================================================================
"""
                    f_out.write(gpt_prompt)
                    f_out.flush()

            if self.use_ai_var.get() and not self.stop_requested and collected_data:
                self.log("\n===========================================")
                self.log("Bắt đầu giai đoạn 2: Kiểm tra kết quả đáng ngờ bằng AI...")

                mode = self.ai_mode_var.get()
                if mode == "manual":
                    self.log("Chế độ: THỦ CÔNG (bạn tự dán Prompt vào AI ngoài).")
                    self.run_manual_flow(collected_data)
                else:
                    self.log("Chế độ: TỰ ĐỘNG (hệ thống tự dò & gọi AI miễn phí).")
                    self.run_auto_flow(collected_data)

            elif not self.use_ai_var.get():
                self.log("\n[INFO] AI tự động đang TẮT. File TXT đã có Prompt để bạn tự check thủ công.")

            if collected_data:
                self.log("\nBắt đầu tạo file Excel...")
                self.excel_output_path = self.export_excel(collected_data, self.excel_output_path)
                self.log(f"-> Đã lưu Excel: {self.excel_output_path}")

            if not self.stop_requested:
                self.log(f"\n[THÀNH CÔNG] Toàn bộ tiến trình hoàn tất!")
                messagebox.showinfo("Hoàn tất", f"Đã xuất thành công:\n1. {self.txt_output_path}\n2. {self.excel_output_path}")

        except Exception as e:
            self.log(f"\n[LỖI] Đã xảy ra sự cố: {e}")
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra:\n{e}")
        finally:
            self.reset_ui()

    def run_auto_flow(self, collected_data):
        BATCH_SIZE = 20
        n = len(collected_data)
        self.log(f"Tổng {n} dòng -> chia thành các batch {BATCH_SIZE} dòng/lượt gọi AI.")

        idx = 0
        while idx < n:
            if self.stop_requested:
                return
            batch = collected_data[idx: idx + BATCH_SIZE]
            batch_no = idx // BATCH_SIZE + 1
            total_batches = (n + BATCH_SIZE - 1) // BATCH_SIZE
            self.log(f"-> Đang thử gọi AI tự động cho batch {batch_no}/{total_batches} ({len(batch)} dòng)...")

            results_map = self.check_with_ai_batch_auto(batch)

            if results_map is None:
                self.log("    [!] Không có nhà cung cấp AI miễn phí nào phản hồi thành công cho batch này.")
                remaining_items = collected_data[idx:]
                want_manual = self.ask_switch_to_manual(
                    f"Hệ thống không gọi được AI tự động.\n"
                    f"Còn {len(remaining_items)} dòng chưa được kiểm tra."
                )
                if want_manual:
                    self.log("-> Người dùng chọn chuyển sang chế độ THỦ CÔNG cho các dòng còn lại.")
                    self.run_manual_flow(remaining_items)
                else:
                    self.log("-> Người dùng chọn bỏ qua kiểm tra AI cho các dòng còn lại.")
                    for item in remaining_items:
                        self._apply_ai_result_with_analysis(
                            item, None,
                            missing_reason="Bỏ qua kiểm tra AI (không gọi được AI tự động)"
                        )
                return

            for item in batch:
                rid = item["Reference Id"]
                res = results_map.get(rid)
                self._apply_ai_result_with_analysis(item, res)

            self.log(f"-> Batch {batch_no}/{total_batches} hoàn tất.")
            idx += BATCH_SIZE
            if idx < n:
                time.sleep(1.0)

    def check_with_ai_batch_auto(self, items):
        prompt = self._build_ai_prompt(items)
        raw = self.call_any_available_ai(prompt)
        if raw is None:
            return None
        return self._parse_ai_json_array(raw)

    def call_any_available_ai(self, prompt):
        gemini_key = self.entry_ai_key.get().strip()
        groq_key = self.entry_groq_key.get().strip()
        openrouter_key = self.entry_openrouter_key.get().strip()

        providers = []
        if gemini_key and HAS_GEMINI:
            providers.append(("Gemini", lambda: self._call_gemini_fast(prompt, gemini_key)))
        if groq_key:
            providers.append(("Groq", lambda: self._call_groq(prompt, groq_key)))
        if openrouter_key:
            providers.append(("OpenRouter", lambda: self._call_openrouter(prompt, openrouter_key)))

        if not providers:
            self.log("    [!] Chưa nhập bất kỳ API Key nào cho chế độ Tự động.")
            return None

        for name, fn in providers:
            if self.stop_requested:
                return None
            self.log(f"    -> Đang thử {name}...")
            try:
                raw = fn()
                if raw and raw.strip():
                    self.log(f"    [OK] {name} phản hồi thành công.")
                    return raw
                self.log(f"    [!] {name} trả về rỗng, thử provider kế tiếp...")
            except Exception as e:
                self.log(f"    [!] {name} lỗi: {str(e)[:120]}")
                continue

        return None

    def _call_gemini_fast(self, prompt, api_key):
        client = genai.Client(api_key=api_key)
        models = self._quick_discover_gemini_models(client) or GEMINI_FALLBACK_MODELS
        last_err = None

        for model_name in models[:2]:
            for attempt in range(2):
                if self.stop_requested:
                    return None
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt,
                        config=genai_types.GenerateContentConfig(
                            temperature=0.0,
                            response_mime_type="application/json"
                        )
                    )
                    return (response.text or "").strip()
                except Exception as e:
                    last_err = e
                    es = str(e).lower()
                    if "404" in es or "not found" in es:
                        break
                    if "429" in es or "503" in es or "overloaded" in es or "unavailable" in es or "quota" in es:
                        if attempt == 0:
                            time.sleep(3)
                            continue
                        break
                    break

        if last_err:
            raise last_err
        raise RuntimeError("Gemini không khả dụng")

    def _quick_discover_gemini_models(self, client):
        try:
            discovered = []
            for m in client.models.list():
                name = getattr(m, "name", "") or ""
                short = name.split("/")[-1] if "/" in name else name
                if not short:
                    continue
                actions = getattr(m, "supported_actions", None) or []
                if actions and "generateContent" not in actions:
                    continue
                low = short.lower()
                if any(x in low for x in ["embedding", "tts", "image", "video", "vision", "aqa", "gemma"]):
                    continue
                if "flash" not in low and "pro" not in low:
                    continue
                discovered.append(short)

            def priority(name):
                n = name.lower()
                if "flash-lite" in n:
                    return 1
                if n.endswith("-latest") and "flash" in n:
                    return 0
                if "flash" in n:
                    return 2
                if "pro" in n:
                    return 3
                return 4

            discovered = sorted(set(discovered), key=priority)
            return discovered
        except Exception:
            return []

    def _call_groq(self, prompt, api_key):
        url = "https://api.groq.com/openai/v1/chat/completions"
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        last_err = None
        for model_name in GROQ_MODELS[:2]:
            body = {
                "model": model_name,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0
            }
            try:
                r = self.session.post(url, headers=headers, json=body, timeout=20)
                if r.status_code == 200:
                    data = r.json()
                    return data["choices"][0]["message"]["content"]
                elif r.status_code == 404:
                    last_err = RuntimeError(f"Model '{model_name}' không tồn tại trên Groq")
                    continue
                elif r.status_code == 429:
                    last_err = RuntimeError("Groq quá tải/hết quota (429)")
                    time.sleep(2)
                    continue
                else:
                    last_err = RuntimeError(f"Groq lỗi HTTP {r.status_code}: {r.text[:150]}")
                    continue
            except requests.exceptions.RequestException as e:
                last_err = e
                continue
        raise last_err or RuntimeError("Groq thất bại")

    def _call_openrouter(self, prompt, api_key):
        url = "https://openrouter.ai/api/v1/chat/completions"
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        last_err = None
        for model_name in OPENROUTER_MODELS:
            body = {
                "model": model_name,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0
            }
            try:
                r = self.session.post(url, headers=headers, json=body, timeout=20)
                if r.status_code == 200:
                    data = r.json()
                    return data["choices"][0]["message"]["content"]
                elif r.status_code == 404:
                    last_err = RuntimeError(f"Model '{model_name}' không tồn tại trên OpenRouter")
                    continue
                elif r.status_code == 429:
                    last_err = RuntimeError("OpenRouter quá tải/hết quota (429)")
                    time.sleep(2)
                    continue
                else:
                    last_err = RuntimeError(f"OpenRouter lỗi HTTP {r.status_code}: {r.text[:150]}")
                    continue
            except requests.exceptions.RequestException as e:
                last_err = e
                continue
        raise last_err or RuntimeError("OpenRouter thất bại")

    def run_manual_flow(self, items):
        if not items:
            return
        self.log(f"[THỦ CÔNG] Đang tạo Prompt cho {len(items)} dòng để bạn sao chép...")
        prompt = self._build_ai_prompt(items)

        json_text = self.get_manual_json_from_user(prompt, len(items))

        if not json_text:
            self.log("    [!] Bạn đã bỏ qua bước nhập JSON thủ công -> giữ nguyên cờ Analysis (nếu có), còn lại 'unknown'.")
            for item in items:
                self._apply_ai_result_with_analysis(
                    item, None,
                    missing_reason="Bỏ qua kiểm tra AI thủ công"
                )
            return

        results_map = self._parse_ai_json_array(json_text)
        if not results_map:
            self.log("    [!] Không đọc được JSON hợp lệ từ nội dung bạn dán vào -> các dòng này sẽ ở trạng thái 'unknown'.")

        for item in items:
            rid = item["Reference Id"]
            res = results_map.get(rid)
            if res is None:
                self._apply_ai_result_with_analysis(item, None, missing_reason="Không thấy id này trong JSON đã nhập")
            else:
                self._apply_ai_result_with_analysis(item, res)

        self.log(f"    -> Đã áp dụng kết quả thủ công cho {len(items)} dòng.")

    def _apply_ai_result_with_analysis(self, item, res, missing_reason="Không nhận được kết quả từ AI cho dòng này"):
        """Gộp kết quả AI (true/suspect/false) với cờ nghi ngờ tính sẵn từ 'analysis'
        của Smarty (dpv_vacant, dpv_cmra, dpv_no_stat, active, dpv_match_code...).
        Nguyên tắc: nếu analysis đã cảnh báo, AI KHÔNG được phép hạ xuống 'true'."""
        analysis_reasons = item.get("_analysis_reasons", [])
        analysis_text = "; ".join(analysis_reasons)

        if res is None:
            if analysis_reasons:
                # Đã có cờ nghi ngờ từ analysis từ trước -> giữ nguyên "suspect", chỉ bổ sung ghi chú
                base_reason = item.get("Lý do đáng ngờ", "") or f"Cảnh báo Analysis: {analysis_text}"
                item["Trạng thái xử lý"] = "suspect"
                item["Lý do đáng ngờ"] = f"{base_reason} | {missing_reason}"
            else:
                item["Trạng thái xử lý"] = "unknown"
                item["Lý do đáng ngờ"] = missing_reason
            return

        ai_status = res.get("status", "unknown")
        ai_reason = (res.get("reason", "") or "").strip()

        if not analysis_reasons:
            item["Trạng thái xử lý"] = ai_status
            item["Lý do đáng ngờ"] = ai_reason
            return

        # Có cảnh báo từ analysis: AI nói "false" -> vẫn giữ "false" (nặng hơn).
        # AI nói "true" hoặc "suspect" -> luôn là "suspect" tối thiểu, không được là "true".
        if ai_status == "false":
            final_status = "false"
        else:
            final_status = "suspect"

        reason_parts = []
        if ai_reason:
            reason_parts.append(f"AI: {ai_reason}")
        reason_parts.append(f"Cảnh báo Analysis: {analysis_text}")
        item["Trạng thái xử lý"] = final_status
        item["Lý do đáng ngờ"] = " | ".join(reason_parts)

    def _ask_retry_cancel(self, title, message):
        """Hỏi người dùng 'Thử lại' / 'Huỷ' theo kiểu thread-safe (an toàn khi gọi từ luồng xử lý
        nền process_data, vì messagebox phải được tạo trên luồng GUI chính)."""
        result_holder = {"value": False}
        event = threading.Event()

        def _build():
            try:
                result_holder["value"] = messagebox.askretrycancel(title, message)
            finally:
                event.set()

        self.root.after(0, _build)
        event.wait()
        return result_holder["value"]

    def _alt_path_with_timestamp(self, path):
        """Sinh ra 1 đường dẫn khác (thêm hậu tố thời gian) để lưu dự phòng khi không thể ghi
        đè vào file gốc, tránh làm mất toàn bộ dữ liệu đã xử lý."""
        base, ext = os.path.splitext(path)
        return f"{base}_{time.strftime('%Y%m%d_%H%M%S')}{ext}"

    def _open_output_txt_with_retry(self, path):
        """Mở file TXT kết quả ở chế độ ghi (tạo mới/ghi đè). Nếu file đang bị một chương trình
        khác (Excel, Notepad, v.v.) mở và khoá nên không ghi được, hỏi người dùng:
        - 'Retry': thử mở lại (sau khi họ đã đóng file đó).
        - 'Cancel': tự động đổi sang tên file khác (kèm timestamp) để không mất dữ liệu."""
        current = path
        while True:
            try:
                return open(current, 'w', encoding='utf-8')
            except PermissionError:
                self.log(f"  [!] Không thể ghi file TXT '{os.path.basename(current)}' - có thể đang mở ở chương trình khác.")
                retry = self._ask_retry_cancel(
                    "Không thể ghi file TXT",
                    f"File:\n{current}\n\ncó thể đang được MỞ trong một chương trình khác (Excel, Notepad, v.v.) "
                    f"nên không thể ghi vào.\n\nBấm 'Retry' SAU KHI bạn đã đóng file đó lại.\n"
                    f"Bấm 'Cancel' để tự động lưu sang một tên file khác (không mất dữ liệu)."
                )
                if retry:
                    continue
                current = self._alt_path_with_timestamp(path)
                self.log(f"  [i] Sẽ lưu file TXT sang tên khác: {current}")

    def _save_workbook_with_retry(self, wb, desired_path):
        """Lưu workbook openpyxl vào desired_path. Đây chính là chỗ hay gặp lỗi 'PermissionError'
        khi file Excel kết quả đang được người dùng MỞ SẴN (trùng tên với file họ định lưu).
        Thay vì crash toàn bộ tiến trình và bắt người dùng tự đóng file rồi chạy lại từ đầu, ở đây
        sẽ hỏi:
        - 'Retry': thử lưu lại (sau khi người dùng đã đóng file Excel đang mở đó).
        - 'Cancel': tự động lưu kết quả sang một tên file khác (kèm timestamp), không mất dữ liệu
          đã xử lý (đặc biệt quan trọng vì bước này diễn ra SAU KHI đã gọi API + AI xong)."""
        path = desired_path
        while True:
            try:
                wb.save(path)
                return path
            except PermissionError:
                self.log(f"  [!] Không thể lưu Excel '{os.path.basename(path)}' - file đang được MỞ ở chương trình khác.")
                retry = self._ask_retry_cancel(
                    "Không thể lưu file Excel",
                    f"File:\n{path}\n\ncó thể đang được MỞ trong Excel (hoặc chương trình khác) nên không thể "
                    f"ghi đè lên được.\n\nHãy ĐÓNG file đó lại, sau đó bấm 'Retry' để lưu lại.\n"
                    f"Bấm 'Cancel' để tự động lưu kết quả sang một tên file khác (không mất dữ liệu đã xử lý)."
                )
                if retry:
                    continue
                path = self._alt_path_with_timestamp(desired_path)
                self.log(f"  [i] Sẽ lưu file Excel sang tên khác: {path}")
            except OSError as e:
                # Các lỗi OS khác (đĩa đầy, đường dẫn không hợp lệ...) cũng không nên làm mất
                # toàn bộ dữ liệu đã xử lý -> vẫn cho người dùng cơ hội thử lại / đổi tên.
                self.log(f"  [!] Lỗi khi lưu file Excel: {e}")
                retry = self._ask_retry_cancel(
                    "Không thể lưu file Excel",
                    f"Đã xảy ra lỗi khi lưu file:\n{path}\n\nLỗi: {e}\n\n"
                    f"Bấm 'Retry' để thử lưu lại.\nBấm 'Cancel' để tự động lưu sang một tên file khác."
                )
                if retry:
                    continue
                path = self._alt_path_with_timestamp(desired_path)
                self.log(f"  [i] Sẽ lưu file Excel sang tên khác: {path}")

    def ask_switch_to_manual(self, reason_text):
        result_holder = {"value": False}
        event = threading.Event()

        def _build():
            try:
                ans = messagebox.askyesno(
                    "Không gọi được AI tự động",
                    f"{reason_text}\n\nBạn có muốn chuyển sang chế độ THỦ CÔNG "
                    f"(tự dán Prompt vào AI ngoài, rồi dán JSON kết quả về) không?"
                )
                result_holder["value"] = bool(ans)
            finally:
                event.set()

        self.root.after(0, _build)
        event.wait()
        return result_holder["value"]

    def get_manual_json_from_user(self, prompt_text, item_count):
        result_holder = {"value": None}
        event = threading.Event()

        def _build():
            dlg = tk.Toplevel(self.root)
            dlg.title("Chế độ thủ công - Dán Prompt vào AI ngoài")
            dlg.geometry("720x650")
            dlg.grab_set()

            tk.Label(
                dlg,
                text=(f"Có {item_count} dòng cần kiểm tra.\n"
                      "1) Bấm 'Sao chép Prompt' rồi dán vào 1 AI bất kỳ (ChatGPT, Gemini web, Claude...).\n"
                      "2) Sao chép TOÀN BỘ phần JSON mà AI trả về, dán vào ô phía dưới.\n"
                      "3) Bấm 'Xác nhận' để tool tự tạo Excel."),
                justify="left", anchor="w"
            ).pack(anchor="w", padx=10, pady=(10, 5))

            tk.Label(dlg, text="Prompt để dán vào AI ngoài:", font=("Arial", 9, "bold")).pack(anchor="w", padx=10)
            txt_prompt = scrolledtext.ScrolledText(dlg, height=14, wrap="word", font=("Consolas", 9))
            txt_prompt.insert("1.0", prompt_text)
            txt_prompt.pack(fill="both", expand=True, padx=10, pady=(0, 5))

            def copy_prompt():
                dlg.clipboard_clear()
                dlg.clipboard_append(prompt_text)
                lbl_copied.config(text="Đã sao chép vào Clipboard!", fg="#28a745")

            frame_copy = tk.Frame(dlg)
            frame_copy.pack(pady=2)
            tk.Button(frame_copy, text="Sao chép Prompt", command=copy_prompt, bg="#0078D7", fg="white", width=20).pack(side="left", padx=5)
            lbl_copied = tk.Label(frame_copy, text="", font=("Arial", 9))
            lbl_copied.pack(side="left", padx=5)

            tk.Label(dlg, text="Dán JSON kết quả từ AI vào đây:", font=("Arial", 9, "bold")).pack(anchor="w", padx=10, pady=(10, 0))
            txt_json = scrolledtext.ScrolledText(dlg, height=10, wrap="word", font=("Consolas", 9))
            txt_json.pack(fill="both", expand=True, padx=10, pady=(0, 5))

            def on_confirm():
                content = txt_json.get("1.0", "end").strip()
                result_holder["value"] = content if content else None
                dlg.grab_release()
                dlg.destroy()
                event.set()

            def on_skip():
                result_holder["value"] = None
                dlg.grab_release()
                dlg.destroy()
                event.set()

            frame_btn = tk.Frame(dlg)
            frame_btn.pack(pady=10)
            tk.Button(frame_btn, text="Xác nhận", command=on_confirm, bg="#28a745", fg="white", width=14, font=("Arial", 10, "bold")).pack(side="left", padx=5)
            tk.Button(frame_btn, text="Bỏ qua kiểm tra AI cho các dòng này", command=on_skip, bg="#dc3545", fg="white", width=32).pack(side="left", padx=5)

            dlg.protocol("WM_DELETE_WINDOW", on_skip)

        self.root.after(0, _build)
        event.wait()
        return result_holder["value"]

    def _build_ai_prompt(self, items):
        payload = [
            {
                "id": it["Reference Id"],
                "in": it["Chuỗi đầu vào"],
                "out": it["Chuỗi đầu ra"],
                "analysis": {
                    "dpv_match_code": it.get("DPV Match Code", ""),
                    "dpv_vacant": it.get("DPV Vacant", ""),
                    "dpv_cmra": it.get("DPV CMRA", ""),
                    "dpv_no_stat": it.get("DPV No Stat", ""),
                    "active": it.get("Active", ""),
                    "dpv_footnotes": it.get("DPV Footnotes", ""),
                    "footnotes": it.get("Footnotes", ""),
                },
            }
            for it in items
        ]

        prompt = f"""Bạn là chuyên gia kiểm duyệt địa chỉ vận chuyển quốc tế, làm việc rất khắt khe và chính xác.
Dưới đây là một danh sách JSON gồm nhiều đơn hàng. Mỗi phần tử có:
- "id": mã đơn hàng
- "in": chuỗi địa chỉ gốc
- "out": chuỗi địa chỉ đã chuẩn hóa từ API Smarty
- "analysis": dữ liệu xác thực DPV (Delivery Point Validation) của USPS đi kèm, gồm:
   + "dpv_match_code": "Y" = khớp hoàn toàn, đã xác nhận; "D"/"S"/"N" hoặc rỗng = CHƯA xác nhận đầy đủ.
   + "dpv_vacant": "Y" = USPS liệt kê địa chỉ này ĐANG BỊ BỎ TRỐNG (vacant) — rất đáng ngờ dù chuỗi khớp.
   + "dpv_cmra": "Y" = đây là hộp thư nhận hộ (CMRA), không phải địa chỉ nhà/công ty thực tế.
   + "dpv_no_stat": "Y" = USPS không có đủ dữ liệu để xác thực địa chỉ này.
   + "active": "N" = địa chỉ hiện KHÔNG còn hoạt động.

QUAN TRỌNG: Việc "out" khớp với "in" về mặt CHUỖI KÝ TỰ là CHƯA ĐỦ để kết luận địa chỉ chắc chắn đúng. Bạn PHẢI kết hợp cả phần "analysis" để đánh giá — đây chính là mục đích của việc kiểm tra này.

Với MỖI phần tử, hãy gán MỘT trong 3 trạng thái:
- status = "true": CHỈ khi "out" khớp tốt với "in" (đủ số nhà, tên đường, thành phố, mã ZIP) VÀ "analysis" không có cảnh báo nào (dpv_match_code="Y", dpv_vacant khác "Y", dpv_cmra khác "Y", dpv_no_stat khác "Y", active khác "N").
- status = "suspect": khi "out" khớp với "in" về chuỗi NHƯNG "analysis" có ít nhất một cảnh báo (dpv_vacant="Y", dpv_cmra="Y", dpv_no_stat="Y", active="N", hoặc dpv_match_code khác "Y"). Đây là các địa chỉ ĐƯỢC TÌM THẤY nhưng cần con người xem lại trước khi giao hàng — KHÔNG được gán "true" cho các trường hợp này.
- status = "false": khi "out" chứa các cụm lỗi như "Lỗi", "Không tìm thấy kết quả", "Không tìm thấy delivery_line_1", hoặc "out" thiếu/sai lệch nghiêm trọng so với "in" (mất số nhà, tên đường khác hẳn, sai ZIP nghiêm trọng).

YÊU CẦU ĐẦU RA (BẮT BUỘC):
- CHỈ trả về DUY NHẤT một mảng JSON hợp lệ, KHÔNG kèm markdown code fence, KHÔNG giải thích thêm, KHÔNG viết bất kỳ chữ nào trước hoặc sau mảng JSON.
- Số phần tử trong mảng trả về PHẢI khớp chính xác với số phần tử đầu vào (mỗi "id" xuất hiện đúng 1 lần).
- Định dạng mỗi phần tử: {{"id": "<id gốc>", "status": "true" | "suspect" | "false", "reason": "<lý do ngắn gọn; BẮT BUỘC điền khi status là 'suspect' hoặc 'false', để trống '' khi 'true'>"}}

Danh sách đầu vào (JSON):
{json.dumps(payload, ensure_ascii=False)}
"""
        return prompt

    def _parse_ai_json_array(self, raw_text):
        text = raw_text.strip()
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip())
        try:
            data = json.loads(text)
        except json.JSONDecodeError:
            match = re.search(r"\[.*\]", text, re.DOTALL)
            if not match:
                self.log("    [!] Không parse được JSON trả về từ AI.")
                return {}
            try:
                data = json.loads(match.group(0))
            except json.JSONDecodeError:
                self.log("    [!] JSON trả về từ AI bị sai định dạng.")
                return {}

        result_map = {}
        if isinstance(data, list):
            for entry in data:
                if not isinstance(entry, dict) or "id" not in entry:
                    continue
                status = str(entry.get("status", "")).strip().lower()
                status = status if status in ("true", "false", "suspect") else "unknown"
                result_map[str(entry["id"])] = {
                    "status": status,
                    "reason": entry.get("reason", "") or ""
                }
        return result_map

    def export_excel(self, collected_data, output_path):
        for item in collected_data:
            item.setdefault("DPV Match Code", "")
            item.setdefault("DPV Vacant", "")
            item.setdefault("DPV CMRA", "")
            item.setdefault("DPV No Stat", "")
            item.setdefault("Active", "")
            item.setdefault("DPV Footnotes", "")
            item.setdefault("Footnotes", "")
            item.setdefault("Trạng thái xử lý", "")
            item.setdefault("Lý do đáng ngờ", "")

        # false (lỗi) đứng trước, kế đến suspect (nghi ngờ - cần xem lại analysis),
        # rồi unknown (chưa kiểm tra được), cuối cùng true (chắc chắn đúng).
        status_priority = {"false": 0, "suspect": 1, "unknown": 2, "": 2, "true": 3}
        sorted_data = sorted(
            collected_data,
            key=lambda x: status_priority.get(str(x.get("Trạng thái xử lý", "")).lower(), 2)
        )

        columns = [
            "Reference Id", "Chuỗi đầu vào", "Chuỗi đầu ra",
            "DPV Match Code", "DPV Vacant", "DPV CMRA", "DPV No Stat", "Active",
            "DPV Footnotes", "Footnotes",
            "Trạng thái xử lý", "Lý do đáng ngờ",
        ]
        out_df = pd.DataFrame(sorted_data, columns=columns)

        # QUAN TRỌNG: xử lý/định dạng dữ liệu ra 1 FILE TẠM trước (không đụng tới output_path lúc
        # này), để việc file đích đang bị khoá (đang mở sẵn ở Excel) không làm hỏng hay chặn mất
        # công đoạn xử lý dữ liệu đã tốn công gọi API/AI. Chỉ bước SAO CHÉP/LƯU vào output_path ở
        # cuối mới cần cơ chế retry khi bị khoá.
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(tmp_fd)
        try:
            out_df.to_excel(tmp_path, index=False)

            import openpyxl
            wb = openpyxl.load_workbook(tmp_path)
            ws = wb.active

            fill_false = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            font_false = Font(color="9C0006", bold=True)
            fill_suspect = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
            font_suspect = Font(color="C65911", bold=True)
            fill_unknown = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            font_unknown = Font(color="9C6500")
            fill_true = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            font_true = Font(color="006100")

            status_col_idx = columns.index("Trạng thái xử lý") + 1

            for row_idx in range(2, ws.max_row + 1):
                status_val = str(ws.cell(row=row_idx, column=status_col_idx).value or "").strip().lower()
                if status_val == "false":
                    fill, font = fill_false, font_false
                elif status_val == "suspect":
                    fill, font = fill_suspect, font_suspect
                elif status_val == "true":
                    fill, font = fill_true, font_true
                else:
                    fill, font = fill_unknown, font_unknown

                for col_idx in range(1, len(columns) + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.fill = fill
                    if col_idx == status_col_idx:
                        cell.font = font

            for col_idx, col_name in enumerate(columns, start=1):
                max_len = max(
                    [len(str(col_name))] + [len(str(row.get(col_name, ""))) for row in sorted_data]
                )
                ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = min(max_len + 2, 60)

            ws.freeze_panes = "A2"

            # Bước lưu thực sự vào đích cuối - có retry nếu file đang bị khoá (đang mở ở chương
            # trình khác). Trả về đường dẫn thực tế đã lưu thành công (có thể khác output_path
            # nếu người dùng chọn đổi tên thay vì đóng file đang mở).
            final_path = self._save_workbook_with_retry(wb, output_path)
            return final_path
        finally:
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    def call_smarty_api(self, street_input, identity):
        # match lấy từ lựa chọn của người dùng trên GUI (mặc định "strict" - giống hệt kết
        # quả khi nhập thủ công vào Smarty, chỉ khớp dựa hoàn toàn trên dữ liệu gốc, không tự
        # "vá"/suy đoán). Người dùng có thể tự đổi sang "enhanced" nếu muốn Smarty khớp mạnh
        # tay hơn, nhưng "strict" là lựa chọn được khuyến khích.
        match_mode = self.match_mode_var.get() if self.match_mode_var.get() in ("strict", "enhanced") else "strict"
        params = {
            "key": identity["api_key"], "agent": "smarty (website:demo)",
            "match": match_mode, "candidates": "5", "geocode": "true",
            "license": "us-rooftop-geocoding-cloud", "street": street_input
        }

        proxy_url = identity.get("proxy_url")
        proxies = {"http": proxy_url, "https": proxy_url} if proxy_url else None

        try:
            # Xoá cookie trước mỗi lần gọi để mỗi request trông như 1 phiên hoàn toàn mới,
            # tránh Smarty gộp nhiều request lại thành 1 "client" duy nhất qua cookie.
            self.session.cookies.clear()
            res = self.session.get(
                API_URL, params=params, headers=identity["headers"],
                proxies=proxies, timeout=15
            )

            rate_limited = res.status_code == 429
            if res.status_code != 200:
                try:
                    msg = res.json()['errors'][0].get('message', 'Unknown Error')
                except Exception:
                    msg = f"HTTP {res.status_code}"
                if "too many requests" in msg.lower() or "rate limit" in msg.lower():
                    rate_limited = True
                return {
                    "result": f"Lỗi API: {msg}",
                    "analysis": {},
                    "rate_limited": rate_limited,
                    "status_code": res.status_code,
                    "success": False,
                }

            data = res.json()
            if isinstance(data, list) and len(data) > 0:
                c = data[0]
                dl = c.get("delivery_line_1", "")
                ll = c.get("last_line", "")
                result_string = f"{dl} {ll}".strip() if dl or ll else "Không tìm thấy delivery_line_1 hoặc last_line"
                analysis = c.get("analysis", {}) or {}
                return {"result": result_string, "analysis": analysis, "rate_limited": False, "status_code": 200, "success": True}
            return {"result": "Không tìm thấy kết quả từ Server", "analysis": {}, "rate_limited": False, "status_code": 200, "success": False}
        except requests.exceptions.RequestException as e:
            # Lỗi mạng/timeout/kết nối: đây thường là lỗi TẠM THỜI (proxy chập chờn, mạng lag...),
            # KHÔNG phải lỗi do dữ liệu địa chỉ sai -> phải được retry giống như rate-limit,
            # nếu không sẽ bỏ sót cả những dòng đáng ra hợp lệ chỉ vì 1 lần mạng bị giật.
            return {
                "result": f"Lỗi mạng / Timeout: {str(e)[:120]}",
                "analysis": {}, "rate_limited": False, "network_error": True, "status_code": None, "success": False,
            }
        except json.JSONDecodeError:
            return {
                "result": "Lỗi phân tích JSON",
                "analysis": {}, "rate_limited": False, "network_error": True, "status_code": None, "success": False,
            }

    def call_smarty_api_with_retry(self, street_input, max_retries, row_num):
        base_wait = 1.5

        for attempt in range(max_retries):
            if self.stop_requested:
                return {"result": "Bị dừng", "analysis": {}}

            # Mỗi LƯỢT GỌI (kể cả lần đầu, không chỉ khi bị chặn) đều dùng 1 danh tính
            # mới: proxy khác + Smarty key khác (nếu có) + User-Agent/Accept-Language khác.
            # Cách này rải traffic ra ngay từ đầu thay vì chỉ phản ứng sau khi đã bị chặn.
            identity = self._next_identity()
            outcome = self.call_smarty_api(street_input, identity)

            is_rate_limited = outcome.get("rate_limited")
            is_network_error = outcome.get("network_error")

            if is_rate_limited or is_network_error:
                proxy_info = f"proxy #{identity['index'] % len(self.proxy_list)}" if self.proxy_list else "không có proxy"
                reason_label = "Bị giới hạn request" if is_rate_limited else "Lỗi mạng/Timeout"
                self.log(
                    f"  [!] Dòng {row_num}: {reason_label} (lần {attempt + 1}/{max_retries}, {proxy_info}). "
                    f"Đang thử lại với danh tính khác..."
                )

                jitter = random.uniform(0.3, 1.3)
                if self.proxy_list:
                    # Đã có proxy mới ở lượt sau -> chỉ cần chờ ngắn ngẫu nhiên để tránh dồn request.
                    time.sleep(jitter)
                else:
                    # Không có proxy để đổi -> chỉ còn cách giãn thời gian theo cấp số nhân + jitter.
                    sleep_for = base_wait + jitter
                    self.log(f"      Không có proxy khả dụng, chờ {sleep_for:.1f}s trước khi thử danh tính kế tiếp...")
                    time.sleep(sleep_for)
                    base_wait = min(base_wait * 1.8, 30)
                continue

            return outcome

        return {
            "result": "Lỗi: Quá nhiều request (đã xoay vòng Proxy/User-Agent/Key nhưng vẫn bị giới hạn hoặc lỗi mạng liên tục)",
            "analysis": {},
            "success": False,
        }

    # ---- Bảng mã tra cứu chính thức của Smarty (US Street API) ----
    # Nguồn: https://www.smarty.com/docs/apis/us-street-api/reference

    # dpv_footnotes: chuỗi ghép các mã 2 ký tự KHÔNG có dấu phân cách (vd "AABB", "AAC1"...)
    # Với mỗi mã, mô tả + có chặn (block) hay không khi chạy chế độ khắt khe.
    DPV_FOOTNOTE_INFO = {
        "AA": ("Street/city/state/ZIP hợp lệ", False),
        "A1": ("Địa chỉ KHÔNG có trong dữ liệu USPS", True),
        "BB": ("Toàn bộ địa chỉ hợp lệ", False),
        "CC": ("Thông tin phụ (Apt/Suite...) không được nhận diện, KHÔNG bắt buộc để giao hàng", True),
        "C1": ("Thông tin phụ không được nhận diện, và BẮT BUỘC phải có để giao hàng", True),
        "F1": ("Địa chỉ quân sự/ngoại giao", True),
        "G1": ("Địa chỉ General Delivery (nhận tại bưu cục)", True),
        "M1": ("Thiếu số nhà (primary number)", True),
        "M3": ("Số nhà không hợp lệ", True),
        "N1": ("Thiếu thông tin phụ (Apt/Suite...) BẮT BUỘC để giao hàng", True),
        "PB": ("Địa chỉ dạng PO Box kiểu đường phố", True),
        "P1": ("Thiếu số hộp PO/RR/HC", True),
        "P3": ("Số hộp PO/RR/HC không hợp lệ", True),
        "RR": ("Địa chỉ xác nhận có thông tin hộp thư riêng (PMB)", True),
        "R1": ("Địa chỉ xác nhận KHÔNG có thông tin hộp thư riêng (PMB)", False),
        "R7": ("Địa chỉ hợp lệ nhưng KHÔNG được USPS giao hàng tận nhà", True),
        "TA": ("Số nhà chỉ khớp được sau khi bỏ ký tự chữ cái ở cuối", True),
        "U1": ("ZIP Code dạng 'unique' (đặc thù)", True),
    }

    # footnotes (khác dpv_footnotes): các mã 1 ký tự, ngăn cách bởi dấu '#', vd "L#", "H#L#"
    # Các mã thuần "thông tin", KHÔNG coi là cảnh báo (không chặn):
    FOOTNOTE_INFO_ONLY = {"N", "Q", "Y", "Z", "LL", "LI"}
    FOOTNOTE_INFO = {
        "A": "USPS đã sửa lại ZIP Code khác với ZIP đã nhập",
        "B": "USPS đã sửa lại chính tả tên thành phố/tiểu bang",
        "C": "Không xác định được ZIP (thiếu/sai city+state hoặc ZIP)",
        "D": "Địa chỉ KHÔNG có trong dữ liệu USPS (không có ZIP+4)",
        "E": "Nhiều bản ghi cùng chung 1 ZIP Code (mơ hồ)",
        "F": "KHÔNG tìm thấy địa chỉ như đã nhập trong thành phố/ZIP đã cho",
        "G": "Đã dùng dữ liệu từ trường Addressee để ghép vào địa chỉ",
        "H": "Thiếu số phụ (Apt/Suite...)",
        "I": "Dữ liệu địa chỉ không đủ/không chính xác để xác định 1 ZIP+4 duy nhất",
        "J": "Địa chỉ bị trùng 2 địa chỉ (dual address)",
        "K": "Chỉ khớp được sau khi ĐỔI hướng (N/S/E/W - cardinal rule match)",
        "L": "Một thành phần địa chỉ (suffix hoặc directional) đã bị THÊM/SỬA/XOÁ để khớp",
        "M": "Đã sửa lại chính tả tên đường (street name)",
        "N": "Đã chuẩn hoá viết tắt (vd STREET -> ST) - chỉ mang tính chuẩn hoá",
        "O": "Có nhiều ZIP+4 phù hợp, đã lấy mã thấp nhất",
        "P": "Địa chỉ có tên khác được ưu tiên hơn (better address exists)",
        "Q": "ZIP Code dạng 'unique' - chỉ mang tính thông tin",
        "R": "Chưa khớp được, nhưng EWS báo sẽ sớm có dữ liệu (địa chỉ mới)",
        "S": "Thông tin phụ (Apt/Suite...) không được USPS nhận diện",
        "T": "Trùng 'magnet street syndrome' - không đủ điều kiện trả ZIP+4 chuẩn CASS",
        "U": "Tên thành phố không chính thức (đã đổi sang tên chuẩn USPS)",
        "V": "KHÔNG xác thực được city/state khớp với ZIP đã cho",
        "W": "ZIP Code này KHÔNG có giao hàng tận đường, phải dùng PO Box/General Delivery",
        "X": "ZIP Code 'unique' dùng mã ZIP+4 mặc định - địa chỉ có thể không có thật",
        "Y": "Địa chỉ quân sự (military match) - chỉ mang tính thông tin",
        "Z": "Đã khớp qua ZIPMOVE (ZIP đã chuyển) - chỉ mang tính thông tin",
        "LL": "Được đánh dấu gửi sang LACSLink - chỉ mang tính thông tin",
        "LI": "Được đánh dấu gửi sang LACSLink - chỉ mang tính thông tin",
    }

    def _parse_dpv_footnotes(self, raw):
        """Tách chuỗi dpv_footnotes (ghép liên tục từng cặp 2 ký tự, vd 'AABBCC') thành list mã."""
        s = str(raw or "").strip().upper()
        return [s[i:i + 2] for i in range(0, len(s) - 1, 2)] if len(s) >= 2 else ([s] if s else [])

    def _parse_footnotes(self, raw):
        """Tách chuỗi footnotes (ngăn cách bởi '#', vd 'H#L#') thành list mã chữ cái."""
        s = str(raw or "").strip().upper()
        if not s:
            return []
        return [code for code in s.split("#") if code]

    # Các mẫu regex phát hiện dữ liệu ĐẦU VÀO (từ Excel gốc) bị lỗi encode / còn dính HTML entity.
    # Đây là lớp kiểm tra ĐỘC LẬP với Smarty: dù Smarty (đặc biệt ở chế độ "enhanced" - khớp mạnh
    # tay, có thể tự "vá" qua các cụm rác không đọc hiểu được) có trả về dpv_match_code=Y hay
    # không, một input còn dính rác kiểu này vẫn cho thấy DỮ LIỆU NGUỒN có vấn đề và không nên
    # được tin tưởng tuyệt đối - phải luôn đưa vào diện cần người kiểm tra lại.
    SUSPICIOUS_INPUT_PATTERNS = [
        (re.compile(r"&#\d+;"), "chứa mã HTML numeric entity (vd \"&#39;\")"),
        (re.compile(r"&#x[0-9a-fA-F]+;"), "chứa mã HTML hex entity (vd \"&#x27;\")"),
        (re.compile(r"&(amp|lt|gt|quot|apos|nbsp);", re.IGNORECASE), "chứa mã HTML named entity (vd \"&amp;\", \"&quot;\")"),
        (re.compile(r"[\uFFFD]"), "chứa ký tự lỗi encode (replacement character \\uFFFD)"),
        (re.compile(r"<[^>]+>"), "chứa thẻ HTML còn sót lại (vd \"<br>\")"),
    ]

    def _detect_input_anomalies(self, input_string):
        """Quét CHUỖI ĐẦU VÀO GỐC (chưa qua Smarty) để tìm dấu hiệu lỗi encode/HTML entity còn sót
        lại từ nguồn dữ liệu (Excel/CRM/website export...). Trả về danh sách lý do (rỗng nếu input
        sạch). Đây là kiểm tra độc lập với 'analysis' của Smarty - áp dụng NGAY CẢ KHI Smarty (đặc
        biệt match=enhanced) vẫn khớp được địa chỉ, vì bản thân input bẩn là 1 rủi ro về chất lượng
        dữ liệu, không nên tự động coi là 'true' chỉ vì Smarty đã cố gắng vá được."""
        reasons = []
        if not input_string:
            return reasons
        for pattern, desc in self.SUSPICIOUS_INPUT_PATTERNS:
            if pattern.search(input_string):
                reasons.append(f"Input gốc {desc} - dữ liệu nguồn có thể bị lỗi encode, cần kiểm tra lại thủ công")
        return reasons

    def _compute_analysis_flags(self, analysis):
        """Đối chiếu TOÀN BỘ object 'analysis' Smarty trả về (không chỉ vacant/CMRA/no_stat/
        active/match_code, mà còn cả 'dpv_footnotes' VÀ 'footnotes' - trường trước đây bị bỏ sót).
        Nguyên tắc KHẮT KHE: hễ còn BẤT KỲ dòng cảnh báo nào (kể cả các mã như 'L#' - đổi suffix/
        directional - vốn không ảnh hưởng dpv_match_code) thì vẫn bị đánh dấu 'suspect', không được
        coi là hợp lệ tuyệt đối chỉ vì dpv_match_code=Y. Trả về danh sách lý do (rỗng nếu thực sự
        sạch, không có bất kỳ cảnh báo nào)."""
        reasons = []
        if not analysis:
            return reasons

        vacant = str(analysis.get("dpv_vacant", "")).strip().upper()
        cmra = str(analysis.get("dpv_cmra", "")).strip().upper()
        no_stat = str(analysis.get("dpv_no_stat", "")).strip().upper()
        active = str(analysis.get("active", "")).strip().upper()
        match_code = str(analysis.get("dpv_match_code", "")).strip().upper()

        if vacant == "Y":
            reasons.append("USPS đánh dấu địa chỉ đang BỊ BỎ TRỐNG (dpv_vacant=Y)")
        if cmra == "Y":
            reasons.append("Địa chỉ là hộp thư nhận hộ CMRA, không phải địa chỉ thực (dpv_cmra=Y)")
        if no_stat == "Y":
            reasons.append("USPS không có đủ dữ liệu xác thực cho địa chỉ này (dpv_no_stat=Y)")
        if active == "N":
            reasons.append("Địa chỉ hiện KHÔNG hoạt động (active=N)")
        if match_code and match_code != "Y":
            reasons.append(f"Mã khớp DPV chưa xác nhận đầy đủ (dpv_match_code={match_code})")

        # --- MỚI: kiểm tra dpv_footnotes theo từng mã 2 ký tự ---
        for code in self._parse_dpv_footnotes(analysis.get("dpv_footnotes", "")):
            info = self.DPV_FOOTNOTE_INFO.get(code)
            if info is None:
                # Mã lạ / chưa có trong bảng tra -> vẫn đưa vào diện nghi ngờ để con người xem lại,
                # thay vì mặc định coi là an toàn.
                reasons.append(f"Mã dpv_footnotes lạ, cần xem lại thủ công ({code})")
                continue
            desc, should_block = info
            if should_block:
                reasons.append(f"Cảnh báo dpv_footnotes: {desc} ({code})")

        # --- MỚI (quan trọng nhất, khắc phục đúng lỗi bạn báo): kiểm tra 'footnotes' ---
        # Đây là trường trước đây KHÔNG được kiểm tra, nên các cảnh báo như 'L#'
        # (đổi suffix/directional để khớp) đã lọt qua dù dpv_match_code=Y.
        for code in self._parse_footnotes(analysis.get("footnotes", "")):
            if code in self.FOOTNOTE_INFO_ONLY:
                continue  # Mã thuần thông tin (N#, Q#, Y#, Z#, LL#, LI#) - không chặn
            desc = self.FOOTNOTE_INFO.get(code, "Mã footnotes lạ, cần xem lại thủ công")
            reasons.append(f"Cảnh báo footnotes: {desc} ({code}#)")

        return reasons

    def reset_ui(self):
        self.btn_start.config(state=tk.NORMAL, text="2. Bắt đầu xử lý & Xuất file (TXT + Excel)")
        self.btn_select_excel.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self.entry_delay.config(state=tk.NORMAL)
        self.rb_match_strict.config(state=tk.NORMAL)
        self.rb_match_enhanced.config(state=tk.NORMAL)
        self.chk_use_ai.config(state=tk.NORMAL)
        if self.use_ai_var.get():
            self._set_ai_inputs_state(tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = SmartyApp(root)
    root.mainloop()