# -*- coding: utf-8 -*-
"""
Cong cu tao file DEBIT hoan chinh (debit_editted) tu file debit goc + ct.xlsx
Giao dien co ban bang Tkinter - chay truc tiep tren may Windows.

Cach chay:
    python debit_tool_gui.py
(neu may chua co Python: cai Python 3 tu python.org, tick "Add to PATH" khi cai)

File engine.py phai nam CUNG THU MUC voi file nay.
"""
import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from engine import build_debit_editted, FLIGHT_TITLE_RE

TITLE_EXAMPLE = "297 3998 4210 / 25 AUG JFK / CI0794"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DEBIT TOOL: tool chỉ áp dụng cho các mawb đi mỹ và chưa tính giảm trừ nếu khách 370 > 500 kg:")
        self.geometry("720x620")
        self.resizable(True, True)

        pad = {"padx": 10, "pady": 6}

        tk.Label(self, text="1. File debit goc (.xlsx):", anchor="w").grid(row=0, column=0, sticky="w", **pad)

        self.debit_var = tk.StringVar()
        tk.Entry(self, textvariable=self.debit_var, width=70).grid(row=1, column=0, columnspan=2, sticky="we", padx=10)
        tk.Button(self, text="Chon file...", command=self.pick_debit).grid(row=1, column=2, padx=10)

        tk.Label(self, text="2. File ct.xlsx (trong thu muc chua cong thuc):", anchor="w").grid(row=2, column=0, sticky="w", **pad)
        self.ct_var = tk.StringVar()
        tk.Entry(self, textvariable=self.ct_var, width=70).grid(row=3, column=0, columnspan=2, sticky="we", padx=10)
        tk.Button(self, text="Chon file...", command=self.pick_ct).grid(row=3, column=2, padx=10)

        tk.Label(
            self,
            text="3. Tieu de chuyen bay day du (bat buoc dung dinh dang mau):",
            anchor="w",
        ).grid(row=4, column=0, sticky="w", **pad)
        tk.Label(
            self,
            text=f'    Vi du: "{TITLE_EXAMPLE}"   (MAWB / NGAY THANG SAN BAY DEN / MA CI)',
            anchor="w", fg="#555555",
        ).grid(row=5, column=0, columnspan=3, sticky="w", padx=10)
        self.title_var = tk.StringVar()
        self.title_entry = tk.Entry(self, textvariable=self.title_var, width=70)
        self.title_entry.grid(row=6, column=0, columnspan=2, sticky="we", padx=10)
        self.title_status = tk.Label(self, text="", anchor="w", fg="#b71c1c")
        self.title_status.grid(row=7, column=0, columnspan=3, sticky="w", padx=10)
        self.title_var.trace_add("write", self._on_title_change)

        tk.Label(
            self,
            text="Bo trong o tren neu muon tool tu dong tao tieu de tu du lieu (chi nhap ma CI ben duoi):",
            anchor="w", fg="#555555",
        ).grid(row=8, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 0))
        tk.Label(self, text="   Ma CI / so invoice (vd: CI0792) - tuy chon:", anchor="w").grid(
            row=9, column=0, sticky="w", **pad
        )
        self.ci_var = tk.StringVar()
        tk.Entry(self, textvariable=self.ci_var, width=30).grid(row=10, column=0, sticky="w", padx=10)

        tk.Label(
            self,
            text='Ma EC hien thi o dong tieu de bang (vd: "EC260802-") - tuy chon, bo trong de tool tu dong tao:',
            anchor="w",
        ).grid(row=11, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 0))
        self.ec_var = tk.StringVar()
        tk.Entry(self, textvariable=self.ec_var, width=30).grid(row=12, column=0, sticky="w", padx=10)

        tk.Label(self, text="4. Noi luu file ket qua (debit_editted.xlsx):", anchor="w").grid(row=13, column=0, sticky="w", **pad)
        self.out_var = tk.StringVar()
        tk.Entry(self, textvariable=self.out_var, width=70).grid(row=14, column=0, columnspan=2, sticky="we", padx=10)
        tk.Button(self, text="Chon noi luu...", command=self.pick_out).grid(row=14, column=2, padx=10)

        self.run_btn = tk.Button(self, text="CHAY - Tao file DEBIT hoan chinh", command=self.run_clicked,
                                  bg="#2e7d32", fg="white", font=("Segoe UI", 11, "bold"), height=2)
        self.run_btn.grid(row=15, column=0, columnspan=3, sticky="we", padx=10, pady=14)

        tk.Label(self, text="Nhat ky xu ly:", anchor="w").grid(row=16, column=0, sticky="w", padx=10)
        self.log_box = scrolledtext.ScrolledText(self, height=16)
        self.log_box.grid(row=17, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(17, weight=1)

    def _on_title_change(self, *args):
        val = self.title_var.get().strip()
        if not val:
            self.title_status.config(text="")
        elif FLIGHT_TITLE_RE.match(val):
            self.title_status.config(text="Dinh dang hop le.", fg="#2e7d32")
        else:
            self.title_status.config(
                text=f'Dinh dang chua dung. Vui long nhap theo mau: "{TITLE_EXAMPLE}"',
                fg="#b71c1c",
            )

    def pick_debit(self):
        p = filedialog.askopenfilename(title="Chon file debit goc", filetypes=[("Excel files", "*.xlsx")])
        if p:
            self.debit_var.set(p)
            if not self.out_var.get():
                base = os.path.splitext(p)[0]
                self.out_var.set(base + "_editted.xlsx")

    def pick_ct(self):
        p = filedialog.askopenfilename(title="Chon file ct.xlsx", filetypes=[("Excel files", "*.xlsx")])
        if p:
            self.ct_var.set(p)

    def pick_out(self):
        p = filedialog.asksaveasfilename(title="Luu file ket qua", defaultextension=".xlsx",
                                          filetypes=[("Excel files", "*.xlsx")])
        if p:
            self.out_var.set(p)

    def log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.update_idletasks()

    def confirm_zero_fields(self, risky_fields):
        """confirm_callback cho engine.build_debit_editted: duoc goi CHI KHI
        mot hoac nhieu truong trong ZERO_CHECK_FIELDS khong chac chan bang 0
        (tong khac 0, hoac co gia tri khong doc duoc dang so) - tuc la neu
        tu dong xoa cot nay khoi output co the gay sai lech du lieu.

        Hien hop thoai canh bao liet ke cac truong do, cho nguoi dung chon:
          - "Yes" (Co)   -> tra ve True: engine se BO QUA tao file lan nay
                            (khong xu ly), nhung tool (cua so GUI) van tiep
                            tuc chay binh thuong de nguoi dung kiem tra lai
                            du lieu dau vao va thu lai.
          - "No" (Khong) -> tra ve False: engine se raise loi va DUNG hoan
                            toan qua trinh xu ly lan nay, khong tao file.
        """
        msg = (
            "Phat hien cac truong ZERO_CHECK_FIELDS co du lieu KHONG chac "
            "chan bang 0 (neu tu dong bo qua/xoa co the lam sai lech du lieu):\n\n"
            + "\n".join(f"- {f}" for f in risky_fields)
            + "\n\nCac truong nay se KHONG bi tu dong xoa khoi output.\n\n"
            "Ban co muon TIEP TUC khong?\n"
            "  - Co (Yes): BO QUA tao file lan nay (chua tao debit_editted), "
            "vui long kiem tra lai du lieu dau vao (file debit goc/ct.xlsx) "
            "roi chay lai. Cong cu se KHONG dong, ban co the thu lai ngay.\n"
            "  - Khong (No): DUNG han qua trinh xu ly lan nay, khong tao file."
        )
        return messagebox.askyesno(
            "Canh bao: du lieu bat thuong o ZERO_CHECK_FIELDS", msg
        )

    def run_clicked(self):
        debit_path = self.debit_var.get().strip()
        ct_path = self.ct_var.get().strip()
        out_path = self.out_var.get().strip()
        ci_code = self.ci_var.get().strip()
        flight_title = self.title_var.get().strip()
        ec_code = self.ec_var.get().strip()

        if not debit_path or not os.path.isfile(debit_path):
            messagebox.showerror("Loi", "Vui long chon file debit goc hop le.")
            return
        if not ct_path or not os.path.isfile(ct_path):
            messagebox.showerror("Loi", "Vui long chon file ct.xlsx hop le.")
            return
        if not out_path:
            messagebox.showerror("Loi", "Vui long chon noi luu file ket qua.")
            return

        if flight_title and not FLIGHT_TITLE_RE.match(flight_title):
            messagebox.showerror(
                "Loi",
                "Tieu de chuyen bay chua dung dinh dang.\n"
                f'Vui long nhap theo mau: "{TITLE_EXAMPLE}"\n'
                "(MAWB / NGAY THANG SAN BAY DEN / MA CI)\n"
                "Hoac de trong o nay neu muon tool tu dong tao tieu de.",
            )
            return

        self.run_btn.config(state="disabled", text="Dang xu ly...")
        self.log_box.delete("1.0", tk.END)

        def worker():
            try:
                result = build_debit_editted(
                    debit_path, ct_path, out_path,
                    ci_code=ci_code,
                    flight_title=(flight_title or None),
                    ec_code=(ec_code or None),
                    progress=self.log,
                    confirm_callback=self.confirm_zero_fields,
                )
                if result is None:
                    # Nguoi dung da chon "Co" trong hop thoai canh bao
                    # ZERO_CHECK_FIELDS -> lan nay KHONG co file nao duoc
                    # tao. Khong duoc bao "XONG" trong truong hop nay.
                    self.log("\n=> DA BO QUA: chua tao file debit_editted lan nay "
                              "vi phat hien du lieu bat thuong o cac truong "
                              "ZERO_CHECK_FIELDS (xem canh bao ben tren).")
                    self.log("Vui long kiem tra lai file debit goc / ct.xlsx roi bam "
                              "CHAY lai.")
                    messagebox.showwarning(
                        "Chua tao file",
                        "Da bo qua buoc tao file lan nay do phat hien du lieu bat "
                        "thuong o cac truong ZERO_CHECK_FIELDS.\n"
                        "Vui long kiem tra lai du lieu dau vao roi thu lai.",
                    )
                else:
                    self.log("\n=> XONG! File da luu tai:\n" + result)
                    self.log("\nLuu y: mo file bang Excel, bam Ctrl+Alt+F9 (hoac File > Options > "
                              "cho phep tinh toan lai) de Excel tinh lai toan bo cong thuc.")
                    messagebox.showinfo("Hoan tat", "Da tao file DEBIT hoan chinh thanh cong!")
            except Exception as e:
                self.log("\nLOI: " + str(e))
                self.log(traceback.format_exc())
                messagebox.showerror("Loi", f"Co loi xay ra:\n{e}")
            finally:
                self.run_btn.config(state="normal", text="CHAY - Tao file DEBIT hoan chinh")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()