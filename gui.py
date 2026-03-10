"""
MOPS PDF Packager — GUI 版本
讓不懂程式的人也能一鍵下載台股公開資訊報告。
使用 tkinter (Python 內建)，不需額外安裝任何 GUI 套件。
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import sys
import os
import io

# 確保能找到 src 模組
if getattr(sys, 'frozen', False):
    # PyInstaller 打包後的路徑
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from src.utils.downloader import MOPSDownloader


class RedirectText(io.StringIO):
    """將 print 的輸出導向到 tkinter Text widget"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.after(0, self._append, string)

    def _append(self, string):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass


class MOPSApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MOPS PDF Packager")
        self.root.geometry("700x520")
        self.root.resizable(True, True)
        self.root.configure(bg="#f5f5f5")

        # 設定 icon (如果有的話)
        try:
            self.root.iconbitmap(default="")
        except:
            pass

        self._build_ui()
        self.is_running = False

    def _build_ui(self):
        # ===== 標題 =====
        title_frame = tk.Frame(self.root, bg="#1a5276", pady=12)
        title_frame.pack(fill=tk.X)

        tk.Label(
            title_frame,
            text="MOPS PDF Packager",
            font=("Microsoft JhengHei", 18, "bold"),
            fg="white", bg="#1a5276"
        ).pack()

        tk.Label(
            title_frame,
            text="台股公開資訊報告 一鍵下載工具",
            font=("Microsoft JhengHei", 10),
            fg="#d5d8dc", bg="#1a5276"
        ).pack()

        # ===== 輸入區 =====
        input_frame = tk.Frame(self.root, bg="#f5f5f5", pady=15, padx=20)
        input_frame.pack(fill=tk.X)

        tk.Label(
            input_frame,
            text="股票代碼:",
            font=("Microsoft JhengHei", 12),
            bg="#f5f5f5"
        ).pack(side=tk.LEFT)

        self.ticker_var = tk.StringVar()
        self.ticker_entry = tk.Entry(
            input_frame,
            textvariable=self.ticker_var,
            font=("Consolas", 14),
            width=10,
            justify="center"
        )
        self.ticker_entry.pack(side=tk.LEFT, padx=(10, 20))
        self.ticker_entry.focus_set()

        # Enter 鍵觸發下載
        self.ticker_entry.bind('<Return>', lambda e: self._start_download())

        self.download_btn = tk.Button(
            input_frame,
            text="開始下載",
            font=("Microsoft JhengHei", 11, "bold"),
            bg="#2980b9", fg="white",
            activebackground="#1f618d", activeforeground="white",
            relief="flat", padx=20, pady=5,
            cursor="hand2",
            command=self._start_download
        )
        self.download_btn.pack(side=tk.LEFT)

        # ===== 下載內容說明 =====
        info_frame = tk.Frame(self.root, bg="#f5f5f5", padx=20)
        info_frame.pack(fill=tk.X)

        tk.Label(
            info_frame,
            text="下載內容: 年報 / 財報 / 關係企業三書表 / 法說會簡報 / ESG永續報告書 / 公開說明書",
            font=("Microsoft JhengHei", 9),
            fg="#7f8c8d", bg="#f5f5f5"
        ).pack(anchor="w")

        tk.Label(
            info_frame,
            text="儲存位置: 桌面 / {代碼} {公司名} NotebookLM上傳文件",
            font=("Microsoft JhengHei", 9),
            fg="#7f8c8d", bg="#f5f5f5"
        ).pack(anchor="w")

        # ===== 進度區 =====
        progress_frame = tk.Frame(self.root, bg="#f5f5f5", padx=20, pady=10)
        progress_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            progress_frame,
            text="下載進度:",
            font=("Microsoft JhengHei", 10),
            bg="#f5f5f5"
        ).pack(anchor="w")

        self.log_text = scrolledtext.ScrolledText(
            progress_frame,
            font=("Consolas", 9),
            bg="#2c3e50", fg="#ecf0f1",
            insertbackground="white",
            state='disabled',
            wrap=tk.WORD,
            height=15
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        # ===== 底部 =====
        footer = tk.Frame(self.root, bg="#f5f5f5", pady=5)
        footer.pack(fill=tk.X)

        tk.Label(
            footer,
            text="資料來源: 公開資訊觀測站 (MOPS) / ESG 數位平台 (TWSE)",
            font=("Microsoft JhengHei", 8),
            fg="#bdc3c7", bg="#f5f5f5"
        ).pack()

    def _start_download(self):
        ticker = self.ticker_var.get().strip()

        if not ticker:
            messagebox.showwarning("提醒", "請輸入股票代碼！")
            self.ticker_entry.focus_set()
            return

        if not ticker.isdigit():
            messagebox.showwarning("提醒", "股票代碼應為數字 (例如: 2330)")
            self.ticker_entry.focus_set()
            return

        if self.is_running:
            messagebox.showinfo("進行中", "下載正在進行中，請等待完成。")
            return

        # 清空 log
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')

        # 開始下載 (另開 thread 避免 GUI 卡住)
        self.is_running = True
        self.download_btn.configure(state='disabled', text="下載中...", bg="#7f8c8d")
        self.ticker_entry.configure(state='disabled')

        thread = threading.Thread(target=self._run_download, args=(ticker,), daemon=True)
        thread.start()

    def _run_download(self, ticker):
        # 將 stdout/stderr 導向到 GUI log
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        redirector = RedirectText(self.log_text)
        sys.stdout = redirector
        sys.stderr = redirector

        try:
            downloader = MOPSDownloader(ticker=ticker)
            downloader.run()
            self.root.after(0, lambda: messagebox.showinfo(
                "完成",
                f"{ticker} 報告下載完成！\n\n儲存位置:\n{downloader.save_dir}"
            ))
        except Exception as e:
            print(f"\n*** 錯誤: {e} ***")
            self.root.after(0, lambda: messagebox.showerror("錯誤", f"下載過程發生錯誤:\n{e}"))
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.is_running = False
            self.root.after(0, self._reset_ui)

    def _reset_ui(self):
        self.download_btn.configure(state='normal', text="開始下載", bg="#2980b9")
        self.ticker_entry.configure(state='normal')
        self.ticker_entry.focus_set()


def main():
    root = tk.Tk()
    app = MOPSApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
