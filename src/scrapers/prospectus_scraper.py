import requests
from bs4 import BeautifulSoup
import re
import os
import urllib3
import urllib.request
import urllib.parse
import ssl
import time
import random

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

SEARCH_URL = "https://doc.twse.com.tw/server-java/t57sb01"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
}

# 建立不驗證 SSL 的 context (doc.twse.com.tw 憑證有時會有問題)
SSL_CTX = ssl.create_default_context()
SSL_CTX.check_hostname = False
SSL_CTX.verify_mode = ssl.CERT_NONE


def _post_urllib(url, data_dict):
    """
    用 urllib.request 做 POST，避免 requests 遇到
    doc.twse.com.tw step=9 回傳 BadStatusLine 的問題。
    """
    encoded = urllib.parse.urlencode(data_dict).encode("utf-8")
    req = urllib.request.Request(url, data=encoded, headers=HEADERS)
    resp = urllib.request.urlopen(req, context=SSL_CTX, timeout=30)
    return resp.read()


def download_prospectus(ticker, save_dir, max_reports=1):
    """
    從 doc.twse.com.tw 下載最近一期（或多期）公開說明書 PDF。
    mtype=B 代表公開說明書。

    Args:
        ticker: 股票代號 (e.g. "2330")
        save_dir: 儲存目錄
        max_reports: 最多下載幾份 (預設 1，即只抓最新一期)

    Returns:
        list: 已下載的檔案路徑列表
    """
    ticker_str = str(ticker)

    # Step 1: 查詢公開說明書列表 (year 留空 = 全部年度)
    payload_search = {
        "step": "1",
        "colorchg": "1",
        "co_id": ticker_str,
        "year": "",
        "seamon": "",
        "mtype": "B",
    }

    print(f"[公開說明書] 正在查詢 {ticker_str} 的公開說明書...")

    try:
        time.sleep(random.uniform(1, 2))
        raw_html = _post_urllib(SEARCH_URL, payload_search)
        print(f"[公開說明書] 查詢成功，回應大小: {len(raw_html)} bytes")
    except Exception as e:
        print(f"[公開說明書] 查詢 {ticker_str} 失敗: {e}")
        return []

    soup = BeautifulSoup(raw_html, "html.parser", from_encoding="big5")

    # 解析所有檔案連結
    links = soup.find_all("a")
    all_files = []
    for a in links:
        text = a.text.strip()
        if text.endswith(".pdf") or text.endswith(".zip") or text.endswith(".doc"):
            all_files.append(text)

    print(f"[公開說明書] 共找到 {len(all_files)} 個檔案連結")

    if not all_files:
        print(f"[公開說明書] 找不到 {ticker_str} 的公開說明書資料。")
        return []

    # 只取 .pdf 檔案
    pdf_files = [f for f in all_files if f.endswith(".pdf")]
    if not pdf_files:
        pdf_files = all_files

    print(f"[公開說明書] 共 {len(pdf_files)} 個 PDF，最新: {pdf_files[-1]}")

    # 取最新的幾份
    target_files = pdf_files[-max_reports:]
    target_files.reverse()

    saved_paths = []
    for target_filename in target_files:
        try:
            time.sleep(random.uniform(2, 4))
            print(f"[公開說明書] 準備下載: {target_filename}")

            # Step 2: POST step=9 取得真實 PDF 轉址頁
            payload_download = {
                "step": "9",
                "kind": "B",
                "co_id": ticker_str,
                "filename": target_filename,
            }
            raw_jump = _post_urllib(SEARCH_URL, payload_download)

            soup_jump = BeautifulSoup(raw_jump, "html.parser", from_encoding="big5")

            # 解析轉址中真實下載路徑 <a href="/pdf/...">
            pdf_a = soup_jump.find("a", href=re.compile(r"/pdf/"))
            if not pdf_a:
                pdf_a = soup_jump.find("a", href=re.compile(r"\.pdf"))
                if not pdf_a:
                    jump_text = soup_jump.get_text(strip=True)[:300]
                    print(f"[公開說明書] 解析轉址失敗: {jump_text}")
                    continue

            href = pdf_a["href"]
            if href.startswith("http"):
                real_pdf_url = href
            else:
                real_pdf_url = "https://doc.twse.com.tw" + href

            # Step 3: 下載 PDF
            os.makedirs(save_dir, exist_ok=True)

            month_match = re.search(r"^(\d+)_", target_filename)
            period_str = month_match.group(1) if month_match else ""

            readable_filename = f"{ticker_str}_公開說明書_{period_str}.pdf"
            save_path = os.path.join(save_dir, readable_filename)

            print(f"[公開說明書] 正在下載 {real_pdf_url} ...")

            req_pdf = urllib.request.Request(real_pdf_url, headers=HEADERS)
            resp_pdf = urllib.request.urlopen(req_pdf, context=SSL_CTX, timeout=120)
            pdf_bytes = resp_pdf.read()

            with open(save_path, "wb") as f:
                f.write(pdf_bytes)

            size_mb = len(pdf_bytes) / (1024 * 1024)
            print(f"[公開說明書] 已成功儲存 {readable_filename} ({size_mb:.1f} MB)")
            saved_paths.append(save_path)

        except Exception as e:
            print(f"[公開說明書] 下載 {target_filename} 失敗: {e}")
            import traceback
            traceback.print_exc()
            continue

    return saved_paths


if __name__ == "__main__":
    download_prospectus("2330", "./test_prospectus", max_reports=1)
