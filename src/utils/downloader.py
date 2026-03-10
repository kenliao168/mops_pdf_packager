import os
import datetime
import requests
import urllib3
from src.scrapers.ebook_scraper import download_mops_pdf
from src.scrapers.mopsov_scraper import download_briefing_selenium, download_financials_selenium, download_affiliated_selenium
from src.scrapers.esg_scraper import download_esg_report
from src.scrapers.prospectus_scraper import download_prospectus

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def get_recent_years(count=2):
    current_year = datetime.datetime.now().year - 1911
    # 預防此時年報還沒出，往前回推
    return [current_year, current_year - 1][:count]


def lookup_company_name(ticker):
    """
    查詢公司簡稱 (例如 2330 -> 台積電)
    嘗試多個來源:
      1. ESG 數位平台 API
      2. TWSE 公司查詢 API
      3. fallback: 空字串
    """
    ticker_str = str(ticker)

    # 方法 1: ESG 數位平台 (最可靠，因為有 shortName)
    for market_type in [0, 1]:
        try:
            current_year = datetime.datetime.now().year
            res = requests.post(
                "https://esggenplus.twse.com.tw/api/api/MopsSustainReport/data",
                json={
                    "companyCodeList": [ticker_str],
                    "year": current_year,
                    "industryNameList": [],
                    "marketType": market_type,
                    "industryName": "all",
                    "companyCode": ticker_str,
                },
                headers={
                    "Content-Type": "application/json",
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Referer": "https://esggenplus.twse.com.tw/inquiry/report",
                    "Origin": "https://esggenplus.twse.com.tw",
                },
                verify=False,
                timeout=10,
            )
            if res.status_code == 200:
                data = res.json()
                if data.get("success") and data.get("data"):
                    name = data["data"][0].get("shortName", "")
                    if name:
                        print(f"[公司查詢] {ticker_str} = {name}")
                        return name
        except Exception:
            pass

    # 方法 1b: ESG 去年
    for market_type in [0, 1]:
        try:
            last_year = datetime.datetime.now().year - 1
            res = requests.post(
                "https://esggenplus.twse.com.tw/api/api/MopsSustainReport/data",
                json={
                    "companyCodeList": [ticker_str],
                    "year": last_year,
                    "industryNameList": [],
                    "marketType": market_type,
                    "industryName": "all",
                    "companyCode": ticker_str,
                },
                headers={
                    "Content-Type": "application/json",
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Referer": "https://esggenplus.twse.com.tw/inquiry/report",
                    "Origin": "https://esggenplus.twse.com.tw",
                },
                verify=False,
                timeout=10,
            )
            if res.status_code == 200:
                data = res.json()
                if data.get("success") and data.get("data"):
                    name = data["data"][0].get("shortName", "")
                    if name:
                        print(f"[公司查詢] {ticker_str} = {name}")
                        return name
        except Exception:
            pass

    # 方法 2: MOPS 公開資訊觀測站 (公司基本資料)
    try:
        res = requests.post(
            "https://mopsov.twse.com.tw/mops/web/ajax_t05st03",
            data={
                "encodeURIComponent": "1",
                "step": "1",
                "firstin": "1",
                "off": "1",
                "queryName": "co_id",
                "inpuType": "co_id",
                "TYPEK": "all",
                "co_id": ticker_str,
            },
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
            },
            verify=False,
            timeout=10,
        )
        if res.status_code == 200:
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(res.text, "html.parser")
            # 找 "公司簡稱" 欄位
            for td in soup.find_all("td"):
                text = td.get_text(strip=True)
                if "公司簡稱" in text:
                    next_td = td.find_next_sibling("td")
                    if next_td:
                        name = next_td.get_text(strip=True)
                        if name:
                            print(f"[公司查詢] {ticker_str} = {name} (MOPS)")
                            return name
    except Exception:
        pass

    print(f"[公司查詢] 無法查到 {ticker_str} 的公司名稱，將只使用代碼。")
    return ""


def get_desktop_path():
    """取得使用者真實桌面路徑 (考慮 OneDrive 等重導向)"""
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        )
        desktop_path, _ = winreg.QueryValueEx(key, 'Desktop')
        return desktop_path
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")


class MOPSDownloader:
    def __init__(self, ticker, target_year=None):
        self.ticker = str(ticker)
        # 如果使用者沒指定年份，預設抓最近兩年 (先試今年，沒有再抓去年)
        if target_year is None:
            self.recent_years = get_recent_years(count=2)
        else:
            self.recent_years = [target_year]

        # 查詢公司名稱
        self.company_name = lookup_company_name(self.ticker)

        # 建立輸出資料夾: "{ticker} {公司名} NotebookLM上傳文件"
        desktop_path = get_desktop_path()
        if self.company_name:
            folder_name = f"{self.ticker} {self.company_name} NotebookLM上傳文件"
        else:
            folder_name = f"{self.ticker} NotebookLM上傳文件"

        self.save_dir = os.path.join(desktop_path, folder_name)

    def run(self):
        print(f"=== 開始打包 {self.ticker} 資料至 {self.save_dir} ===")
        os.makedirs(self.save_dir, exist_ok=True)

        # 定義我們要抓取的各項數量目標
        target_counts = {
            "年報": 5,
            "財報": 999, # 改為無限大，由年份來決定停止
            "關係企業三書表": 5,
            "法說會簡報": 5
        }
        fetched_counts = {k: 0 for k in target_counts}
        seen_paths = set()

        # 由於財報一年有四期，5 期財報通常只需跨 2 年。
        # 但年報一年只有一期，5 年年報需要跨至少 5~6 年。
        # 因此我們擴充 recent_years 至少往前推 6 年
        current_year = datetime.datetime.now().year - 1911
        years_to_search = [current_year - i for i in range(0, 7)]

        for year in years_to_search:
            # 判斷是否全部抓滿
            if all(fetched_counts[k] >= target_counts[k] for k in target_counts):
                break

            print(f"\n==> 正在搜尋 {year} 年度資料...")

            # 1. 抓取年報 (每年1期)
            if fetched_counts["年報"] < target_counts["年報"]:
                try:
                    path = download_mops_pdf(self.ticker, year, "年報", self.save_dir)
                    if path:
                        fetched_counts["年報"] += 1
                except Exception as e:
                    print(f"Warning: {year} 年報下載失敗。({e})")

            # 2. 抓取財報 (最新年度前四期全拿，往前推五年的歷史年度只拿第四季(Q4)而且個體與合併全收)
            if fetched_counts["財報"] < target_counts["財報"]:
                try:
                    # 使用 download_all=True 抓取該年度所有財報
                    paths = download_mops_pdf(self.ticker, year, "財報", self.save_dir, download_all=True)
                    if paths:
                        for path in reversed(paths):
                            filename = os.path.basename(path)

                            # 判斷這個報告是哪一季。檔名通常是 2330_113_202404_財報_合併.pdf -> '04_'
                            is_q4 = ("_04_" in filename or "_12_" in filename or "04_財報" in filename)

                            # 如果這份報告不是第一年的最新報告 (亦即我們已經抓滿了最新的一整年)，那歷史年度就強迫只收第四季
                            if fetched_counts["財報"] >= 4 and not is_q4:
                                try:
                                    os.remove(path)
                                except:
                                    pass
                                continue

                            fetched_counts["財報"] += 1
                except Exception as e:
                    print(f"Warning: {year} 財報下載失敗。({e})")

            # 3. 抓取關係企業三書表 (每年1期)
            if fetched_counts["關係企業三書表"] < target_counts["關係企業三書表"]:
                try:
                    path = download_mops_pdf(self.ticker, year, "關係企業三書表", self.save_dir)
                    if path:
                        fetched_counts["關係企業三書表"] += 1
                except Exception as e:
                    print(f"Warning: {year} 關係企業三書表下載失敗。({e})")

            # 4. 抓取法說會簡報 (每年可能多期，我們算次數，抓5次最新)
            if fetched_counts["法說會簡報"] < target_counts["法說會簡報"]:
                try:
                    from src.scrapers.briefing_scraper import download_briefing_pdf
                    paths = download_briefing_pdf(self.ticker, year, self.save_dir, download_all=True)
                    if paths:
                        # 從最新的開始算
                        for path in reversed(paths):
                            if path in seen_paths:
                                continue # 已經算進之前的數量了，不重複計算

                            if fetched_counts["法說會簡報"] < target_counts["法說會簡報"]:
                                fetched_counts["法說會簡報"] += 1
                                seen_paths.add(path)
                            else:
                                try:
                                    os.remove(path)
                                except:
                                    pass
                except Exception as e:
                    print(f"Warning: {year} 法說會簡報下載失敗。({e})")

        # 5. 抓取永續報告書 (ESG Report) — 不受年度迴圈限制，一次搜尋全部
        print(f"\n==> 正在搜尋 {self.ticker} 永續報告書 (ESG 數位平台)...")
        try:
            esg_paths = download_esg_report(self.ticker, self.save_dir, max_reports=3)
            if esg_paths:
                print(f"成功下載 {len(esg_paths)} 份永續報告書。")
            else:
                print(f"Warning: 無法自動下載 {self.ticker} 的永續報告書。可手動至 https://esggenplus.twse.com.tw/inquiry/report 下載。")
        except Exception as e:
            print(f"Warning: 永續報告書下載失敗。({e})")
            print(f"您可以手動至 https://esggenplus.twse.com.tw/inquiry/report 搜尋 {self.ticker} 下載。")

        # 6. 抓取最近一期公開說明書
        print(f"\n==> 正在搜尋 {self.ticker} 最近一期公開說明書...")
        try:
            prospectus_paths = download_prospectus(self.ticker, self.save_dir, max_reports=1)
            if prospectus_paths:
                print(f"成功下載 {len(prospectus_paths)} 份公開說明書。")
            else:
                print(f"Warning: 無法自動下載 {self.ticker} 的公開說明書。可手動至 https://mopsov.twse.com.tw/mops/web/t57sb01_q3 下載。")
        except Exception as e:
            print(f"Warning: 公開說明書下載失敗。({e})")
            print(f"您可以手動至 https://mopsov.twse.com.tw/mops/web/t57sb01_q3 搜尋 {self.ticker} 下載。")

        print("\n=== 打包完成 ===")
        print(f"所有報告已存放在 {os.path.abspath(self.save_dir)} 資料夾中。")
        print("您可以直接將這個資料夾內的 PDF 放入 Google NotebookLM 中進行分析。")
