import requests
from bs4 import BeautifulSoup
import re
import urllib3
import time
import random

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 新版的 MOPS 電子書查詢網址
MOPS_EBOOK_URL = "https://mops.twse.com.tw/mops/web/ajax_t164sb03"

import requests
import datetime
import os
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

import requests
import datetime
import os
import urllib3
import re
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def get_recent_years(count=2):
    """取得最近幾年的民國年分"""
    current_year = datetime.datetime.now().year - 1911
    # 通常今年中的時候才會出完去年的年報，因此我們從去年開始抓
    return [current_year - i for i in range(1, count + 2)]

def download_mops_pdf(ticker, year, doc_type, save_dir, download_all=False):
    """
    透過 MOPS /server-java/t57sb01 介面下載 PDF。
    """
    search_url = 'https://doc.twse.com.tw/server-java/t57sb01'
    session = requests.Session()
    
    search_mtypes = ['F']
    if doc_type == "財報":
        search_mtypes = ['A']
    elif doc_type == "關係企業三書表":
        search_mtypes = ['K'] # 關係企業三書表(K)
        
    for mtype in search_mtypes:
        payload_search = {
            'step': '1', 
            'colorchg': '1', 
            'co_id': str(ticker), 
            'year': str(year), 
            'mtype': mtype
        }
        
        try:
            res_search = session.post(search_url, data=payload_search, verify=False, timeout=15)
            res_search.raise_for_status()
            
            soup = BeautifulSoup(res_search.content, 'html.parser', from_encoding='big5')
            links = soup.find_all('a')
            
            target_filenames = []
            
            if doc_type == "年報":
                # 尋找 F04.pdf (股東會年報)
                for a in links:
                    if a.text.strip().endswith("F04.pdf"):
                        target_filenames.append(a.text.strip())
            elif doc_type == "財報":
                # 財報為 mtype=A，合併財報 AI1，個體 AI3/AI2
                for a in links:
                    txt = a.text.strip()
                    if "AI1.pdf" in txt or "AI3.pdf" in txt or "AI2.pdf" in txt:
                        target_filenames.append(txt)
            elif doc_type == "關係企業三書表":
                # 在 K 專區，所有附檔皆視為三書表相關
                for a in links:
                    if ".pdf" in a.text.lower():
                        target_filenames.append(a.text.strip())

            if target_filenames:
                break # 找到了就不用再試下一個 mtype
        except requests.exceptions.RequestException as e:
            print(f"Request Error for {ticker} ({year}) mtype={mtype}: {e}")
            pass
            
    if not target_filenames:
        print(f"Warning: 無法在 {year} 年度的 MOPS 中找到 {ticker} 的 {doc_type}。")
        return [] if download_all else None
        
    try:
        # 若只要求最新的一份，我們抓字串排序最後的 (代表最新月份或最後一季)
        target_filenames.sort()
        if not download_all:
            target_filenames = [target_filenames[-1]]
            
        saved_paths = []
        for target_filename in target_filenames:
            time.sleep(random.uniform(2, 4)) # 下載前延遲
            print(f"找到 {ticker} {year} {doc_type} 的伺服器檔名: {target_filename}，準備下載...")
            
            # 步驟 2: 點擊下載，取得真實 PDF 轉址
            payload_download = {
                'step': '9', 
                'kind': mtype, 
                'co_id': str(ticker), 
                'filename': target_filename
            }
            res_jump = session.post(search_url, data=payload_download, verify=False, timeout=15)
            soup_jump = BeautifulSoup(res_jump.content, 'html.parser', from_encoding='big5')
            
            # 解析轉址中真實下載路徑 <a href="/pdf/...">
            pdf_a = soup_jump.find('a', href=re.compile(r'/pdf/'))
            if not pdf_a:
                print(f"解析 {target_filename} PDF 轉址失敗, html: {soup_jump.text[:100]}")
                continue
                
            real_pdf_url = "https://doc.twse.com.tw" + pdf_a['href']
            
            # 步驟 3: 實際下載 PDF 寫入硬碟
            os.makedirs(save_dir, exist_ok=True)
            # 解析檔名來分辨月份/季度與報表種類
            # 例如 202404_2330_AI1.pdf -> 202404，種類為 AI1
            month_match = re.search(r'^(\d+)_', target_filename)
            period_str = f"_{month_match.group(1)}" if month_match else ""
            
            type_code = ""
            if "AI1" in target_filename:
                type_code = "_合併"
            elif "AI2" in target_filename or "AI3" in target_filename:
                type_code = "_個體"
                
            readable_filename = f"{ticker}_{year}{period_str}_{doc_type}{type_code}.pdf"
            save_path = os.path.join(save_dir, readable_filename)
            
            print(f"正在從 {real_pdf_url} 下載...")
            res_pdf = session.get(real_pdf_url, verify=False, timeout=30)
            res_pdf.raise_for_status()
            
            with open(save_path, "wb") as f:
                f.write(res_pdf.content)
                
            print(f"已成功儲存 {readable_filename}\n")
            saved_paths.append(save_path)
            
        return saved_paths if download_all else (saved_paths[0] if saved_paths else None)

    except Exception as e:
        print(f"Error downloading {doc_type} for {ticker} year {year}: {e}")
        return [] if download_all else None

if __name__ == "__main__":
    # Test
    download_mops_pdf("2330", "112", "年報", "./test_pack")
    download_mops_pdf("2330", "113", "財報", "./test_pack")
    
