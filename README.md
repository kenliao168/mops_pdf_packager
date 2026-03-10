# MOPS PDF Packager

台股公開資訊報告一鍵下載工具。輸入股票代碼，自動從公開資訊觀測站 (MOPS) 及 ESG 數位平台下載報告，方便上傳至 Google NotebookLM 進行 AI 分析。

## 下載內容

| 報告類型 | 來源 | 數量 |
|---------|------|------|
| 年報 | doc.twse.com.tw | 最近 5 年 |
| 財報 | doc.twse.com.tw | 最近年度全季 + 歷史年度 Q4 |
| 關係企業三書表 | doc.twse.com.tw | 最近 5 年 |
| 法說會簡報 | mopsov.twse.com.tw | 最近 5 場 |
| ESG 永續報告書 | esggenplus.twse.com.tw | 最近 3 年 |
| 公開說明書 | doc.twse.com.tw | 最近 1 期 |

## 使用方式

### 方式一：下載 EXE（推薦，不需安裝 Python）

1. 到 [Releases](../../releases) 頁面下載 `MOPS_PDF_Packager.exe`
2. 雙擊執行
3. 輸入股票代碼 → 點擊「開始下載」
4. 報告會自動存到桌面的 `{代碼} {公司名} NotebookLM上傳文件` 資料夾

### 方式二：用 Python 執行

```bash
# 安裝依賴
pip install -r requirements.txt

# GUI 模式
python gui.py

# CLI 模式
python main.py 2330
python main.py 2330 --year 113
```

### 方式三：用 BAT 檔

雙擊 `下載MOPS全部報告.bat`，輸入股票代碼即可。

## 自己打包 EXE

```bash
# 需要 Python 3.8+
pip install pyinstaller
# 雙擊 build_exe.bat 或手動執行:
pyinstaller --onefile --windowed --name "MOPS_PDF_Packager" --add-data "src;src" gui.py
```

打包後的 EXE 會在 `dist/` 資料夾裡。

## 專案結構

```
mops_pdf_packager/
├── main.py                # CLI 入口
├── gui.py                 # GUI 入口 (tkinter)
├── requirements.txt
├── build_exe.bat          # 一鍵打包 EXE
├── 下載MOPS全部報告.bat     # BAT 快捷方式
├── 只下載ESG永續報告書.bat   # ESG 單獨下載
└── src/
    ├── scrapers/
    │   ├── ebook_scraper.py       # 年報/財報/三書表
    │   ├── briefing_scraper.py    # 法說會簡報
    │   ├── esg_scraper.py         # ESG 永續報告書
    │   ├── prospectus_scraper.py  # 公開說明書
    │   └── mopsov_scraper.py      # Selenium 備援
    └── utils/
        └── downloader.py          # 主流程控制
```

## 注意事項

- 部分公司可能沒有 ESG 永續報告書，工具會自動跳過
- 下載速度取決於 MOPS / TWSE 伺服器回應速度，請耐心等待
- 如遇下載失敗，可能是伺服器暫時繁忙，稍後重試即可
