@echo off
chcp 65001 >nul 2>&1
title 測試公開說明書下載
cd /d "%~dp0"

echo ============================================
echo   測試公開說明書下載功能
echo ============================================
echo.

set /p TICKER=請輸入股票代碼 (例如 2330):

echo.
echo 正在測試下載 %TICKER% 的公開說明書...
echo.

python -c "import sys; sys.path.insert(0, '.'); from src.scrapers.prospectus_scraper import download_prospectus; paths = download_prospectus('%TICKER%', './test_prospectus', max_reports=1); print(); print('=== 結果 ==='); print('下載路徑:', paths if paths else '無'); print('成功!' if paths else '失敗!')"

echo.
echo ============================================
if exist "test_prospectus\*.pdf" (
    echo 成功! PDF 已下載到 test_prospectus 資料夾
    dir test_prospectus\*.pdf
) else (
    echo 失敗! 沒有下載到任何 PDF
)
echo ============================================
echo.
pause
