@echo off
chcp 65001 >nul 2>&1
title ESG 永續報告書 獨立下載工具

echo ============================================
echo   ESG 永續報告書 獨立下載工具
echo   (esggenplus.twse.com.tw)
echo ============================================
echo.

set /p TICKER="請輸入股票代碼 (例如 2330): "

if "%TICKER%"=="" (
    echo [錯誤] 請輸入股票代碼!
    pause
    exit /b 1
)

echo.
echo 正在從 ESG 數位平台下載 %TICKER% 的永續報告書...
echo.

REM 切到 mops_pdf_packager 專案目錄
cd /d "%~dp0"
if not exist "main.py" (
    for %%D in (
        "%USERPROFILE%\OneDrive\下載報告\mops_pdf_packager"
        "%USERPROFILE%\OneDrive - 博信資產管理股份有限公司\下載報告\mops_pdf_packager"
        "%USERPROFILE%\下載報告\mops_pdf_packager"
        "%USERPROFILE%\Desktop\mops_pdf_packager"
        "%USERPROFILE%\桌面\mops_pdf_packager"
    ) do (
        if exist "%%~D\main.py" (
            cd /d "%%~D"
            goto :found
        )
    )
    for /d %%O in ("%USERPROFILE%\OneDrive*") do (
        if exist "%%O\下載報告\mops_pdf_packager\main.py" (
            cd /d "%%O\下載報告\mops_pdf_packager"
            goto :found
        )
    )
    echo [錯誤] 找不到專案資料夾! 請確認 mops_pdf_packager 位置。
    pause
    exit /b 1
)

:found
echo 專案路徑: %CD%
echo.

REM 使用桌面路徑作為儲存位置
python -c "import sys; sys.path.insert(0,'.'); from src.scrapers.esg_scraper import download_esg_report; download_esg_report('%TICKER%', './%TICKER%_ESG_Reports')"

echo.
echo ============================================
echo   完成! 請查看 %TICKER%_ESG_Reports 資料夾。
echo ============================================
pause
