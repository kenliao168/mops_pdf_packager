@echo off
chcp 65001 >nul 2>&1
title MOPS PDF Packager - 全部報告下載

echo ============================================
echo   MOPS PDF Packager (含 ESG 永續報告書)
echo ============================================
echo.

set /p TICKER="請輸入股票代碼 (例如 2330): "

if "%TICKER%"=="" (
    echo [錯誤] 請輸入股票代碼!
    pause
    exit /b 1
)

echo.
echo 正在下載 %TICKER% 的所有報告...
echo (包含: 年報、財報、關係企業三書表、法說會簡報、ESG永續報告書)
echo.

REM 切到 mops_pdf_packager 專案目錄 (bat 可能放在桌面，所以不能用 %~dp0)
cd /d "%~dp0"
if not exist "main.py" (
    REM 如果當前目錄沒有 main.py，嘗試找到正確路徑
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
    REM 最後用 OneDrive 特殊路徑
    for /d %%O in ("%USERPROFILE%\OneDrive*") do (
        if exist "%%O\下載報告\mops_pdf_packager\main.py" (
            cd /d "%%O\下載報告\mops_pdf_packager"
            goto :found
        )
    )
    echo [錯誤] 找不到 main.py! 請確認 mops_pdf_packager 資料夾位置。
    pause
    exit /b 1
)

:found
echo 專案路徑: %CD%
echo.
python main.py %TICKER%

echo.
echo ============================================
echo   下載完成! 請查看桌面資料夾。
echo ============================================
pause
