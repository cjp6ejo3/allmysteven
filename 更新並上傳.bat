@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   更新電子券清單並上傳到 GitHub
echo ========================================
echo.

python extract_and_upload.py --upload

echo.
pause
