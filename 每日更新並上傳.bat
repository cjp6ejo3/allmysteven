@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   每日更新：檢查已使用 + 新網址 + 上傳
echo ========================================
echo.
echo 1. 掃描 Yahoo 查詢結果（含今日新檔）
echo 2. 重新檢查每張券是否已使用
echo 3. 上傳到 GitHub
echo.

python extract_and_upload.py --refresh --upload

echo.
pause
