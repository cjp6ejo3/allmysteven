@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   每日更新：可兌換券全部重查 + 上傳
echo ========================================
echo.
echo 1. 彙整全部 Yahoo 查詢結果
echo 2. 每次都把尚未標已兌換的券全部連網重查
echo 3. 查到已使用就從開啟清單拿掉，不上傳
echo 4. 已標已兌換的不再重抓
echo 5. 有變更才上傳 GitHub
echo.
echo 若要全量重抓：python extract_and_upload.py --refresh --all
echo 若要含已兌換上傳：python extract_and_upload.py --refresh --days 10 --include-used
echo.

python extract_and_upload.py --refresh --days 10

echo.
pause
