@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   每日更新：近10天檢查 + 新網址 + 上傳
echo ========================================
echo.
echo 1. 彙整全部 Yahoo 查詢結果（HTML 完整保留）
echo 2. 只重新檢查「最近 10 天」券是否已使用
echo 3. 新網址才連網；舊快取不重抓
echo 4. 有變更才上傳 GitHub（沒新的就略過）
echo.
echo 若要全量重抓：python extract_and_upload.py --refresh --all
echo.

python extract_and_upload.py --refresh --days 10

echo.
pause
