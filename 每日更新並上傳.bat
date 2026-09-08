@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   每日更新：近10天檢查 + 新網址 + 上傳
echo ========================================
echo.
echo 1. 彙整全部 Yahoo 查詢結果
echo 2. 只重新檢查「最近 10 天」券是否已使用
echo 3. 到期卻未標已兌換 → 強制重查（修未換過誤判）
echo 4. 已兌換預設不上傳頁面（僅可兌換）
echo 5. 有變更才上傳 GitHub
echo.
echo 若要全量重抓：python extract_and_upload.py --refresh --all
echo 若要含已兌換上傳：python extract_and_upload.py --refresh --days 10 --include-used
echo.

python extract_and_upload.py --refresh --days 10

echo.
pause
