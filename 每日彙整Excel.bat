@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   每日更新：檢查已使用 + 彙整成 Excel
echo ========================================
echo.
echo 1. 掃描 Yahoo 查詢結果（含今日新檔）
echo 2. 重新檢查每張券是否已使用
echo 3. 匯出 Excel
echo.

python extract_to_excel.py --refresh

echo.
echo 💡 Excel 檔案已產生在 github 資料夾內。
echo.
pause
