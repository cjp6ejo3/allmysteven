@echo off
chcp 65001 >nul
echo ========================================
echo   第一次上傳到 GitHub - allmysteven
echo ========================================
echo.

REM 檢查 Git 是否已安裝
git --version >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 尚未安裝 Git！
    echo 請先到 https://git-scm.com/download/win 下載並安裝
    echo 安裝完成後重新執行此腳本。
    pause
    exit /b 1
)

cd /d "%~dp0"

echo [1/7] 設定 Git 使用者資訊...
echo 請輸入你的 GitHub 註冊信箱（在 https://github.com/settings/emails 可查）
set /p GIT_EMAIL=信箱: 
if "%GIT_EMAIL%"=="" (
    echo 未輸入信箱，使用預設值（之後可再修改）
    set GIT_EMAIL=cjp6ejo3@users.noreply.github.com
)
git config user.email "%GIT_EMAIL%"
git config user.name "cjp6ejo3"
echo 已設定: %GIT_EMAIL%
echo.

echo [2/7] 初始化 Git 儲存庫...
if not exist ".git" (
    git init
    echo 完成
) else (
    echo 已是 Git 儲存庫，略過
)
echo.

echo [3/7] 設定遠端...
git remote remove origin 2>nul
git remote add origin https://github.com/cjp6ejo3/allmysteven.git
echo 完成
echo.

echo [4/7] 加入所有檔案...
git add .
echo 完成
echo.

echo [5/7] 建立提交...
git commit -m "初次上傳：電子券清單與 Telegram 獎品網址整理"
if errorlevel 1 (
    echo 沒有新變更可提交，或已是空提交
)
echo.

echo [6/7] 設定主分支...
git branch -M main
echo 完成
echo.

echo [7/7] 推送到 GitHub...
echo 若跳出登入視窗，請用 GitHub 帳號登入
echo 若要求密碼，請使用 Personal Access Token（非 GitHub 密碼）
echo.
git push -u origin main

if errorlevel 1 (
    echo.
    echo [可能失敗原因]
    echo - 未登入 GitHub：請在跳出的瀏覽器中登入
    echo - 密碼錯誤：GitHub 已不支援密碼，請用 Token
    echo   建立 Token: https://github.com/settings/tokens
    echo   勾選 repo 權限後產生，在要求密碼時貼上
) else (
    echo.
    echo ========================================
    echo   上傳成功！
    echo   網址: https://cjp6ejo3.github.io/allmysteven/
    echo ========================================
)

echo.
pause
