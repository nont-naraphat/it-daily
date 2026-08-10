@echo off
chcp 65001 >nul
setlocal

REM ===== push.bat : git add -> commit -> push =====
REM ใช้:  push.bat "commit message"   (ถ้าไม่ใส่ จะใช้ timestamp)

if "%~1"=="" (
    set "MSG=update %date% %time%"
) else (
    set "MSG=%~1"
)

REM ครั้งแรก: init + ตั้ง remote (แก้ URL ให้ตรง repo ของพี่)
if not exist ".git" (
    echo [init] first time setup...
    git init
    git branch -M main
    git remote add origin https://github.com/nont-naraphat/it-daily.git
)

git add .
git commit -m "%MSG%"
git push -u origin main

echo.
echo [done] pushed: %MSG%
endlocal
pause
