@echo off
chcp 65001 >nul

REM ============================================
REM  init.bat : ตั้ง git ครั้งแรก + push ขึ้น GitHub
REM  แก้แค่บรรทัด REPO ด้านล่างให้ตรง repo ของพี่
REM ============================================
set REPO=https://github.com/nont-naraphat/it-daily.git

git init
git branch -M main
git add .
git commit -m "first commit"
git remote add origin %REPO%
git push -u origin main

echo.
echo [done] ขึ้น git แล้ว  ครั้งต่อไปใช้ push.bat พอ
pause
