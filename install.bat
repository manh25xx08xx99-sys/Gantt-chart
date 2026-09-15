@echo off
rem Cai dat Add-in "Cong trinh do" (工程表ツール) cho Excel tren may nay.
rem Taskpane duoc host san tren GitHub Pages - khong can chay server cuc bo.
rem Double-click de chay - khong can mo terminal thu cong.
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo [LOI] Chua tim thay Node.js tren may nay.
  echo Vui long cai Node.js tai https://nodejs.org/ ^(ban LTS^) roi chay lai file nay.
  pause
  exit /b 1
)

echo Dang dang ky Add-in vao Excel...
node install.cjs
if errorlevel 1 (
  pause
  exit /b 1
)

echo.
echo ================================================================
echo  Hoan tat! Mo Excel, vao tab Home, tim nut "工程表ツール".
echo  (Neu chua thay, dong het Excel roi mo lai.)
echo ================================================================
pause
