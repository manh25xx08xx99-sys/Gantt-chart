@echo off
rem Cai dat Add-in "工程表ツール" cho Excel tren may nay.
rem Khong can Node.js, khong can chay server (taskpane duoc host san tren GitHub Pages).
rem Ngoai viec dang ky Add-in, file nay con bat che do TU DONG CAP NHAT manifest.xml
rem tu GitHub (moi ngay 9:00 va moi lan dang nhap Windows).
rem Double-click de chay.
setlocal
set "DIR=%~dp0"
set "MANIFEST=%DIR%manifest.xml"
set "UPDATER=%DIR%auto-update.ps1"
set "RAW=https://raw.githubusercontent.com/manh25xx08xx99-sys/Gantt-chart/main"

rem (1) Neu chua co auto-update.ps1 thi tai tu GitHub (de chi can install.bat cung cai duoc)
if not exist "%UPDATER%" (
  echo Dang tai auto-update.ps1 tu GitHub...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%RAW%/auto-update.ps1' -OutFile '%UPDATER%' -UseBasicParsing" 2>nul
)

rem (2) auto-update.ps1 lo luon: tai manifest.xml moi nhat, dang ky Add-in, bat tu dong cap nhat
if exist "%UPDATER%" (
  echo Dang cai dat va bat tu dong cap nhat...
  powershell -NoProfile -ExecutionPolicy Bypass -File "%UPDATER%" -Setup
  if not errorlevel 1 goto done
  echo [Chu y] Khong bat duoc tu dong cap nhat. Se chi dang ky Add-in.
)

rem (3) Du phong: khong co mang / PowerShell bi chan -> chi dang ky manifest.xml co san
if not exist "%MANIFEST%" (
  echo [LOI] Khong tim thay manifest.xml va cung khong tai duoc tu GitHub.
  echo Hay de manifest.xml cung thu muc voi install.bat nay, roi chay lai.
  pause
  exit /b 1
)
echo Dang dang ky Add-in vao Excel...
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer" /v "dfc3fd23-ff19-4d33-b312-a15d117dd27d" /t REG_SZ /d "%MANIFEST%" /f
if errorlevel 1 (
  echo [LOI] Dang ky khong thanh cong.
  pause
  exit /b 1
)

:done
echo.
echo ================================================================
echo  Hoan tat! Mo Excel, vao tab Home, tim nut "工程表ツール".
echo  (Neu chua thay, dong het Excel roi mo lai.)
echo ================================================================
pause
