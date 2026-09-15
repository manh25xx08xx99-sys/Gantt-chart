@echo off
rem Cai dat Add-in "工程表ツール" cho Excel tren may nay.
rem Chi can file nay + manifest.xml nam CUNG THU MUC voi nhau - khong can Node.js,
rem khong can chay server (taskpane duoc host san tren GitHub Pages).
rem Double-click de chay.
setlocal
set "MANIFEST=%~dp0manifest.xml"

if not exist "%MANIFEST%" (
  echo [LOI] Khong tim thay manifest.xml cung thu muc voi install.bat nay.
  echo Hay dam bao 2 file nam chung 1 cho.
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

echo.
echo ================================================================
echo  Hoan tat! Mo Excel, vao tab Home, tim nut "工程表ツール".
echo  (Neu chua thay, dong het Excel roi mo lai.)
echo ================================================================
pause
