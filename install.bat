@echo off
rem File nay luu dang UTF-8 (khong BOM) va co chua tieng Nhat, nen doi console sang
rem UTF-8 truoc khi in ra. Neu khong, chu tieng Nhat se hien thanh ky tu la.
chcp 65001 >nul
rem Cai dat Add-in "工程表ツール" cho Excel tren may nay.
rem Khong can Node.js, khong can chay server (taskpane duoc host san tren GitHub Pages).
rem Ngoai viec dang ky Add-in, file nay con bat che do TU DONG CAP NHAT manifest.xml
rem tu GitHub (moi ngay 9:00 va moi lan dang nhap Windows).
rem
rem manifest.xml va auto-update.ps1 duoc dat vao thu muc co dinh %LOCALAPPDATA%\GanttAddin,
rem khong phu thuoc vao noi de file install.bat nay. Vi vay cai dat xong thi install.bat
rem tu xoa chinh no (nguoi dung khong can giu lai file nao).
rem Double-click de chay.
setlocal
set "SRC=%~dp0"
set "DIR=%LOCALAPPDATA%\GanttAddin\"
set "MANIFEST=%DIR%manifest.xml"
set "UPDATER=%DIR%auto-update.ps1"
set "RAW=https://raw.githubusercontent.com/manh25xx08xx99-sys/Gantt-chart/main"

if not exist "%DIR%" mkdir "%DIR%"
if not exist "%DIR%" (
  echo [LOI] Khong tao duoc thu muc cai dat: %DIR%
  pause
  exit /b 1
)

rem (1) Luon tai ban moi nhat cua auto-update.ps1 tu GitHub, ghi de len ban cu.
rem Tai vao file tam roi moi doi ten, de neu tai loi thi ban cu van con dung duoc.
echo Dang tai auto-update.ps1 tu GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%RAW%/auto-update.ps1' -OutFile '%UPDATER%.new' -UseBasicParsing" 2>nul
if exist "%UPDATER%.new" move /y "%UPDATER%.new" "%UPDATER%" >nul
rem Khong tai duoc va chua co ban nao: dung ban de cung thu muc voi install.bat (neu co)
if not exist "%UPDATER%" if exist "%SRC%auto-update.ps1" copy /y "%SRC%auto-update.ps1" "%UPDATER%" >nul

rem (2) auto-update.ps1 lo luon: tai manifest.xml moi nhat, dang ky Add-in, bat tu dong cap nhat
if exist "%UPDATER%" (
  echo Dang cai dat va bat tu dong cap nhat...
  powershell -NoProfile -ExecutionPolicy Bypass -File "%UPDATER%" -Setup
  if errorlevel 1 echo [Chu y] Khong bat duoc tu dong cap nhat. Se chi dang ky Add-in.
)

rem (3) Dam bao manifest.xml co trong thu muc cai dat va da dang ky vao Excel.
rem Khi co mang thi buoc (2) da lam xong; buoc nay danh cho truong hop khong co mang
rem (dung manifest.xml de cung thu muc voi install.bat).
if not exist "%MANIFEST%" if exist "%SRC%manifest.xml" copy /y "%SRC%manifest.xml" "%MANIFEST%" >nul
if not exist "%MANIFEST%" (
  echo [LOI] Khong tim thay manifest.xml va cung khong tai duoc tu GitHub.
  echo Hay de manifest.xml cung thu muc voi install.bat nay, roi chay lai.
  pause
  exit /b 1
)
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer" /v "dfc3fd23-ff19-4d33-b312-a15d117dd27d" /t REG_SZ /d "%MANIFEST%" /f >nul
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

rem (4) Cai dat thanh cong thi tu xoa install.bat.
rem Khong xoa khi dang chay trong thu muc ma nguon (co .git), de khong mat file trong kho.
if exist "%SRC%.git\" (
  pause
  exit /b 0
)
echo.
echo File install.bat nay se tu xoa sau khi ban nhan phim bat ky.
pause
rem "(goto)" ket thuc file bat truoc, roi lenh del moi chay, nen xoa duoc chinh no ma khong bao loi
(goto) 2>nul & del /f /q "%~f0"
