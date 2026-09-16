@echo off
rem Go Add-in "工程表ツール" khoi Excel tren may nay (khong can Node.js).
rem Xoa ca tac vu tu dong cap nhat manifest.xml.
schtasks /delete /tn "GanttAddin_AutoUpdateManifest" /f >nul 2>&1
if errorlevel 1 (
  echo Khong co tac vu tu dong cap nhat (hoac da xoa truoc do^).
) else (
  echo Da xoa tac vu tu dong cap nhat.
)
reg delete "HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer" /v "dfc3fd23-ff19-4d33-b312-a15d117dd27d" /f
if errorlevel 1 (
  echo Khong xoa duoc dang ky Add-in (co the da go truoc do^).
) else (
  echo Da xoa dang ky. Khoi dong lai Excel de an hoan toan.
)
pause
