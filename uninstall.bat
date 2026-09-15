@echo off
rem Go Add-in "工程表ツール" khoi Excel tren may nay (khong can Node.js).
reg delete "HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer" /v "dfc3fd23-ff19-4d33-b312-a15d117dd27d" /f
if errorlevel 1 (
  echo Khong xoa duoc (co the da go truoc do^).
) else (
  echo Da xoa dang ky. Khoi dong lai Excel de an hoan toan.
)
pause
