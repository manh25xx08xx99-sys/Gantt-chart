@echo off
rem Start the static server for the Gantt Excel add-in (keeps this window open)
cd /d "%~dp0"
node server.js
pause
