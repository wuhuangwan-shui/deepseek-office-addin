@echo off
chcp 65001 >nul
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Node.js not found. Please install Node.js first.
  pause
  exit /b 1
)
echo ==============================================
echo   DeepSeek Office Add-in - Local Server
echo   URL: https://localhost:3000
echo   Keep this window open while using Word/Excel.
echo   (First start prepares the local HTTPS cert.)
echo ==============================================
node server.js
pause
