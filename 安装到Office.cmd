@echo off
title DeepSeek Office Add-in Installer
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
if errorlevel 1 (
  echo.
  echo Install FAILED - see messages above.
  pause
  exit /b 1
)
echo.
pause
