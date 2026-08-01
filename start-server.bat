@echo off
setlocal

set "APP_DIR=%LOCALAPPDATA%\OllamaOutlookAddin"
if not exist "%APP_DIR%\manifest.xml" (
  set "APP_DIR=%~dp0"
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%APP_DIR%\start-server.ps1"
if errorlevel 1 pause
