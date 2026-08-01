@echo off
setlocal

set "APP_NAME=OllamaOutlookAddin"
set "DISPLAY_NAME=CoOllama"
set "CATALOG_GUID={6B7A37F4-4E07-4B67-B965-81C57E1D7F14}"
set "DEST=%LOCALAPPDATA%\%APP_NAME%"
set "START_MENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs\CoOllama"
set "STARTUP_LINK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Start CoOllama Server.lnk"
set "REGKEY=HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%"

echo Uninstalling %DISPLAY_NAME% Outlook add-in...
echo.
echo If the CoOllama server window is open, close it before continuing.
echo.

reg delete "%REGKEY%" /f >nul 2>nul

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$names=@('Ask Ollama Localhost Dev','CoOllama Localhost Dev'); foreach ($store in @('Cert:\CurrentUser\My','Cert:\CurrentUser\Root')) { Get-ChildItem $store ^| Where-Object { $names -contains $_.FriendlyName } ^| ForEach-Object { Remove-Item -LiteralPath $_.PSPath -Force } }" >nul 2>nul

if exist "%STARTUP_LINK%" del /F /Q "%STARTUP_LINK%"
if exist "%START_MENU%" rmdir /S /Q "%START_MENU%"
if exist "%DEST%" rmdir /S /Q "%DEST%"

echo %DISPLAY_NAME% files, shortcuts, startup entry, and local certificates removed.
echo.
echo If you manually added the add-in in Outlook, remove it from Outlook's My add-ins page too:
echo https://aka.ms/olksideload
echo.
pause
