@echo off
setlocal

set "APP_NAME=OllamaOutlookAddin"
set "DISPLAY_NAME=CoOllama"
set "CATALOG_GUID={6B7A37F4-4E07-4B67-B965-81C57E1D7F14}"
set "SRC=%~dp0"
set "DEST=%LOCALAPPDATA%\%APP_NAME%"
set "START_MENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs\CoOllama"
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "REGKEY=HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%"

cd /d "%SRC%"

echo Installing %DISPLAY_NAME% Outlook add-in...
echo.

if not exist "%DEST%" mkdir "%DEST%"
if not exist "%DEST%\pane" mkdir "%DEST%\pane"
if not exist "%DEST%\assets" mkdir "%DEST%\assets"
if not exist "%START_MENU%" mkdir "%START_MENU%"
if not exist "%STARTUP%" mkdir "%STARTUP%"

copy /Y "%SRC%manifest.xml" "%DEST%\" >nul
copy /Y "%SRC%start-server.bat" "%DEST%\" >nul
copy /Y "%SRC%start-server.ps1" "%DEST%\" >nul
copy /Y "%SRC%README.md" "%DEST%\" >nul 2>nul
copy /Y "%SRC%ALPHA_BACKLOG.md" "%DEST%\" >nul 2>nul
xcopy /E /I /Y "%SRC%pane" "%DEST%\pane" >nul
xcopy /E /I /Y "%SRC%assets" "%DEST%\assets" >nul

reg add "%REGKEY%" /v Id /t REG_SZ /d "%CATALOG_GUID%" /f >nul
reg add "%REGKEY%" /v Url /t REG_SZ /d "%DEST%" /f >nul
reg add "%REGKEY%" /v Flags /t REG_DWORD /d 1 /f >nul

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$paths=@('%START_MENU%\Start CoOllama Server.lnk','%STARTUP%\Start CoOllama Server.lnk'); $w=New-Object -ComObject WScript.Shell; foreach ($p in $paths) { $s=$w.CreateShortcut($p); $s.TargetPath='%DEST%\start-server.bat'; $s.WorkingDirectory='%DEST%'; $s.IconLocation='%DEST%\assets\icon-80.png'; $s.Save() }" >nul 2>nul

where ollama >nul 2>nul
if errorlevel 1 (
  echo Ollama status: not found on PATH.
  echo Install Ollama from https://ollama.com before using %DISPLAY_NAME%.
) else (
  echo Ollama status: found on PATH.
)

set "OLLAMA_TAGS=%TEMP%\coollama-ollama-tags.json"
curl.exe -fsS --max-time 5 "http://127.0.0.1:11434/api/tags" -o "%OLLAMA_TAGS%" >nul 2>nul
if errorlevel 1 (
  echo Ollama status: installed, but not responding at http://127.0.0.1:11434 right now.
  echo If Ollama is already running, the add-in pane can still retry from inside Outlook.
) else (
  findstr /C:"\"name\"" "%OLLAMA_TAGS%" >nul 2>nul
  if errorlevel 1 (
    echo Ollama status: running, but no local models were found.
    echo Pull at least one model before using %DISPLAY_NAME%, for example: ollama pull llama3.2
  ) else (
    echo Ollama status: running with at least one installed model.
  )
)
if exist "%OLLAMA_TAGS%" del /F /Q "%OLLAMA_TAGS%" >nul 2>nul

echo.
echo %DISPLAY_NAME% files installed to:
echo %DEST%
echo.
echo Startup behavior:
echo - A Start Menu shortcut was created at CoOllama ^> Start CoOllama Server.
echo - A Startup shortcut was created so the CoOllama local server starts when Windows starts.
echo - You can still run the Start Menu shortcut manually if you closed the server window.
echo.
echo Next steps for New Outlook:
echo 1. Start the server now from Start Menu ^> CoOllama ^> Start CoOllama Server.
echo    Leave that server window open while using the add-in.
echo 2. Open https://aka.ms/olksideload in a browser.
echo 3. My add-ins ^> Custom Addins ^> Add a custom add-in ^> Add from File.
echo 4. Select this manifest:
echo    %DEST%\manifest.xml
echo 5. In New Outlook, open an email and look under Apps / More apps / ... for CoOllama.
echo.
echo Note: New Outlook may require reopening the add-in when switching emails.
echo.
pause


