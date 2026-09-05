@echo off
chcp 65001 >nul
setlocal

echo.
echo Collecting local environment information...
echo The report intentionally avoids usernames, computer names and secrets.
echo.

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Collect-Local-Environment.ps1"
set "ERR=%ERRORLEVEL%"

echo.
if not "%ERR%"=="0" (
  echo Environment collection returned an error. Please keep the console text for diagnosis.
) else (
  echo Done. Check Local_Environment_Collector\reports\
  echo Commit and Push the generated JSON/TXT files, then ask ChatGPT to read them.
)
echo.
pause
exit /b %ERR%
