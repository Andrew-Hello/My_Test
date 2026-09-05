@echo off
chcp 65001 >nul
setlocal

if "%~1"=="" (
  echo Drag a DOC or DOCX file onto this CMD file.
  pause
  exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-WordFixedFormat2-HQ.ps1" "%~1"
set "ERR=%ERRORLEVEL%"
echo.
pause
exit /b %ERR%
