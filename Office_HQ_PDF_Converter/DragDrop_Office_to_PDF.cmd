@echo off
chcp 65001 >nul
setlocal

if "%~1"=="" (
    echo.
    echo High Quality Office -^> PDF Converter
    echo.
    echo Drag one or more Office files onto this CMD file.
    echo Supported: DOCX DOC XLSX XLS PPTX PPT
    echo.
    pause
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Convert-OfficeToPDF.ps1" %*
set "ERR=%ERRORLEVEL%"

echo.
if not "%ERR%"=="0" (
    echo One or more files failed. Read the messages above for details.
)
pause
exit /b %ERR%
