@echo off
chcp 65001 >nul
setlocal

if "%~1"=="" (
    echo.
    echo Office HQ PDF Converter v2
    echo.
    echo Drag one or more Office files onto this CMD file.
    echo Supported: DOCX DOC XLSX XLS PPTX PPT
    echo.
    echo Preferred engine: Adobe Acrobat PDFMaker
    echo Fallback: Microsoft Office native PDF export
    echo Logs: Office_HQ_PDF_Converter\logs
    echo.
    pause
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Convert-OfficeToPDF-v2.ps1" %*
set "ERR=%ERRORLEVEL%"

echo.
if not "%ERR%"=="0" (
    echo One or more files failed. Please check the logs folder.
) else (
    echo Conversion finished. Please check the generated PDF and logs folder.
)
pause
exit /b %ERR%
