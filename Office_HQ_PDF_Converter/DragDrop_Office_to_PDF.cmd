@echo off
chcp 65001 >nul
setlocal

if "%~1"=="" (
    echo.
    echo Office HQ PDF Converter v3
    echo.
    echo Drag one or more Office files onto this CMD file.
    echo Supported: DOCX DOC XLSX XLS PPTX PPT
    echo.
    echo Stable HQ path for DOCX/XLSX/PPTX:
    echo   Office layout engine ^> structural PDF ^> lossless original-image rebuild
    echo.
    echo PDFMaker and Word ExportAsFixedFormat2 are NOT used.
    echo Logs: Office_HQ_PDF_Converter\logs
    echo Diagnostics: Office_HQ_PDF_Converter\diagnostics
    echo.
    pause
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Convert-OfficeToPDF-v3.ps1" %*
set "ERR=%ERRORLEVEL%"

echo.
if not "%ERR%"=="0" (
    echo One or more files failed or could only produce NATIVE_ONLY output.
    echo Please check the logs folder.
) else (
    echo Conversion finished.
    echo HQ_REBUILT means the original OOXML image pixels were restored losslessly.
)
pause
exit /b %ERR%
