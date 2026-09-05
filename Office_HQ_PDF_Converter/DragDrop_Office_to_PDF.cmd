@echo off
chcp 65001 >nul
setlocal

if "%~1"=="" (
    echo.
    echo Office HQ PDF Converter
    echo.
    echo Drag one or more Office files onto this CMD file.
    echo Supported: DOCX DOC XLSX XLS PPTX PPT
    echo.
    echo Production pipeline:
    echo Office layout engine -^> temporary structural PDF -^> HQ image restoration
    echo.
    echo HQ restoration dependencies: Python 3 + PyMuPDF + Pillow
    echo ImageHash / Acrobat PDFMaker / ExportAsFixedFormat2 are NOT required.
    echo.
    pause
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Convert-OfficeToPDF.ps1" %*
set "ERR=%ERRORLEVEL%"

echo.
if not "%ERR%"=="0" (
    echo One or more files failed or were degraded to NATIVE_ONLY.
    echo Check logs\ and diagnostics\ for details.
    echo If the cause is environment-related, run ..\Local_Environment_Collector\Collect_Local_Environment.cmd
) else (
    echo Conversion finished successfully.
    echo HQ outputs are named *_HQ.pdf.
)
echo.
pause
exit /b %ERR%
