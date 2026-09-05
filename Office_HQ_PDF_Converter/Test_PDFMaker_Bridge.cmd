@echo off
chcp 65001 >nul
setlocal

if "%~1"=="" (
  echo Drag a DOC or DOCX file onto this CMD file.
  pause
  exit /b 1
)

set "EXT=%~x1"
if /I not "%EXT%"==".docx" if /I not "%EXT%"==".doc" (
  echo This bridge test currently supports Word DOC/DOCX only.
  pause
  exit /b 2
)

set "OUT=%~dpn1_AdobeBridge.pdf"
echo Source: %~1
echo Output: %OUT%
echo.
cscript.exe //nologo "%~dp0PdfMaker-Bridge.vbs" "%~1" "%OUT%"
set "ERR=%ERRORLEVEL%"
echo.
if "%ERR%"=="0" (
  echo PDFMaker bridge completed successfully.
) else (
  echo PDFMaker bridge failed with exit code %ERR%.
)
pause
exit /b %ERR%
