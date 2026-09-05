@echo off
chcp 65001 >nul
setlocal

echo.
echo [RETIRED TEST]
echo Word ExportAsFixedFormat2 is not used by the production converter anymore.
echo On Word 16.0.17932 this COM call can block indefinitely in PowerShell automation.
echo.
echo Use instead:
echo   DragDrop_Office_to_PDF.cmd
echo.
echo The v3 production path uses Office only for layout, then restores original OOXML image pixels.
echo.
pause
exit /b 3
