@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js was not found. Please install Node.js first.
  pause
  exit /b 1
)

if not exist node_modules (
  echo Installing dependencies...
  call npm install
  if errorlevel 1 (
    echo Dependency installation failed.
    pause
    exit /b 1
  )
)

start "BioMechanism AI Server" /D "%~dp0" cmd /k npm run dev -- --host 127.0.0.1 --port 5173
timeout /t 3 /nobreak >nul
start "" "http://127.0.0.1:5173/"
endlocal
