@echo off
setlocal
title Client Background Investigator - Portable
cd /d "%~dp0"

if not exist "runtime\node.exe" goto :missing_runtime
if not exist "app\dist\server.cjs" goto :missing_application
if not exist "app\dist\index.html" goto :missing_application

netstat -ano | findstr ":3000" | findstr "LISTENING" >nul 2>nul
if not errorlevel 1 goto :already_running

set "NODE_ENV=production"
cd /d "%~dp0app"

echo ==============================================
echo   Client Background Investigator - Portable
echo ==============================================
echo [OK] Embedded Node.js runtime is ready.
echo [WEB]  http://localhost:3000
echo [STOP] Close this window or press Ctrl+C to stop.
echo ==============================================
echo.

start "" /min cmd /c "ping -n 4 127.0.0.1 >nul && start http://localhost:3000"
"%~dp0runtime\node.exe" "dist\server.cjs"
set "SERVICE_EXIT_CODE=%errorlevel%"

echo.
if not "%SERVICE_EXIT_CODE%"=="0" (
  echo [ERROR] Service exited unexpectedly with code %SERVICE_EXIT_CODE%.
) else (
  echo Service stopped.
)
pause
exit /b %SERVICE_EXIT_CODE%

:already_running
echo [INFO] Port 3000 is already in use.
echo        Opening the existing service: http://localhost:3000
start "" http://localhost:3000
echo.
pause
exit /b 0

:missing_runtime
echo [ERROR] Embedded Node.js runtime is missing: runtime\node.exe
echo         Please extract the complete ZIP package again.
echo.
pause
exit /b 1

:missing_application
echo [ERROR] Portable application files are incomplete.
echo         Please extract the complete ZIP package again.
echo.
pause
exit /b 1
