@echo off
title Client Background Investigator - Launcher
cd /d "%~dp0"

set "MODE=prod"
set "ACTION=start"

REM Parse optional args: dev / fast / restart (any order)
for %%a in (%*) do (
    if /i "%%a"=="dev"     set "MODE=dev"
    if /i "%%a"=="fast"    set "ACTION=fast"
    if /i "%%a"=="restart" set "ACTION=restart"
)

echo ==============================================
echo   Client Background Investigator - Launcher
echo ==============================================
echo.

REM ----- 0. location check -----
if not exist "package.json" (
    echo [ERROR] package.json not found.
    echo         Please put this script in the project root folder.
    echo.
    pause
    exit /b 1
)

REM ----- 1. Node.js check -----
where node >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Node.js not found. Please install Node.js 18 or newer.
    echo         Download: https://nodejs.org/
    echo.
    pause
    exit /b 1
)
for /f "delims=" %%v in ('node --version') do echo [OK] Node.js %%v

REM ----- 2. dependencies check -----
if not exist "node_modules" (
    echo [..] Installing dependencies, please wait ...
    echo.
    call npm.cmd install
    if errorlevel 1 (
        echo.
        echo [ERROR] npm install failed. Please check your network.
        echo.
        pause
        exit /b 1
    )
    echo.
)
echo [OK] Dependencies ready

REM ----- 3. restart mode: stop existing service on port 3000 -----
if /i "%ACTION%"=="restart" (
    echo [..] Restart mode: stopping existing service on port 3000 ...
    for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":3000" ^| findstr "LISTENING"') do taskkill /F /PID %%p >nul 2>nul
    ping -n 3 127.0.0.1 >nul
)

REM ----- 4. build (production mode rebuilds by default) -----
if /i not "%MODE%"=="dev" if /i not "%ACTION%"=="fast" (
    echo [..] Building latest production files, about 10 seconds ...
    echo.
    call npm.cmd run build
    if errorlevel 1 (
        echo.
        echo [ERROR] Build failed. See messages above.
        echo.
        pause
        exit /b 1
    )
    echo.
    echo [OK] Build finished
)

REM ----- 5. port check -----
netstat -ano | findstr ":3000" | findstr "LISTENING" >nul 2>nul
if not errorlevel 1 (
    echo [INFO] A service is already running on port 3000.
    echo        Open your browser at: http://localhost:3000
    echo.
    pause
    exit /b 0
)

REM ----- 6. launch -----
echo.
if /i "%MODE%"=="dev" (
    echo [OK] Starting DEV mode - hot reload.
) else (
    echo [OK] Starting PRODUCTION mode.
)
echo [WEB]  http://localhost:3000
echo [STOP] Close this window or press Ctrl+C to stop.
echo ==============================================
echo.

REM Open browser after a short delay (wait for the server)
start "" /min cmd /c "ping -n 4 127.0.0.1 >nul && start http://localhost:3000"

if /i "%MODE%"=="dev" (
    call npm.cmd run dev
) else (
    call npm.cmd start
)
set "SERVICE_EXIT_CODE=%errorlevel%"

echo.
if not "%SERVICE_EXIT_CODE%"=="0" (
    echo [ERROR] Service exited unexpectedly with code %SERVICE_EXIT_CODE%.
) else (
    echo Service stopped.
)
pause
exit /b %SERVICE_EXIT_CODE%
