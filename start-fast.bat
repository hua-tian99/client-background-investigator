@echo off
title Client Background Investigator - Fast Start
cd /d "%~dp0"

if not exist "dist\server.cjs" goto :missing_build
if not exist "dist\index.html" goto :missing_build

echo [FAST START] Using existing production files. Build skipped.
echo.
call "%~dp0start.bat" fast
exit /b %errorlevel%

:missing_build
echo [ERROR] Production files are missing.
echo         Run start.bat once to build the project first.
echo.
pause
exit /b 1
