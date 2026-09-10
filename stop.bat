@echo off
title Stop ExcelFlow
color 0C
cd /d "%~dp0"
cls

echo =======================================================================
echo                     STOPPING EXCELFLOW SERVICES
echo =======================================================================
echo.

echo [*] Terminating processes on port 8000 (Backend)...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8000" ^| findstr "LISTENING"') do (
    taskkill /f /pid %%a >nul 2>&1
    echo [*] Stopped Backend process PID %%a
)

echo [*] Terminating processes on port 3000 (Frontend)...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":3000" ^| findstr "LISTENING"') do (
    taskkill /f /pid %%a >nul 2>&1
    echo [*] Stopped Frontend process PID %%a
)

echo.
echo [SUCCESS] All ExcelFlow services have been stopped.
timeout /t 2 >nul
exit /b 0
