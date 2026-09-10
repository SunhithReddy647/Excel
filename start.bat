@echo off
title ExcelFlow Launcher
color 0B
cls

echo =======================================================================
echo               EXCELFLOW - INTELLIGENT WORKBOOK PLATFORM
echo =======================================================================
echo.

:: 1. Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not found in PATH! Please install Python 3.10+.
    pause
    exit /b 1
)

:: 2. Check Node
node --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Node.js is not found in PATH! Please install Node.js 18+.
    pause
    exit /b 1
)

:: 3. Setup Backend Environment
echo [*] Preparing Backend...
if not exist "backend\venv\Scripts\python.exe" (
    echo [*] Creating virtual environment in backend\venv...
    python -m venv backend\venv
    echo [*] Installing backend dependencies...
    backend\venv\Scripts\pip install -r backend\requirements.txt
)

:: Copy .env if backend/.env is missing
if not exist "backend\.env" (
    if exist ".env" (
        copy /y ".env" "backend\.env" >nul
        echo [*] Copied root .env to backend\.env
    ) else (
        copy /y ".env.example" "backend\.env" >nul
        echo [!] Created default backend\.env. Please add your GEMINI_API_KEY if needed.
    )
)

:: Ensure storage directory exists
if not exist "storage" mkdir storage
if not exist "storage\uploads" mkdir storage\uploads
if not exist "storage\generated" mkdir storage\generated
if not exist "storage\temp" mkdir storage\temp

:: 4. Setup Frontend Environment
echo [*] Preparing Frontend...
if not exist "frontend\node_modules" (
    echo [*] Installing frontend npm packages (first time only)...
    cd frontend && npm install && cd ..
)

:: Copy frontend .env.local if missing
if not exist "frontend\.env.local" (
    echo NEXT_PUBLIC_API_URL=http://localhost:8000> "frontend\.env.local"
)

echo.
echo =======================================================================
echo   STARTING SERVICES:
echo   - Backend API:  http://localhost:8000 (FastAPI + SQLite + Native Pivot Engine)
echo   - Frontend App: http://localhost:3000 (Next.js Dashboard)
echo =======================================================================
echo.

:: 5. Launch Backend in new window
start "ExcelFlow Backend (Port 8000)" cmd /k "cd backend && venv\Scripts\python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload"

:: 6. Launch Frontend in new window
start "ExcelFlow Frontend (Port 3000)" cmd /k "cd frontend && npm run dev"

:: 7. Wait 3 seconds and open browser
timeout /t 3 /nobreak >nul
echo [*] Opening ExcelFlow in your browser...
start http://localhost:3000

echo.
echo [*] All services are running!
echo [*] To stop ExcelFlow, run stop.bat or close the opened terminal windows.
echo.
pause
