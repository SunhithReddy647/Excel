# ExcelFlow PowerShell Launcher
Write-Host "=======================================================================" -ForegroundColor Cyan
Write-Host "              EXCELFLOW - INTELLIGENT WORKBOOK PLATFORM" -ForegroundColor Cyan
Write-Host "=======================================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Check Python
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) {
    Write-Error "Python not found in PATH! Please install Python 3.10+."
    exit 1
}

# 2. Check Node
$nodeCmd = Get-Command node -ErrorAction SilentlyContinue
if (-not $nodeCmd) {
    Write-Error "Node.js not found in PATH! Please install Node.js 18+."
    exit 1
}

# 3. Setup Backend
Write-Host "[*] Preparing Backend..." -ForegroundColor Yellow
if (-not (Test-Path "backend\venv\Scripts\python.exe")) {
    Write-Host "[*] Creating virtual environment in backend\venv..." -ForegroundColor Gray
    python -m venv backend\venv
    Write-Host "[*] Installing backend dependencies..." -ForegroundColor Gray
    & "backend\venv\Scripts\pip.exe" install -r backend\requirements.txt
}

if (-not (Test-Path "backend\.env")) {
    if (Test-Path ".env") {
        Copy-Item ".env" "backend\.env"
        Write-Host "[*] Copied root .env to backend\.env" -ForegroundColor Green
    } else {
        Copy-Item ".env.example" "backend\.env"
        Write-Host "[!] Created default backend\.env. Please add your GEMINI_API_KEY." -ForegroundColor Yellow
    }
}

# Ensure storage directories
@("storage", "storage\uploads", "storage\generated", "storage\temp") | ForEach-Object {
    if (-not (Test-Path $_)) { New-Item -ItemType Directory -Path $_ -Force | Out-Null }
}

# 4. Setup Frontend
Write-Host "[*] Preparing Frontend..." -ForegroundColor Yellow
if (-not (Test-Path "frontend\node_modules")) {
    Write-Host "[*] Installing frontend npm packages..." -ForegroundColor Gray
    Push-Location frontend
    npm install
    Pop-Location
}

if (-not (Test-Path "frontend\.env.local")) {
    "NEXT_PUBLIC_API_URL=http://localhost:8000" | Out-File -FilePath "frontend\.env.local" -Encoding utf8
}

Write-Host ""
Write-Host "=======================================================================" -ForegroundColor Green
Write-Host "  STARTING SERVICES:" -ForegroundColor Green
Write-Host "  - Backend API:  http://localhost:8000 (FastAPI + SQLite + Native Pivot Engine)" -ForegroundColor White
Write-Host "  - Frontend App: http://localhost:3000 (Next.js Dashboard)" -ForegroundColor White
Write-Host "=======================================================================" -ForegroundColor Green
Write-Host ""

# 5. Launch Backend in new window
Start-Process -FilePath "cmd.exe" -ArgumentList '/k', 'cd backend && venv\Scripts\python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload'

# 6. Launch Frontend in new window
Start-Process -FilePath "cmd.exe" -ArgumentList '/k', 'cd frontend && npm run dev'

# 7. Open Browser
Start-Sleep -Seconds 3
Write-Host "[*] Opening ExcelFlow in default browser..." -ForegroundColor Cyan
Start-Process "http://localhost:3000"

Write-Host "[SUCCESS] ExcelFlow is running!" -ForegroundColor Green
Write-Host "[*] To stop services, run .\stop.bat or close the process windows." -ForegroundColor Gray
