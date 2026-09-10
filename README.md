# ExcelFlow — Intelligent Excel Assignment Automation Platform

> Turn raw datasets and assignment questions into professionally formatted, human-grade Excel workbooks featuring **real formulas, authentic native PivotTables, interactive Slicers, charts, and executive dashboards**.

---

## ⚡ Super Easy Run (1-Click Local Launch)

### On Windows:
Simply double-click:
```cmd
start.bat
```
or in PowerShell:
```powershell
.\start.ps1
```
This automatically sets up Python dependencies, frontend packages, starts both services, and opens `http://localhost:3000` in your default browser!

To stop all services:
```cmd
stop.bat
```

---

## 🚀 Vercel Publish Ready

The frontend is configured for seamless zero-config deployment on **Vercel**:

- **Framework**: Next.js 16 (React 19, Turbopack, Tailwind CSS v4)
- **Monorepo / Subfolder Ready**: Both root [`vercel.json`](file:///c:/Users/sunhi/Desktop/Excellence/excel-flow/vercel.json) and [`frontend/vercel.json`](file:///c:/Users/sunhi/Desktop/Excellence/excel-flow/frontend/vercel.json) are preconfigured.
- **Security Headers & API Rewrites**: Pre-configured in [`frontend/next.config.ts`](file:///c:/Users/sunhi/Desktop/Excellence/excel-flow/frontend/next.config.ts) to proxy `/api/*` to your backend without CORS friction.
- **Full Guide**: See [`DEPLOY_VERCEL.md`](file:///c:/Users/sunhi/Desktop/Excellence/excel-flow/DEPLOY_VERCEL.md).

---

## 🏗️ Technical Architecture

| Layer | Technologies | Features |
|---|---|---|
| **Frontend** | Next.js 16, React 19, TypeScript, Tailwind CSS | Modern executive dark-mode UI, step-by-step progress, rich validation metrics & preview |
| **Backend API** | FastAPI, Uvicorn, Python 3.10+ | Async REST endpoints, background task queue, CORS support |
| **Database** | SQLite + async SQLAlchemy (`storage/excelflow.db`) | Complete persistence of datasets, assignments, generation logs, and workbooks |
| **AI Engine** | Google Gemini (`gemini-3.6-flash`) | Schema understanding, question parsing, OCR for image/PDF/DOCX, validated plan generation |
| **Pivot Engine** | `pywin32` Excel COM Automation (Windows) | **Utmost authentic native Excel PivotTables (`PivotCache`, `PivotTable`)**, interactive Slicers, and linked PivotCharts |
| **Fallback Engine** | `openpyxl`, `pandas`, `xlsxwriter` | Zero-crash fallback producing structured tables & formula summaries on headless cloud hosts |

---

## 🧪 Testing

Run the full backend test suite:
```bash
cd backend
venv\Scripts\pytest -v
```
All 13 tests verify profiler, formula engine, SQLite transactions, and end-to-end workbook generation with native PivotTables.
