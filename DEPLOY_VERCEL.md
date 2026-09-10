# 🚀 Deploying ExcelFlow to Vercel

ExcelFlow is architected with a decoupled **Next.js 16 (React 19) Frontend** and a **FastAPI + SQLite + Native Pivot Engine Backend**.

---

## ⚡ Quick Deploy to Vercel (Frontend)

### Option A: Via GitHub (Recommended)

1. **Push your repository to GitHub**:
   ```bash
   git add .
   git commit -m "feat: complete ExcelFlow platform with native PivotTables and SQLite"
   git push origin main
   ```

2. **Import into Vercel**:
   - Go to [vercel.com/new](https://vercel.com/new)
   - Select your GitHub repository (`excel-flow`)
   - **Framework Preset**: `Next.js` (automatically detected)
   - **Root Directory**: Select `frontend` (or leave default `/` — the root `vercel.json` will automatically build `frontend`)

3. **Configure Environment Variables in Vercel**:
   | Variable | Value | Description |
   |---|---|---|
   | `NEXT_PUBLIC_API_URL` | `https://your-backend-url.com` | URL of your deployed backend |
   | `BACKEND_URL` | `https://your-backend-url.com` | Fallback proxy URL |

4. Click **Deploy**! ✨

---

### Option B: Via Vercel CLI

```bash
# Install Vercel CLI if needed
npm install -g vercel

# Deploy directly from the frontend directory
cd frontend
vercel
```

---

## 🖥️ Deploying the Backend

The backend is built with **FastAPI**, **SQLite**, and the **Gemini AI API**.

### Easy Backend Hosting Options:

#### 1. Render / Railway / Fly.io (Docker or Python)
The repository includes a ready-to-use [`docker-compose.yml`](file:///c:/Users/sunhi/Desktop/Excellence/excel-flow/docker-compose.yml) and [`backend/Dockerfile`](file:///c:/Users/sunhi/Desktop/Excellence/excel-flow/backend/Dockerfile).

**Required Backend Environment Variables**:
```env
GEMINI_API_KEY=your_gemini_api_key_here
DATABASE_URL=sqlite+aiosqlite:///./storage/excelflow.db
STORAGE_PATH=./storage
FRONTEND_URL=https://your-app.vercel.app
BACKEND_URL=https://your-backend-url.com
```

#### 2. Localhost with Tunnel (e.g. ngrok / Cloudflare Tunnel)
If you want your Vercel frontend to leverage the **Native Microsoft Excel COM PivotTable & Slicer Engine** running on your local Windows PC:
```bash
# Expose your local port 8000
ngrok http 8000
```
Set `NEXT_PUBLIC_API_URL=https://xxxx.ngrok-free.app` in your Vercel project settings!

---

## 🎯 Dual-Engine Architecture

- **Windows Desktop / Local Host**: Detects Microsoft Office and uses `pywin32` COM automation to generate **100% genuine Microsoft Excel PivotTables (`PivotCache`, `PivotTable`)**, interactive Slicers (`SlicerCaches`), and linked PivotCharts.
- **Headless Cloud / Docker**: In Linux / headless environments where Excel is not installed, the platform automatically falls back to **high-fidelity formula-driven summary tables (`SUMIFS`, `COUNTIFS`)** and charts. The generation never crashes in any environment!
