"""
app.py

FastAPI Web Application Backend for SBL Trend Analysis.
Hosts the web dashboard and REST APIs for file processing, chart visualization,
and multi-worksheet Excel report generation.

Complies with the Windows VM Web Application Deployment Handbook:
  - Listens on http://127.0.0.1:8000 (proxied via IIS ARR + URL Rewrite)
  - Managed as a Windows Service via NSSM
  - Provides /health, /api/analyze, /api/download, and web dashboard
"""

import io
import os
import time
import uuid
from datetime import datetime
from typing import Any, Dict, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles

from sbl_engine import process_sbl_data

app = FastAPI(
    title="SBL Trend Analysis Web Application",
    description="Statistical Bin Limit (SBL) Trigger Trend & Statistical Analysis Dashboard",
    version="1.0.0",
)

# In-memory storage for generated Excel reports (job_id -> { "bytes": bytes, "filename": str, "timestamp": float })
REPORT_CACHE: Dict[str, Dict[str, Any]] = {}
CACHE_TTL_SECONDS = 3600  # 1 hour cache lifetime


def cleanup_cache():
    """Remove expired cached reports to prevent memory leak."""
    now = time.time()
    expired = [k for k, v in REPORT_CACHE.items() if now - v["timestamp"] > CACHE_TTL_SECONDS]
    for k in expired:
        REPORT_CACHE.pop(k, None)


# ---------------------------------------------------------------------
# API ENDPOINTS
# ---------------------------------------------------------------------
@app.get("/health", tags=["Health"])
def health_check():
    """
    Service healthcheck endpoint used by IIS ARR, load balancers, and monitoring.
    Matches Handbook Section 11 & 30.
    """
    return {
        "status": "healthy",
        "service": "SBLTrend Web Application",
        "timestamp": datetime.now().isoformat(),
        "version": "1.0.0",
    }


@app.post("/api/analyze", tags=["Analysis"])
async def analyze_sbl_file(
    file: UploadFile = File(..., description="SBL / Screening CSV, Excel, or TXT file"),
    months: int = Form(6, description="Max history in months (0 for all-time)"),
    filter_col: Optional[str] = Form(None, description="Optional column to filter on"),
    filter_val: Optional[str] = Form(None, description="Optional filter values (comma-separated)"),
):
    """
    Process an uploaded SBL data file, run deduplication & trend calculations,
    and return JSON analytics along with a download job_id for the generated Excel report.
    """
    cleanup_cache()
    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="Uploaded file is empty.")

        # Execute Core SBL Trend Pipeline
        web_payload, excel_bytes = process_sbl_data(
            file_bytes=content,
            filename=file.filename or "sbl_data.csv",
            months=months,
            filter_col=filter_col.strip() if filter_col else None,
            filter_val=filter_val.strip() if filter_val else None,
        )

        # Store Excel binary for download
        job_id = str(uuid.uuid4())
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_filename = f"SBL_Trend_Report_{ts}.xlsx"

        REPORT_CACHE[job_id] = {
            "bytes": excel_bytes,
            "filename": report_filename,
            "timestamp": time.time(),
        }

        web_payload["job_id"] = job_id
        web_payload["excel_filename"] = report_filename
        return web_payload

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Analysis failed: {str(e)}")


@app.get("/api/download/{job_id}", tags=["Analysis"])
def download_excel_report(job_id: str):
    """
    Download the generated multi-worksheet Excel report containing all tables and embedded charts.
    """
    report = REPORT_CACHE.get(job_id)
    if not report:
        raise HTTPException(
            status_code=404,
            detail="Report not found or has expired. Please run the analysis again.",
        )

    return StreamingResponse(
        io.BytesIO(report["bytes"]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{report["filename"]}"'},
    )


# ---------------------------------------------------------------------
# SERVE FRONTEND WEB DASHBOARD
# ---------------------------------------------------------------------
HTML_DASHBOARD_FILE = os.path.join(os.path.dirname(__file__), "static", "index.html")

# Mount static folder if it exists
static_dir = os.path.join(os.path.dirname(__file__), "static")
if os.path.exists(static_dir):
    app.mount("/static", StaticFiles(directory=static_dir), name="static")


@app.get("/", response_class=HTMLResponse, tags=["Frontend"])
def serve_dashboard():
    """Serve the single-page interactive web dashboard."""
    if os.path.exists(HTML_DASHBOARD_FILE):
        with open(HTML_DASHBOARD_FILE, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>SBL Trend Dashboard</h1><p>Frontend template not found. Ensure static/index.html exists.</p>"


if __name__ == "__main__":
    import uvicorn
    # Default local dev port 8000
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=True)
