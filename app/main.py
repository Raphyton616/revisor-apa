"""
Servidor FastAPI para RevisorAPA (Solo .docx)
"""

import os
from pathlib import Path
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

from app.apa_logic import analyze_document, correct_document

app = FastAPI(title="RevisorAPA", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates" if (BASE_DIR / "templates").exists() else BASE_DIR / "app" / "templates"
STATIC_DIR = BASE_DIR / "static" if (BASE_DIR / "static").exists() else BASE_DIR / "app" / "static"

STATIC_DIR.mkdir(parents=True, exist_ok=True)
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    if not (TEMPLATES_DIR / "index.html").exists():
        return HTMLResponse(
            f"<h3>Error de plantilla</h3><p>No se encontró <code>index.html</code> en <code>{TEMPLATES_DIR}</code>.</p>",
            status_code=500,
        )
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/api/analyze")
async def api_analyze(file: UploadFile = File(...)):
    if not (file.filename or "").lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="¡Solo se admite formato .docx! ⚠️")
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="El archivo está vacío.")
    return analyze_document(data, file.filename or "")


@app.post("/api/correct")
async def api_correct(file: UploadFile = File(...)):
    if not (file.filename or "").lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="¡Solo se admite formato .docx! ⚠️")
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="El archivo está vacío.")
    
    corrected = correct_document(data, file.filename or "")
    out_name = f"{Path(file.filename or 'doc').stem}_APA7_Corregido.docx"
    return Response(
        content=corrected,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
        )
