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

# Definir la raíz absoluta del proyecto evitando duplicar 'app'
CURRENT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = CURRENT_DIR.parent if CURRENT_DIR.name == "app" else CURRENT_DIR

# Buscar carpeta de plantillas (templates) en raíz o subdirectorios
TEMPLATES_DIR = PROJECT_ROOT / "templates"
if not TEMPLATES_DIR.exists():
    TEMPLATES_DIR = PROJECT_ROOT / "app" / "templates"

# Buscar carpeta de estáticos (static)
STATIC_DIR = PROJECT_ROOT / "static"
if not STATIC_DIR.exists():
    STATIC_DIR = PROJECT_ROOT / "app" / "static"

STATIC_DIR.mkdir(parents=True, exist_ok=True)
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    index_path = TEMPLATES_DIR / "index.html"
    if not index_path.exists():
        return HTMLResponse(
            f"<h3>Error de configuración</h3>"
            f"<p>No se encontró <code>index.html</code> en la ruta: <code>{TEMPLATES_DIR}</code>.</p>"
            f"<p>Asegúrate de que la carpeta <strong>templates</strong> que contiene <strong>index.html</strong> esté subida a tu repositorio.</p>",
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
