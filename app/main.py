"""
Servidor FastAPI para RevisorAPA (Solo .docx)
"""

import os
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, Response, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

from app.apa_logic import analyze_document, correct_document

app = FastAPI(title="RevisorAPA", version="2.0.0")

# CORS más seguro para desarrollo local
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:8000",
        "http://127.0.0.1:8000",
        "http://localhost:3000",
        "http://127.0.0.1:3000",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Raíz del proyecto (un nivel arriba de la carpeta "app")
CURRENT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = CURRENT_DIR.parent

TEMPLATES_DIR = PROJECT_ROOT / "templates"
STATIC_DIR = PROJECT_ROOT / "static"

# Crear carpeta static si no existe
STATIC_DIR.mkdir(parents=True, exist_ok=True)

# Montar archivos estáticos
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
            f"<p>Asegúrate de que la carpeta <strong>templates</strong> que contiene <strong>index.html</strong> esté en la raíz del proyecto.</p>",
            status_code=500,
        )
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/health")
async def health_check():
    return JSONResponse({"status": "ok", "service": "RevisorAPA", "version": "2.0.0"})


@app.post("/api/analyze")
async def api_analyze(file: UploadFile = File(...)):
    if not (file.filename or "").lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="¡Solo se admite formato .docx! ⚠️")

    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="El archivo está vacío.")

    try:
        return analyze_document(data, file.filename or "")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al analizar el documento: {str(e)}")


@app.post("/api/correct")
async def api_correct(file: UploadFile = File(...)):
    if not (file.filename or "").lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="¡Solo se admite formato .docx! ⚠️")

    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="El archivo está vacío.")

    try:
        corrected = correct_document(data, file.filename or "")
        out_name = f"{Path(file.filename or 'doc').stem}_APA7_Corregido.docx"
        return Response(
            content=corrected,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al corregir el documento: {str(e)}")
