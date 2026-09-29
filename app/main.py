"""Servidor FastAPI para RevisorAPA."""

from __future__ import annotations

import logging
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

from app.apa_logic import analyze_document, correct_document


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

MAX_FILE_SIZE = 10 * 1024 * 1024
DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

app = FastAPI(
    title="RevisorAPA",
    version="2.1.0",
    description="Revisión automática de formato APA 7 para documentos DOCX.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:8000",
        "http://127.0.0.1:8000",
        "http://localhost:3000",
        "http://127.0.0.1:3000",
    ],
    allow_credentials=True,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

CURRENT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = CURRENT_DIR.parent
TEMPLATES_DIR = PROJECT_ROOT / "templates"
STATIC_DIR = PROJECT_ROOT / "static"

if not TEMPLATES_DIR.exists():
    raise RuntimeError(f"No existe la carpeta de plantillas: {TEMPLATES_DIR}")

if not STATIC_DIR.exists():
    raise RuntimeError(f"No existe la carpeta de archivos estáticos: {STATIC_DIR}")

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


async def read_docx_upload(file: UploadFile) -> tuple[bytes, str]:
    filename = file.filename or ""

    if not filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail="Solo se admiten archivos con extensión .docx.",
        )

    data = await file.read(MAX_FILE_SIZE + 1)

    if not data:
        raise HTTPException(status_code=400, detail="El archivo está vacío.")

    if len(data) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=413,
            detail="El archivo no puede superar los 10 MB.",
        )

    return data, filename


@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/health")
async def health_check():
    return JSONResponse({
        "status": "ok",
        "service": "RevisorAPA",
        "version": "2.1.0",
    })


@app.post("/api/analyze")
async def api_analyze(file: UploadFile = File(...)):
    data, filename = await read_docx_upload(file)

    try:
        return analyze_document(data, filename)
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error
    except Exception as error:
        logger.exception("Error al analizar el documento %s", filename)
        raise HTTPException(
            status_code=400,
            detail="El archivo no es un documento DOCX válido o está dañado.",
        ) from error


@app.post("/api/correct")
async def api_correct(file: UploadFile = File(...)):
    data, filename = await read_docx_upload(file)

    try:
        corrected = correct_document(data, filename)
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error
    except Exception as error:
        logger.exception("Error al corregir el documento %s", filename)
        raise HTTPException(
            status_code=400,
            detail="No fue posible generar el documento corregido. Verifica que el DOCX sea válido.",
        ) from error

    output_name = f"{Path(filename).stem}_APA7_Corregido.docx"
    return Response(
        content=corrected,
        media_type=DOCX_MEDIA_TYPE,
        headers={
            "Content-Disposition": f'attachment; filename="{output_name}"',
            "Content-Length": str(len(corrected)),
        },
)
