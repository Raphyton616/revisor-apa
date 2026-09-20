"""
RevisorAPA - FastAPI backend
"""

from __future__ import annotations

import re
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from io import BytesIO

from app.apa_logic import analyze_document, correct_document

# Intentamos importar pypdf/pypdf2 para procesamiento mejorado de PDF si está disponible
try:
    import pypdf
except ImportError:
    pypdf = None

BASE_DIR = Path(__file__).resolve().parent.parent

app = FastAPI(
    title="RevisorAPA",
    description="Revisión académica con criterios de APA 7.ª edición",
    version="1.1.0",
)

app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "templates")


def _is_allowed(filename: str) -> bool:
    name = (filename or "").lower()
    return name.endswith(".docx") or name.endswith(".pdf")


def _limpiar_texto_pdf(data_bytes: bytes) -> bytes:
    """
    Extrae el texto de un PDF y lo normaliza para que el analizador APA
    no pierda encabezados como 'Referencias' debido a números de página
    o saltos de línea accidentales del formateo PDF.
    """
    if not pypdf:
        return data_bytes

    try:
        reader = pypdf.PdfReader(BytesIO(data_bytes))
        lineas_limpias = []

        for page in reader.pages:
            texto_pagina = page.extract_text() or ""
            for linea in texto_pagina.splitlines():
                linea_str = linea.strip()
                # Descarta números de página aislados antes o después del título
                if re.match(r"^\d+$", linea_str):
                    continue
                if linea_str:
                    lineas_limpias.append(linea_str)

        # Unimos las líneas asegurando que 'Referencias' quede limpio en su propio párrafo
        texto_unificado = "\n".join(lineas_limpias)
        
        # Normaliza variaciones comunes como "5 Referencias" o "Referencias "
        texto_unificado = re.sub(
            r"(?i)(?:\b\d+\s+)?(referencias|referencia bibliográfica|bibliografía)\b",
            r"\n\1\n",
            texto_unificado
        )

        return texto_unificado.encode("utf-8")
    except Exception:
        # Si ocurre alguna excepción con pypdf, retorna los bytes originales
        return data_bytes


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="index.html"
    )


@app.post("/api/analyze")
async def analyze(file: UploadFile = File(...)):
    filename = file.filename or ""
    if not _is_allowed(filename):
        raise HTTPException(
            status_code=400,
            detail="Solo se aceptan archivos .docx o .pdf",
        )

    try:
        data = await file.read()
        if not data:
            raise HTTPException(status_code=400, detail="El archivo está vacío")

        # Si es un PDF, aplicamos la normalización previa del texto
        if filename.lower().endswith(".pdf"):
            data_procesada = _limpiar_texto_pdf(data)
        else:
            data_procesada = data

        result = analyze_document(data_procesada, filename=filename)
        result["filename"] = filename
        return result
    except Exception as exc:
        raise HTTPException(
            status_code=422,
            detail=f"No se pudo analizar el documento: {exc}",
        ) from exc


@app.post("/api/correct")
async def correct(file: UploadFile = File(...)):
    filename = file.filename or "documento"
    if not _is_allowed(filename):
        raise HTTPException(
            status_code=400,
            detail="Solo se aceptan archivos .docx o .pdf",
        )

    try:
        data = await file.read()
        if not data:
            raise HTTPException(status_code=400, detail="El archivo está vacío")

        corrected = correct_document(data, filename=filename)

        base_name = re.sub(r"\.(docx|pdf)$", "", filename, flags=re.I)
        output_filename = f"{base_name}_APA7_corregido.docx"

        return StreamingResponse(
            BytesIO(corrected),
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            headers={
                "Content-Disposition": f'attachment; filename="{output_filename}"'
            },
        )
    except Exception as exc:
        raise HTTPException(
            status_code=422,
            detail=f"No se pudo generar el documento corregido: {exc}",
        ) from exc


@app.get("/health")
async def health():
    return {"status": "ok"}
