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


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="index.html"
    )


@app.post("/api/analyze")
async def analyze(file: UploadFile = File(...)):
    if not _is_allowed(file.filename or ""):
        raise HTTPException(
            status_code=400,
            detail="Solo se aceptan archivos .docx o .pdf",
        )

    try:
        data = await file.read()
        if not data:
            raise HTTPException(status_code=400, detail="El archivo está vacío")
        result = analyze_document(data, filename=file.filename or "")
        result["filename"] = file.filename
        return result
    except Exception as exc:
        raise HTTPException(
            status_code=422,
            detail=f"No se pudo analizar el documento: {exc}",
        ) from exc


@app.post("/api/correct")
async def correct(file: UploadFile = File(...)):
    if not _is_allowed(file.filename or ""):
        raise HTTPException(
            status_code=400,
            detail="Solo se aceptan archivos .docx o .pdf",
        )

    try:
        data = await file.read()
        if not data:
            raise HTTPException(status_code=400, detail="El archivo está vacío")
        corrected = correct_document(data, filename=file.filename or "")

        base_name = re.sub(r"\.(docx|pdf)$", "", file.filename or "documento", flags=re.I)
        filename = f"{base_name}_APA7_corregido.docx"

        return StreamingResponse(
            BytesIO(corrected),
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"'
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
