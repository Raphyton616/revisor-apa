"""
Servidor FastAPI para RevisorAPA.
Procesa análisis y corrección exclusiva de archivos .docx.
"""

from __future__ import annotations

import os
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

templates = Jinja2Templates(directory="templates")
if os.path.exists("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/api/analyze")
async def api_analyze(file: UploadFile = File(...)):
    filename = file.filename or ""
    if not filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail="¡Solo se admite formato .docx! ⚠️",
        )

    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="El archivo subido está vacío.")

    try:
        result = analyze_document(data, filename)
        return result
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error al analizar el documento Word: {str(e)}",
        )


@app.post("/api/correct")
async def api_correct(file: UploadFile = File(...)):
    filename = file.filename or ""
    if not filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail="¡Solo se admite formato .docx! ⚠️",
        )

    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="El archivo subido está vacío.")

    try:
        corrected_data = correct_document(data, filename)
        base_name = os.path.splitext(filename)[0]
        output_name = f"{base_name}_APA7_Corregido.docx"

        return Response(
            content=corrected_data,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{output_name}"'
            },
        )
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error al generar el documento corregido: {str(e)}",
)
