# RevisorAPA — Página web

Revisión académica con criterios de APA 7.ª edición.  
Aplicación web construida con **FastAPI** + HTML/CSS/JS (estilo medieval).

## Características

- Interfaz medieval (pergaminos, tipografía Cinzel / Libre Baskerville)
- Carga de archivos `.docx`
- Revisión automática:
  - Márgenes de 1"
  - Tipografía APA permitida
  - Interlineado doble
  - Sangría de primera línea
  - Numeración de páginas
  - Sección de referencias
  - Citas narrativas y parentéticas
  - Correspondencia citas ↔ referencias
  - Sangría francesa en referencias
- Generación de documento corregido (formato APA aplicado)

## Requisitos

- Python 3.10+

## Instalación

```bash
cd revisor-apa-web
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Ejecutar en local

```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Abre en el navegador: **http://localhost:8000**

## Estructura

```
revisor-apa-web/
├── app/
│   ├── __init__.py
│   ├── main.py          # FastAPI (rutas)
│   └── apa_logic.py     # Análisis y corrección APA
├── static/
│   ├── style.css
│   └── app.js
├── templates/
│   └── index.html
├── requirements.txt
└── README.md
```

## API

| Método | Ruta            | Descripción                          |
|--------|-----------------|--------------------------------------|
| GET    | `/`             | Página principal                     |
| POST   | `/api/analyze`  | Analiza un `.docx` → JSON            |
| POST   | `/api/correct`  | Genera `.docx` corregido (descarga)  |
| GET    | `/health`       | Health check                         |

## Despliegue

Puedes desplegar en:

- **Railway / Render / Fly.io** (recomendado para FastAPI)
- **Docker** (añade un `Dockerfile` si lo necesitas)
- Cualquier VPS con Python + uvicorn + nginx

Ejemplo rápido con Railway o Render: conecta el repositorio y usa el comando de inicio:

```bash
uvicorn app.main:app --host 0.0.0.0 --port $PORT
```
