# RevisorAPA — Página web

Revisión académica con criterios de APA 7.ª edición.

Aplicación web construida con **FastAPI** + HTML/CSS/JS (estilo medieval).

## Características

- Interfaz medieval (pergaminos, tipografía Cinzel / Libre Baskerville)
- Carga de archivos `.docx`
- Revisión automática:
  - Márgenes de 1" (2,54 cm)
  - Tipografía APA permitida
  - Interlineado doble
  - Sangría de primera línea
  - Numeración de páginas
  - Sección de referencias
  - Citas narrativas y parentéticas
  - Correspondencia citas ↔ referencias
  - Listado de citas sin referencia y referencias no citadas
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
