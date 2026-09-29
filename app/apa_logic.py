"""Lógica de análisis y corrección APA 7 para archivos .docx."""

from __future__ import annotations

import io
import re
from dataclasses import asdict, dataclass
from typing import Iterable

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

ALLOWED_FONTS = {
    ("Times New Roman", 12.0),
    ("Arial", 11.0),
    ("Calibri", 11.0),
    ("Georgia", 11.0),
    ("Lucida Sans Unicode", 10.0),
}

HEADING_WORDS = {
    "resumen", "abstract", "introducción", "introduccion", "método", "metodo",
    "resultados", "discusión", "discusion", "conclusión", "conclusion",
}


@dataclass
class Check:
    category: str
    title: str
    status: str
    detail: str
    recommendation: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


def iter_paragraphs(document: Document) -> Iterable:
    for paragraph in document.paragraphs:
        yield paragraph
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs


def non_empty_paragraphs(document: Document) -> list:
    return [paragraph for paragraph in iter_paragraphs(document) if paragraph.text.strip()]


def length_cm(value) -> float | None:
    if value is None:
        return None
    try:
        return round(float(value.cm), 2)
    except (AttributeError, TypeError, ValueError):
        return None


def effective_font(run) -> tuple[str, float]:
    name = run.font.name or "Times New Roman"
    size = run.font.size.pt if run.font.size else 12.0
    return str(name), round(float(size), 1)


def paragraph_font_samples(paragraphs: list) -> list[tuple[str, float]]:
    samples = []
    for paragraph in paragraphs:
        for run in paragraph.runs:
            if run.text.strip():
                samples.append(effective_font(run))
    return samples


def has_page_field(document: Document) -> bool:
    for section in document.sections:
        try:
            if "PAGE" in section.header._element.xml.upper():
                return True
        except Exception:
            continue
    return False


def is_reference_heading(text: str) -> bool:
    cleaned = re.sub(r"^\s*\d+[.)]?\s*|\s*\d+\s*$", "", text.strip())
    return bool(re.fullmatch(
        r"(?:referencias(?:\s+bibliográficas?)?|references|bibliografía|bibliografia)",
        cleaned,
        flags=re.IGNORECASE,
    ))


def find_reference_start(paragraphs: list) -> tuple[int | None, str]:
    # La sección de referencias normalmente aparece al final; buscar desde atrás
    # evita confundir una mención intermedia con el encabezado real.
    for index in range(len(paragraphs) - 1, -1, -1):
        text = paragraphs[index].text.strip()
        if is_reference_heading(text):
            return index, text
    return None, ""


def first_author_surname(author: str) -> str:
    if not isinstance(author, str):
        return ""
    cleaned = re.sub(r"\bet\s+al\.?\b", "", author, flags=re.IGNORECASE)
    cleaned = re.split(r"\s+(?:y|e|and|&)\s+", cleaned, maxsplit=1, flags=re.IGNORECASE)[0]
    cleaned = cleaned.split(",", 1)[0].strip().lower()
    return re.sub(r"[^a-záéíóúüñ0-9 -]", "", cleaned).strip()


def extract_citations(text: str) -> list[dict]:
    citations = []
    parenthetical = re.findall(r"\(([^()]{2,240}?\b(?:19|20)\d{2}[a-z]?[^()]*)\)", text)
    for group in parenthetical:
        for part in re.split(r"\s*;\s*", group):
            match = re.search(
                r"(?P<autor>[A-ZÁÉÍÓÚÑ][^,;()]{1,100}?),\s*"
                r"(?P<anio>(?:19|20)\d{2}[a-z]?)\b",
                part,
                flags=re.IGNORECASE,
            )
            if match:
                citations.append({
                    "autor": match.group("autor").strip(),
                    "anio": match.group("anio").strip(),
                    "tipo": "parentética",
                    "texto": f"({part.strip()})",
                })

    narrative = re.findall(
        r"\b([A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúüñ'’-]{1,}"
        r"(?:\s+(?:y|e|&|and)\s+[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúüñ'’-]{1,})?"
        r"(?:\s+et\s+al\.?)?)\s+\(((?:19|20)\d{2}[a-z]?)\)",
        text,
    )
    for autor, anio in narrative:
        citations.append({
            "autor": autor.strip(),
            "anio": anio.strip(),
            "tipo": "narrativa",
            "texto": f"{autor.strip()} ({anio})",
        })

    unique = []
    seen = set()
    for citation in citations:
        key = (first_author_surname(citation["autor"]), citation["anio"])
        if key[0] and key not in seen:
            seen.add(key)
            unique.append(citation)
    return unique


def extract_reference_entries(paragraphs: list) -> list[dict]:
    entries = []
    for paragraph in paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        year_match = re.search(r"\b((?:19|20)\d{2}[a-z]?)\b", text, flags=re.IGNORECASE)
        year = year_match.group(1) if year_match else ""
        author = text.split(",", 1)[0].strip() if "," in text else text.split(".", 1)[0].strip()
        entries.append({
            "autor": author,
            "anio": year,
            "texto": text,
            "surname": first_author_surname(author),
        })
    return entries


def _check_format(document: Document, body: list, all_paragraphs: list) -> list[Check]:
    checks = []
    section = document.sections[0] if document.sections else None
    if section:
        values = {
            "Superior": length_cm(section.top_margin),
            "Inferior": length_cm(section.bottom_margin),
            "Izquierdo": length_cm(section.left_margin),
            "Derecho": length_cm(section.right_margin),
        }
        bad = [f"{name} ({value} cm)" for name, value in values.items()
               if value is None or abs(value - 2.54) > 0.1]
        checks.append(Check(
            "Formato", "Márgenes de 2,54 cm", "error" if bad else "ok",
            "Se encontraron márgenes incorrectos en: " + ", ".join(bad) + "." if bad
            else "Los cuatro márgenes están configurados correctamente a 2,54 cm (1 pulgada).",
            "Ajusta los cuatro márgenes a 2,54 cm en Disposición > Márgenes de Word." if bad else "",
        ))

    fonts = paragraph_font_samples(all_paragraphs)
    invalid = sorted(set(font for font in fonts if font not in ALLOWED_FONTS))
    checks.append(Check(
        "Formato", "Tipografía APA", "warning" if invalid else "ok",
        "Hay fuentes o tamaños que no coinciden con las opciones estándar de APA 7: " +
        ", ".join(f"{name} {size:g} pt" for name, size in invalid) + "." if invalid
        else "La tipografía detectada coincide con opciones permitidas de APA 7.",
        "Usa Times New Roman 12, Arial 11, Calibri 11, Georgia 11 o Lucida Sans Unicode 10." if invalid else "",
    ))

    spacing_missing = 0
    spacing_bad = 0
    for paragraph in body:
        value = paragraph.paragraph_format.line_spacing
        if value is None:
            spacing_missing += 1
        elif not isinstance(value, (int, float)) or abs(float(value) - 2.0) > 0.05:
            spacing_bad += 1
    spacing_status = "warning" if spacing_bad or spacing_missing else "ok"
    checks.append(Check(
        "Formato", "Interlineado doble", spacing_status,
        "Hay párrafos con interlineado distinto de 2,0." if spacing_bad else
        "El interlineado no está definido explícitamente en algunos párrafos; puede depender del estilo del documento." if spacing_missing else
        "Los párrafos del cuerpo tienen interlineado doble (2,0).",
        "Aplica interlineado doble a todo el texto del documento." if spacing_status != "ok" else "",
    ))

    candidates = [p for p in body[1:] if p.text.strip().lower() not in HEADING_WORDS]
    with_indent = sum(
        1 for p in candidates
        if length_cm(p.paragraph_format.first_line_indent) is not None
        and abs(length_cm(p.paragraph_format.first_line_indent) - 1.27) <= 0.15
    )
    indent_status = "ok" if candidates and with_indent / len(candidates) >= 0.65 else "warning"
    checks.append(Check(
        "Formato", "Sangría de primera línea", indent_status,
        f"{with_indent} de {len(candidates)} párrafos revisados tienen sangría de aproximadamente 1,27 cm."
        if candidates else "No se encontraron suficientes párrafos de cuerpo para verificar la sangría.",
        "Aplica sangría de primera línea de 1,27 cm a cada párrafo del cuerpo." if indent_status != "ok" else "",
    ))

    checks.append(Check(
        "Formato", "Numeración de páginas", "ok" if has_page_field(document) else "warning",
        "El documento contiene un campo de numeración de páginas en el encabezado."
        if has_page_field(document) else "No se detectó un campo de numeración de páginas en el encabezado.",
        "Añade el número de página alineado a la derecha en el encabezado superior." if not has_page_field(document) else "",
    ))
    return checks


def analyze_document(data: bytes, filename: str = "") -> dict:
    if not filename.lower().endswith(".docx"):
        raise ValueError("La aplicación está optimizada exclusivamente para archivos .docx.")
    document = Document(io.BytesIO(data))
    paragraphs = non_empty_paragraphs(document)
    reference_start, heading = find_reference_start(paragraphs)
    body = paragraphs[:reference_start] if reference_start is not None else paragraphs
    reference_paragraphs = paragraphs[reference_start + 1:] if reference_start is not None else []
    checks = _check_format(document, body, paragraphs)

    if reference_start is None:
        checks.append(Check("Referencias", "Sección de referencias", "error",
            "No se encontró un encabezado independiente de Referencias.",
            "Crea una página final titulada exactamente «Referencias»."))
    else:
        good_heading = heading.strip().lower() == "referencias"
        checks.append(Check("Referencias", "Sección de referencias", "ok" if good_heading else "warning",
            "Se encontró la sección «Referencias»." if good_heading else f"Se detectó el título «{heading}».",
            "En APA 7, usa exactamente el título «Referencias»." if not good_heading else ""))

    full_text = "\n".join(p.text for p in body)
    citations = extract_citations(full_text)
    references = extract_reference_entries(reference_paragraphs)
    citation_keys = {(first_author_surname(c["autor"]), c["anio"]) for c in citations}
    reference_keys = {(r["surname"], r["anio"]) for r in references}
    missing = [c for c in citations if (first_author_surname(c["autor"]), c["anio"]) not in reference_keys]
    uncited = [r for r in references if not r["anio"] or (r["surname"], r["anio"]) not in citation_keys]

    checks.append(Check("Citas", "Citas dentro del texto", "ok" if citations else "warning",
        f"Se detectaron {len(citations)} cita(s) con patrón autor-año." if citations else
        "No se detectaron patrones claros de cita autor-año.",
        "Revisa que las afirmaciones tomadas de fuentes externas tengan su cita." if not citations else ""))
    if reference_start is not None:
        checks.append(Check("Citas", "Correspondencia citas-referencias", "ok" if not missing else "warning",
            "Todas las citas detectadas tienen una referencia coincidente." if not missing else
            f"Hay {len(missing)} cita(s) sin referencia coincidente.",
            "Verifica cada cita contra la lista de referencias." if missing else ""))
    checks.append(Check("Referencias", "Entradas bibliográficas", "ok" if references else "warning",
        f"Se detectaron {len(references)} entrada(s) en Referencias." if references else
        "La lista de referencias está vacía o no pudo identificarse.",
        "Incluye una entrada bibliográfica completa por cada fuente citada." if not references else ""))

    return {
        "checks": [check.to_dict() for check in checks],
        "paragraph_count": len(paragraphs),
        "word_count": len(re.findall(r"\b[\wÁÉÍÓÚÜÑáéíóúüñ'-]+\b", full_text)),
        "citation_count": len(citations),
        "citations": citations,
        "reference_count": len(references),
        "references": references,
        "citations_without_reference": missing,
        "uncited_references": uncited,
        "ok_count": sum(check.status == "ok" for check in checks),
        "issue_count": sum(check.status != "ok" for check in checks),
        "is_pdf": False,
    }


def set_run_font(run, name: str = "Times New Roman", size: int = 12) -> None:
    run.font.name = name
    run.font.size = Pt(size)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
        r_fonts.set(qn(f"w:{attribute}"), name)


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run()
    for field_type, text in (("begin", None), ("instrText", " PAGE "), ("separate", None), ("text", "1"), ("end", None)):
        element = OxmlElement("w:fldChar" if field_type in {"begin", "separate", "end"} else "w:t" if field_type == "text" else "w:instrText")
        if field_type == "instrText":
            element.set(qn("xml:space"), "preserve")
        if field_type in {"begin", "separate", "end"}:
            element.set(qn("w:fldCharType"), field_type)
        else:
            element.text = text
        run._r.append(element)
    set_run_font(run)


def correct_document(data: bytes, filename: str = "") -> bytes:
    if not filename.lower().endswith(".docx"):
        raise ValueError("Solo se pueden corregir archivos .docx.")
    document = Document(io.BytesIO(data))

    for section in document.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        header = section.header
        paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        if "PAGE" not in header._element.xml.upper():
            paragraph.clear()
            add_page_number(paragraph)

    try:
        normal = document.styles["Normal"]
        normal.font.name = "Times New Roman"
        normal.font.size = Pt(12)
        normal.paragraph_format.line_spacing = 2
        normal.paragraph_format.space_after = Pt(0)
    except KeyError:
        pass

    paragraphs = non_empty_paragraphs(document)
    reference_start, _ = find_reference_start(paragraphs)
    reference_mode = False
    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        if reference_start is not None and index == reference_start:
            reference_mode = True
            paragraph.text = "Referencias"
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Inches(0)
            paragraph.paragraph_format.left_indent = Inches(0)
            paragraph.paragraph_format.line_spacing = 2
            for run in paragraph.runs:
                set_run_font(run)
                run.bold = True
            continue

        paragraph.paragraph_format.line_spacing = 2
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            set_run_font(run)
        is_heading = text.lower() in HEADING_WORDS or (
            paragraph.style and paragraph.style.name and paragraph.style.name.lower().startswith("heading")
        )
        if is_heading:
            paragraph.paragraph_format.first_line_indent = Inches(0)
            paragraph.paragraph_format.left_indent = Inches(0)
        elif reference_mode:
            paragraph.paragraph_format.left_indent = Inches(0.5)
            paragraph.paragraph_format.first_line_indent = Inches(-0.5)
        else:
            paragraph.paragraph_format.left_indent = Inches(0)
            paragraph.paragraph_format.first_line_indent = Inches(0.5)

    if reference_start is None:
        document.add_page_break()
        heading = document.add_paragraph("Referencias")
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.paragraph_format.line_spacing = 2
        for run in heading.runs:
            set_run_font(run)
            run.bold = True
        placeholder = document.add_paragraph(
            "[Añade aquí las fuentes bibliográficas citadas ordenadas alfabéticamente]."
        )
        placeholder.paragraph_format.line_spacing = 2
        placeholder.paragraph_format.left_indent = Inches(0.5)
        placeholder.paragraph_format.first_line_indent = Inches(-0.5)
        for run in placeholder.runs:
            set_run_font(run)

    output = io.BytesIO()
    document.save(output)
    return output.getvalue()
