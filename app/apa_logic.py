"""
Lógica de análisis y corrección APA 7.ª edición.
Soporta .docx y .pdf (el corregido siempre se entrega como .docx).
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, asdict
from typing import Iterable

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


APA_FONTS = {
    ("Times New Roman", 12),
    ("Arial", 11),
    ("Calibri", 11),
    ("Georgia", 11),
    ("Lucida Sans Unicode", 10),
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
    return [p for p in iter_paragraphs(document) if p.text.strip()]


def length_inches(value) -> float | None:
    if value is None:
        return None
    try:
        return float(value.inches)
    except (AttributeError, TypeError, ValueError):
        return None


def paragraph_font_samples(paragraphs: list) -> list[tuple[str, float]]:
    samples: list[tuple[str, float]] = []
    for paragraph in paragraphs:
        for run in paragraph.runs:
            if not run.text.strip():
                continue
            font_name = run.font.name or "Times New Roman"
            size = run.font.size.pt if run.font.size else 12.0
            samples.append((font_name, round(size, 1)))
    return samples


def has_page_field(document: Document) -> bool:
    for section in document.sections:
        header_xml = section.header._element.xml
        if "PAGE" in header_xml.upper():
            return True
    return False


def is_reference_heading(text: str) -> bool:
    """Detecta el título de referencias de forma flexible e inmune a números de página pegados."""
    text_clean = text.strip()
    text_clean = re.sub(r"^\d+\s+|\s+\d+$", "", text_clean)
    return bool(re.search(r"^\s*(referencias|references|bibliografía|bibliografia)\b", text_clean, re.I))


def find_reference_start(paragraphs: list) -> int | None:
    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip() if hasattr(paragraph, "text") else str(paragraph).strip()
        if is_reference_heading(text):
            return index
    return None


def extract_citations(text: str) -> list[dict]:
    citations: list[dict] = []

    parenthetical_groups = re.findall(
        r"\(([^()]*?(?:19|20)\d{2}[a-z]?(?:[^()]*)?)\)",
        text,
    )
    for group in parenthetical_groups:
        for part in re.split(r"\s*;\s*", group):
            match = re.search(
                r"(?P<autor>[A-ZÁÉÍÓÚÑ][^()]*?),\s*"
                r"(?P<anio>(?:19|20)\d{2}[a-z]?)\b",
                part,
            )
            if not match:
                continue
            autor = match.group("autor").strip()
            anio = match.group("anio")
            citations.append(
                {
                    "autor": autor,
                    "anio": anio,
                    "tipo": "parentética",
                    "texto": f"({part.strip()})",
                }
            )

    narrative = re.findall(
        r"\b([A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúüñ'-]{1,}"
        r"(?:\s+(?:y|e|&|and)\s+"
        r"[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúüñ'-]{1,})?)"
        r"\s+\((\d{4}[a-z]?)\)",
        text,
    )
    for autor, anio in narrative:
        citations.append(
            {
                "autor": autor.strip(),
                "anio": anio,
                "tipo": "narrativa",
                "texto": f"{autor.strip()} ({anio})",
            }
        )

    unique = []
    seen = set()
    for citation in citations:
        key = (citation["autor"].lower(), citation["anio"])
        if key not in seen:
            seen.add(key)
            unique.append(citation)
    return unique


def first_author_surname(author: str) -> str:
    cleaned = re.sub(r"\bet\s+al\.?\b", "", author, flags=re.I)
    cleaned = re.split(r"\s+(?:y|e|and|&)\s+", cleaned, maxsplit=1, flags=re.I)[0]
    cleaned = cleaned.split(",", 1)[0].strip().lower()
    return re.sub(r"[^a-záéíóúüñ0-9 -]", "", cleaned).strip()


def extract_reference_entries(reference_paragraphs: list) -> list[dict]:
    entries: list[dict] = []
    for paragraph in reference_paragraphs:
        text = paragraph.text.strip() if hasattr(paragraph, "text") else str(paragraph).strip()
        if not text:
            continue
        year_match = re.search(r"\b((?:19|20)\d{2}[a-z]?)\b", text)
        if "," in text:
            author = text.split(",", 1)[0].strip()
        else:
            author = text.split(".", 1)[0].strip()
        entries.append(
            {
                "autor": author,
                "anio": year_match.group(1) if year_match else "",
                "texto": text,
                "clave": (
                    first_author_surname(author),
                    year_match.group(1) if year_match else "",
                ),
            }
        )
    return entries


def _is_pdf(data: bytes, filename: str = "") -> bool:
    if filename.lower().endswith(".pdf"):
        return True
    return data[:5] == b"%PDF-"


def extract_text_from_pdf(data: bytes) -> str:
    if pdfplumber is None:
        raise RuntimeError("pdfplumber no está instalado")
    texts = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            if page_text.strip():
                texts.append(page_text)
    return "\n\n".join(texts) 

def text_to_paragraphs(text: str) -> list[str]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    raw = re.split(r"\n\s*\n", text)
    paragraphs = []
    for block in raw:
        block = block.strip()
        if not block:
            continue
        lines = [ln.strip() for ln in block.split("\n") if ln.strip()]
        
        # Filtra números de página solos generados por la extracción del PDF
        lines = [ln for ln in lines if not re.match(r"^\d+$", ln)]
        
        if not lines:
            continue
            
        if len(lines) <= 1:
            paragraphs.append(lines[0])
        else:
            current = lines[0]
            for ln in lines[1:]:
                # Si una línea es el encabezado de referencias, la separa inmediatamente
                if is_reference_heading(ln):
                    if current:
                        paragraphs.append(current)
                    current = ln
                    continue

                if current.endswith((".", "?", "!", ":", ";")) or len(current) > 80:
                    paragraphs.append(current)
                    current = ln
                else:
                    current = current + " " + ln
            if current:
                paragraphs.append(current)
    return paragraphs


def analyze_from_text(paragraphs: list[str], is_pdf: bool = False) -> dict:
    text = "\n".join(paragraphs)
    reference_start = None
    for i, p in enumerate(paragraphs):
        if is_reference_heading(p):
            reference_start = i
            break

    body_paragraphs = paragraphs[:reference_start] if reference_start is not None else paragraphs
    reference_paragraphs = paragraphs[reference_start + 1:] if reference_start is not None else []

    checks: list[Check] = []

    if is_pdf:
        checks.append(
            Check(
                "Formato", "Origen del documento",
                "ok",
                "El archivo es un PDF. Se analizó el texto extraído.",
                "El documento corregido se entregará como .docx con formato APA aplicado.",
            )
        )
        checks.append(
            Check(
                "Formato", "Márgenes, tipografía e interlineado",
                "warning",
                "No se pueden verificar márgenes, tipografía ni interlineado de forma fiable en un PDF.",
                "Al generar el documento corregido se aplicarán márgenes de 2,54 cm, Times New Roman 12 e interlineado doble.",
            )
        )

    references_heading_ok = reference_start is not None
    checks.append(
        Check(
            "Referencias", "Sección de referencias",
            "ok" if references_heading_ok else "error",
            "Se encontró la sección «Referencias»." if references_heading_ok else "No se encontró un encabezado exacto «Referencias».",
            "Añade un encabezado «Referencias» al final del trabajo y coloca allí las fuentes citadas.",
        )
    )

    citation_matches = extract_citations("\n".join(body_paragraphs))
    citation_count = len(citation_matches)
    reference_entries = extract_reference_entries(reference_paragraphs)
    citation_keys = {(first_author_surname(c["autor"]), c["anio"]) for c in citation_matches}
    reference_keys = {e["clave"] for e in reference_entries}
    citations_without_reference = [
        c for c in citation_matches
        if (first_author_surname(c["autor"]), c["anio"]) not in reference_keys
    ]
    uncited_references = [
        e for e in reference_entries
        if not e["anio"] or e["clave"] not in citation_keys
    ]

    checks.append(
        Check(
            "Citas", "Citas dentro del texto",
            "ok" if citation_count > 0 else "warning",
            f"Se detectaron {citation_count} cita(s) con formato autor-año." if citation_count else "No se detectaron patrones claros de cita autor-año.",
            "Revisa que toda idea tomada de otra fuente tenga una cita narrativa o parentética.",
        )
    )

    if citation_count > 0:
        if not citations_without_reference:
            checks.append(Check(
                "Citas", "Correspondencia citas-referencias", "ok",
                "Las citas encontradas tienen una referencia con el mismo autor y año.", ""
            ))
        else:
            unmatched_labels = ", ".join(f"{c['autor']}, {c['anio']}" for c in citations_without_reference[:5])
            extra = len(citations_without_reference) - 5
            checks.append(Check(
                "Citas", "Correspondencia citas-referencias", "warning",
                f"Hay {len(citations_without_reference)} cita(s) sin una referencia coincidente: {unmatched_labels}"
                + (f" y {extra} más." if extra > 0 else "."),
                "Verifica que cada cita tenga su correspondiente entrada en la lista de referencias.",
            ))

    reference_count = len(reference_paragraphs)
    checks.append(
        Check(
            "Referencias", "Entradas bibliográficas",
            "ok" if reference_count > 0 else "error",
            f"Se detectaron {reference_count} entrada(s) después del encabezado de referencias." if reference_count else "La sección de referencias está vacía o no pudo identificarse.",
            "Incluye una entrada completa por cada fuente citada.",
        )
    )

    if reference_count > 0:
        if not citation_count:
            uncited_status, uncited_detail = "warning", "No se detectaron citas en el texto; no se puede confirmar que las referencias estén citadas."
        elif uncited_references:
            uncited_status, uncited_detail = "warning", f"Hay {len(uncited_references)} referencia(s) que no coinciden con ninguna cita detectada."
        else:
            uncited_status, uncited_detail = "ok", "Cada referencia coincide con al menos una cita detectada."
        checks.append(Check(
            "Referencias", "Referencias citadas en el texto",
            uncited_status, uncited_detail,
            "Elimina las referencias no utilizadas o añade la cita correspondiente en el texto.",
        ))

    reference_years = sum(
        bool(re.search(r"\b(19|20)\d{2}[a-z]?\b", p if isinstance(p, str) else getattr(p, "text", "")))
        for p in reference_paragraphs
    )
    reference_content_ok = reference_count > 0 and reference_years >= max(1, int(reference_count * 0.75))
    checks.append(
        Check(
            "Referencias", "Datos básicos de las referencias",
            "ok" if reference_content_ok else "warning",
            f"{reference_years} de {reference_count} referencias contienen un año identificable.",
            "Comprueba autor, fecha, título, fuente y DOI o URL según el tipo de material.",
        )
    )

    return {
        "checks": [c.to_dict() for c in checks],
        "paragraph_count": len(paragraphs),
        "word_count": len(re.findall(r"\b[\wÁÉÍÓÚÜÑáéíóúüñ'-]+\b", text)),
        "citation_count": citation_count,
        "citations": citation_matches,
        "reference_count": reference_count,
        "references": reference_entries,
        "citations_without_reference": citations_without_reference,
        "uncited_references": uncited_references,
        "ok_count": sum(c.status == "ok" for c in checks),
        "issue_count": sum(c.status != "ok" for c in checks),
        "is_pdf": is_pdf,
    }


def analyze_document(data: bytes, filename: str = "") -> dict:
    if _is_pdf(data, filename):
        text = extract_text_from_pdf(data)
        paragraphs = text_to_paragraphs(text)
        return analyze_from_text(paragraphs, is_pdf=True)

    document = Document(io.BytesIO(data))
    paragraphs = non_empty_paragraphs(document)
    reference_start = find_reference_start(paragraphs)
    body_paragraphs = paragraphs[:reference_start] if reference_start is not None else paragraphs
    reference_paragraphs = paragraphs[reference_start + 1:] if reference_start is not None else []

    checks: list[Check] = []

    section = document.sections[0] if document.sections else None
    margins = (
        [length_inches(getattr(section, side, None)) for side in ("top_margin", "bottom_margin", "left_margin", "right_margin")]
        if section else []
    )
    margins_ok = bool(margins) and all(value is not None and abs(value - 1.0) <= 0.04 for value in margins)
    checks.append(
        Check(
            "Formato", "Márgenes de una pulgada",
            "ok" if margins_ok else "error",
            "Los cuatro márgenes están configurados aproximadamente a 2,54 cm." if margins_ok else "Se encontraron márgenes distintos de 2,54 cm en uno o más lados.",
            "Usa márgenes de 2,54 cm en los cuatro lados.",
        )
    )

    font_samples = paragraph_font_samples(paragraphs)
    font_ok = not font_samples or all(s in APA_FONTS for s in font_samples)
    font_label = ", ".join(f"{name} {size:g}" for name, size in sorted(set(font_samples))[:4])
    checks.append(
        Check(
            "Formato", "Tipografía legible y consistente",
            "ok" if font_ok else "warning",
            f"Se detectó: {font_label or 'tipografía heredada'}." if font_ok else "Hay tipografías que no coinciden con las opciones APA 7.",
            "Usa Times New Roman 12, Arial 11, Calibri 11, Georgia 11 o Lucida Sans Unicode 10.",
        )
    )

    spacing_values = []
    for paragraph in body_paragraphs:
        value = paragraph.paragraph_format.line_spacing
        if isinstance(value, (int, float)):
            spacing_values.append(float(value))
    spacing_ok = not spacing_values or all(abs(v - 2.0) <= 0.05 for v in spacing_values)
    checks.append(
        Check(
            "Formato", "Interlineado doble",
            "ok" if spacing_ok else "warning",
            "Los párrafos usan interlineado doble." if spacing_ok else "Hay párrafos con interlineado distinto de 2,0.",
            "Aplica interlineado doble a todo el texto.",
        )
    )

    body_without_headings = [
        p for p in body_paragraphs[1:]
        if not re.match(
            r"^(resumen|abstract|introducción|introduccion|método|metodo|resultados|discusión|discusion|conclusión|conclusion)$",
            p.text.strip(), re.I
        )
    ]
    indent_values = [
        length_inches(p.paragraph_format.first_line_indent)
        for p in body_without_headings
        if p.paragraph_format.first_line_indent is not None
    ]
    indent_ok = not indent_values or sum(abs(v - 0.5) <= 0.06 for v in indent_values) >= max(1, int(len(indent_values) * 0.65))
    checks.append(
        Check(
            "Formato", "Sangría de primera línea",
            "ok" if indent_ok else "warning",
            "La mayoría de los párrafos usa sangría de 1,27 cm." if indent_ok else "La sangría de primera línea es irregular.",
            "Aplica sangría de primera línea de 1,27 cm.",
        )
    )

    page_status = "ok" if has_page_field(document) else "warning"
    checks.append(
        Check(
            "Formato", "Numeración de páginas",
            page_status,
            "El documento contiene numeración de páginas." if page_status == "ok" else "No se detectó numeración de páginas.",
            "Añade el número de página en la esquina superior derecha.",
        )
    )

    text_result = analyze_from_text([p.text for p in paragraphs], is_pdf=False)
    text_result["checks"] = [c.to_dict() for c in checks] + [
        c for c in text_result["checks"] if c["category"] != "Formato"
    ]
    text_result["ok_count"] = sum(1 for c in text_result["checks"] if c["status"] == "ok")
    text_result["issue_count"] = sum(1 for c in text_result["checks"] if c["status"] != "ok")
    text_result["is_pdf"] = False
    return text_result


def set_run_font(run, name: str = "Times New Roman", size: int = 12) -> None:
    run.font.name = name
    run.font.size = Pt(size)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    r_fonts.set(qn("w:ascii"), name)
    r_fonts.set(qn("w:hAnsi"), name)
    r_fonts.set(qn("w:eastAsia"), name)
    r_fonts.set(qn("w:cs"), name)


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instruction, separate, text, end])
    set_run_font(run, size=12)


def build_apa_docx_from_text(paragraphs: list[str]) -> bytes:
    document = Document()

    for section in document.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        header = section.header
        header_paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_paragraph.clear()
        add_page_number(header_paragraph)

    try:
        normal = document.styles["Normal"]
        normal.font.name = "Times New Roman"
        normal.font.size = Pt(12)
        if normal._element.rPr is not None and normal._element.rPr.rFonts is not None:
            normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        normal.paragraph_format.line_spacing = 2
        normal.paragraph_format.space_after = Pt(0)
    except KeyError:
        pass

    reference_mode = False
    for text in paragraphs:
        text = text.strip()
        if not text:
            continue

        if is_reference_heading(text):
            reference_mode = True
            p = document.add_paragraph()
            run = p.add_run("Referencias")
            set_run_font(run)
            p.paragraph_format.line_spacing = 2
            p.paragraph_format.first_line_indent = Inches(0)
            p.paragraph_format.left_indent = Inches(0)
            continue

        p = document.add_paragraph()
        run = p.add_run(text)
        set_run_font(run)
        p.paragraph_format.line_spacing = 2
        p.paragraph_format.space_after = Pt(0)

        is_heading = text.lower() in {
            "resumen", "abstract", "introducción", "introduccion",
            "método", "metodo", "resultados", "discusión", "discusion",
            "conclusión", "conclusion", "referencias"
        }

        if is_heading:
            p.paragraph_format.first_line_indent = Inches(0)
            p.paragraph_format.left_indent = Inches(0)
        elif reference_mode:
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.5)
        else:
            p.paragraph_format.left_indent = Inches(0)
            p.paragraph_format.first_line_indent = Inches(0.5)

    output = io.BytesIO()
    document.save(output)
    return output.getvalue()


def correct_document(data: bytes, filename: str = "") -> bytes:
    if _is_pdf(data, filename):
        text = extract_text_from_pdf(data)
        paragraphs = text_to_paragraphs(text)
        return build_apa_docx_from_text(paragraphs)

    document = Document(io.BytesIO(data))
    for section in document.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        header = section.header
        header_paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        if "PAGE" not in header._element.xml.upper():
            header_paragraph.clear()
            add_page_number(header_paragraph)

    try:
        normal_style = document.styles["Normal"]
        normal_style.font.name = "Times New Roman"
        normal_style.font.size = Pt(12)
        if normal_style._element.rPr is not None and normal_style._element.rPr.rFonts is not None:
            normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        normal_style.paragraph_format.line_spacing = 2
        normal_style.paragraph_format.space_after = Pt(0)
    except KeyError:
        pass

    paragraphs = non_empty_paragraphs(document)
    reference_start = find_reference_start(paragraphs)
    reference_mode = False

    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        if reference_start is not None and index == reference_start:
            reference_mode = True
        paragraph.paragraph_format.line_spacing = 2
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            set_run_font(run)

        is_heading = (
            (paragraph.style and paragraph.style.name and paragraph.style.name.lower().startswith("heading"))
            or text.lower() in {
                "resumen", "abstract", "introducción", "introduccion",
                "método", "metodo", "resultados", "discusión", "discusion",
                "conclusión", "conclusion", "referencias"
            }
        )
        if is_heading:
            paragraph.paragraph_format.first_line_indent = Inches(0)
            paragraph.paragraph_format.left_indent = Inches(0)
            continue

        if reference_mode:
            paragraph.paragraph_format.left_indent = Inches(0.5)
            paragraph.paragraph_format.first_line_indent = Inches(-0.5)
        else:
            paragraph.paragraph_format.left_indent = Inches(0)
            paragraph.paragraph_format.first_line_indent = Inches(0.5)

    output = io.BytesIO()
    document.save(output)
    return output.getvalue()
