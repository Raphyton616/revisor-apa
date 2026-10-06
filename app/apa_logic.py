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

from app.instructions import apply_custom_corrections, check_custom_instructions, parse_instructions

ALLOWED_FONTS = {
    ("Times New Roman", 12.0),
    ("Arial", 11.0),
    ("Calibri", 11.0),
    ("Georgia", 11.0),
    ("Lucida Sans Unicode", 10.0),
}

HEADING_WORDS = {
    "resumen", "abstract", "introducción", "introduccion", "método", "metodo",
    "resultados", "desarrollo", "discusión", "discusion", "conclusión", "conclusion",
}


STATUS_LABELS = {
    "ok": "Verificado",
    "warning": "Revisión requerida",
    "error": "Incumplimiento detectado",
    "not_evaluable": "No evaluable",
}


@dataclass
class Check:
    category: str
    title: str
    status: str
    detail: str
    recommendation: str = ""
    confidence: str = ""

    def to_dict(self) -> dict:
        result = asdict(self)
        result["status_label"] = STATUS_LABELS.get(self.status, "Revisión requerida")
        result["evidence"] = self.detail
        result["confidence"] = self.confidence or {
            "ok": "Alta",
            "error": "Alta",
            "warning": "Media",
            "not_evaluable": "Baja",
        }.get(self.status, "Media")
        return result


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


def effective_line_spacing(paragraph) -> tuple[float | None, bool]:
    """Obtiene el interlineado efectivo y si proviene de un estilo heredado."""
    direct = paragraph.paragraph_format.line_spacing
    if isinstance(direct, (int, float)):
        return float(direct), False
    style = paragraph.style
    visited = set()
    while style is not None and id(style) not in visited:
        visited.add(id(style))
        value = style.paragraph_format.line_spacing
        if isinstance(value, (int, float)):
            return float(value), True
        style = style.base_style
    return None, True


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
    """Devuelve una clave comparable para enlazar citas y referencias."""
    if not isinstance(author, str):
        return ""
    cleaned = re.sub(r"\bet\s+al\.?\b", "", author, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+\[.*?\]", "", cleaned)
    cleaned = re.split(r"\s+(?:y|e|and|&)\s+", cleaned, maxsplit=1, flags=re.IGNORECASE)[0]
    cleaned = cleaned.split(",", 1)[0].strip().lower()
    return re.sub(r"[^a-záéíóúüñ0-9 -]", "", cleaned).strip()


def _citation_key(author: str, year: str) -> tuple[str, str]:
    return first_author_surname(author), str(year).lower().strip()


def _split_parenthetical_parts(group: str) -> list[str]:
    """Separa varias fuentes en una cita parentética sin romper años 2020a."""
    parts = re.split(r"\s*;\s*", group)
    result = []
    for part in parts:
        part = part.strip()
        if re.search(r"\b(?:19|20)\d{2}[a-z]?\b", part, re.IGNORECASE):
            result.append(part)
    return result


def extract_citations(text: str) -> list[dict]:
    citations = []

    # Citas parentéticas: (García, 2020), (García & López, 2020, p. 15), etc.
    parenthetical_groups = re.findall(
        r"\(([^()]{2,260}?\b(?:19|20)\d{2}[a-z]?[^()]*)\)",
        text,
        flags=re.IGNORECASE,
    )
    for group in parenthetical_groups:
        for part in _split_parenthetical_parts(group):
            match = re.search(
                r"(?P<autor>[A-ZÁÉÍÓÚÑ][^,;()]{1,120}?),\s*"
                r"(?P<anio>(?:19|20)\d{2}[a-z]?)\b",
                part,
                flags=re.IGNORECASE,
            )
            if not match:
                continue
            author = match.group("autor").strip()
            year = match.group("anio").strip()
            citations.append({
                "autor": author,
                "anio": year,
                "tipo": "parentética",
                "texto": f"({part})",
                "pagina": _extract_page(part),
                "clave": _citation_key(author, year),
            })

    # Citas narrativas: García (2020), García y López (2020), García et al. (2020).
    narrative = re.finditer(
        r"\b(?:(?:según|como\s+señala|de\s+acuerdo\s+con)\s+)?"
        r"(?P<autor>[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúüñ'’-]{1,}"
        r"(?:\s+[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúüñ'’-]{1,}){0,5}"
        r"(?:\s+(?:y|e|&|and)\s+[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúüñ'’-]{1,})?"
        r"(?:\s+et\s+al\.?)?)\s+"
        r"\((?P<anio>(?:19|20)\d{2}[a-z]?)\)",
        text,
        flags=re.IGNORECASE,
    )
    for match in narrative:
        author = match.group("autor").strip()
        year = match.group("anio").strip()
        citations.append({
            "autor": author,
            "anio": year,
            "tipo": "narrativa",
            "texto": match.group(0),
            "pagina": "",
            "clave": _citation_key(author, year),
        })

    unique = []
    seen = set()
    for citation in citations:
        key = citation["clave"]
        if key[0] and key not in seen:
            seen.add(key)
            unique.append(citation)
    return unique


def _extract_page(text: str) -> str:
    match = re.search(r"\b(?:p|pp)\.\s*([\d-]+)", text, flags=re.IGNORECASE)
    return match.group(1) if match else ""


def _reference_type(text: str) -> str:
    lowered = text.lower()
    if "doi.org/" in lowered or "doi:" in lowered:
        return "con DOI"
    if "http://" in lowered or "https://" in lowered:
        return "con URL"
    if re.search(r"\b\d+\s*\(\s*\d+\s*\)", text):
        return "artículo de revista"
    return "referencia general"


def _reference_missing_elements(text: str, year: str, author: str) -> list[str]:
    missing = []
    if not author or len(author) < 2 or re.match(r"^(?:https?://|doi:?)", author, re.IGNORECASE):
        missing.append("autor")
    if not year:
        missing.append("año")
    # En APA, después del año debe existir al menos un título o descripción.
    if year:
        tail = text[text.lower().find(year.lower()) + len(year):].strip(" .")
        if len(tail) < 4:
            missing.append("título u obra")
    return missing


def _has_valid_doi(text: str) -> bool:
    return bool(re.search(r"(?:https?://doi\.org/|doi:\s*10\.\d{4,9}/\S+)", text, re.IGNORECASE))


def _has_valid_url(text: str) -> bool:
    return bool(re.search(r"https?://[^\s]+", text, re.IGNORECASE))


def _has_hanging_indent(paragraph) -> bool:
    left = paragraph.paragraph_format.left_indent
    first = paragraph.paragraph_format.first_line_indent
    left_cm = length_cm(left)
    first_cm = length_cm(first)
    return (
        left_cm is not None and first_cm is not None and
        abs(left_cm - 1.27) <= 0.15 and abs(first_cm + 1.27) <= 0.15
    )


def _has_italic_text(paragraph) -> bool:
    return any(run.text.strip() and run.italic is True for run in paragraph.runs)


def _publication_pattern(text: str) -> str:
    if re.search(r"\b\d+\s*\(\s*\d+\s*\)", text):
        return "revista"
    if re.search(r"\b(?:volumen|vol\.?|tomo)\b", text, re.IGNORECASE):
        return "revista"
    return "general"


def _has_page_range(text: str) -> bool:
    return bool(re.search(r"\b(?:pp?\.?\s*)?\d{1,5}\s*[-–]\s*\d{1,5}\b", text))


def extract_reference_entries(paragraphs: list) -> list[dict]:
    entries = []
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph.text.strip()
        if not text:
            continue
        year_match = re.search(r"\b((?:19|20)\d{2}[a-z]?)\b", text, flags=re.IGNORECASE)
        year = year_match.group(1) if year_match else ""
        author = text.split(",", 1)[0].strip() if "," in text else text.split(".", 1)[0].strip()
        entries.append({
            "numero": index,
            "autor": author,
            "anio": year,
            "texto": text,
            "surname": first_author_surname(author),
            "clave": _citation_key(author, year),
            "tipo": _reference_type(text),
            "tiene_doi": _has_valid_doi(text),
            "tiene_url": _has_valid_url(text),
            "elementos_faltantes": _reference_missing_elements(text, year, author),
            "sangria_francesa": _has_hanging_indent(paragraph),
            "tiene_cursiva": _has_italic_text(paragraph),
            "patron_publicacion": _publication_pattern(text),
            "tiene_rango_paginas": _has_page_range(text),
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
        value, inherited = effective_line_spacing(paragraph)
        if value is None:
            spacing_missing += 1
        elif abs(value - 2.0) > 0.05:
            spacing_bad += 1
    spacing_status = "warning" if spacing_bad or spacing_missing else "ok"
    checks.append(Check(
        "Formato", "Interlineado doble", spacing_status,
        "Hay párrafos con interlineado distinto de 2,0." if spacing_bad else
        "El interlineado no está definido ni en los párrafos ni en sus estilos." if spacing_missing else
        "Los párrafos del cuerpo tienen interlineado doble (2,0), incluido el heredado desde los estilos." if any(effective_line_spacing(p)[1] for p in body) else
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


def _normalized_heading(text: str) -> str:
    return re.sub(r"[^a-záéíóúüñ ]", "", text.casefold()).strip()


def _is_heading(paragraph) -> bool:
    text = paragraph.text.strip()
    if not text or len(text.split()) > 12:
        return False
    style_name = (paragraph.style.name or "").casefold() if paragraph.style else ""
    return "heading" in style_name or "título" in style_name or text.casefold() in HEADING_WORDS


def _structure_checks(body: list, reference_start: int | None) -> list[Check]:
    """Revisa estructura sin afirmar que un bloque sin título sea incorrecto."""
    if not body:
        return [Check("Estructura", "Estructura del documento", "not_evaluable",
                      "No hay párrafos suficientes para evaluar la estructura académica.",
                      "Añade contenido y confirma la estructura solicitada por el docente.", "Baja")]

    checks = []
    texts = [p.text.strip() for p in body if p.text.strip()]
    normalized = [_normalized_heading(text) for text in texts]
    heading_aliases = {
        "Introducción": {"introducción", "introduccion"},
        "Desarrollo": {"desarrollo", "marco teórico", "marco teorico"},
        "Conclusión": {"conclusión", "conclusion"},
    }

    first = texts[0]
    first_is_heading = _is_heading(body[0])
    title_ok = bool(first and len(first.split()) <= 18 and not first_is_heading and len(first) >= 5)
    checks.append(Check(
        "Estructura", "Título principal", "ok" if title_ok else "warning",
        f"Se identificó como posible título principal: «{first[:100]}»." if title_ok else
        "No se pudo identificar con suficiente seguridad un título principal al inicio del documento.",
        "Confirma que el documento tenga un título principal visible." if not title_ok else "",
        "Alta" if title_ok else "Baja",
    ))

    for label, aliases in heading_aliases.items():
        exact_index = next((i for i, value in enumerate(normalized) if value in aliases), None)
        if exact_index is not None:
            checks.append(Check(
                "Estructura", label, "ok",
                f"Se detectó el encabezado «{texts[exact_index]}».",
                "", "Alta",
            ))
            continue

        if label == "Introducción":
            candidate = texts[1:4]
            has_content = len(candidate) >= 1 and sum(len(t.split()) for t in candidate) >= 25
            detail = "Se identificó contenido al inicio que podría funcionar como introducción, pero no un encabezado visible." if has_content else "No se identificó un encabezado ni suficiente contenido introductorio al inicio."
        elif label == "Desarrollo":
            candidate = texts[1:-1]
            has_content = len(candidate) >= 2 and sum(len(t.split()) for t in candidate) >= 35
            detail = "Se identificó un bloque central de contenido que podría corresponder al desarrollo, pero no un encabezado visible." if has_content else "No se identificó un bloque central suficiente para evaluar el desarrollo."
        else:
            markers = ("en conclusión", "en conclusion", "finalmente", "por lo tanto", "en síntesis", "en sintesis", "para concluir")
            candidate = texts[-3:]
            marked = any(any(marker in t.casefold() for marker in markers) for t in candidate)
            has_content = marked or (len(candidate) >= 1 and len(candidate[-1].split()) >= 25)
            detail = "Se identificó contenido final que podría funcionar como conclusión, pero no un encabezado visible." if has_content else "No se identificó un encabezado ni contenido concluyente suficiente."

        checks.append(Check(
            "Estructura", label, "warning" if has_content else "not_evaluable",
            detail,
            f"Confirma si el docente exige el encabezado visible «{label}»." if has_content else
            f"Confirma si el trabajo debe incluir una sección titulada «{label}».",
            "Media" if has_content else "Baja",
        ))
    return checks


def analyze_document(data: bytes, filename: str = "", instructions_text: str = "") -> dict:
    if not filename.lower().endswith(".docx"):
        raise ValueError("La aplicación está optimizada exclusivamente para archivos .docx.")

    document = Document(io.BytesIO(data))
    paragraphs = non_empty_paragraphs(document)
    reference_start, heading = find_reference_start(paragraphs)
    body = paragraphs[:reference_start] if reference_start is not None else paragraphs
    reference_paragraphs = paragraphs[reference_start + 1:] if reference_start is not None else []
    checks = _check_format(document, body, paragraphs)
    checks.extend(_structure_checks(body, reference_start))
    custom_rules = parse_instructions(instructions_text)
    custom_checks = check_custom_instructions(document, custom_rules)
    checks.extend(Check(**item) for item in custom_checks)

    if reference_start is None:
        checks.append(Check(
            "Referencias", "Sección de referencias", "error",
            "No se encontró un encabezado independiente de Referencias.",
            "Crea una página final titulada exactamente «Referencias».",
        ))
    else:
        good_heading = heading.strip().lower() == "referencias"
        checks.append(Check(
            "Referencias", "Sección de referencias", "ok" if good_heading else "warning",
            "Se encontró la sección «Referencias»." if good_heading else f"Se detectó el título «{heading}».",
            "En APA 7, usa exactamente el título «Referencias»." if not good_heading else "",
        ))

    full_text = "\n".join(p.text for p in body)
    citations = extract_citations(full_text)
    references = extract_reference_entries(reference_paragraphs)
    citation_keys = {c["clave"] for c in citations}
    reference_keys = {r["clave"] for r in references if r["surname"]}
    missing = [c for c in citations if c["clave"] not in reference_keys]
    uncited = [r for r in references if not r["anio"] or r["clave"] not in citation_keys]

    checks.append(Check(
        "Citas", "Citas dentro del texto", "ok" if citations else "warning",
        f"Se detectaron {len(citations)} cita(s) con patrón autor-año." if citations else
        "No se detectaron patrones claros de cita autor-año.",
        "Revisa que las afirmaciones tomadas de fuentes externas tengan su cita." if not citations else "",
    ))
    checks.append(Check(
        "Citas", "Correspondencia citas-referencias", "ok" if not missing and reference_start is not None else "warning",
        "Todas las citas detectadas tienen una referencia coincidente." if not missing and reference_start is not None else
        f"Hay {len(missing)} cita(s) sin referencia coincidente." if missing else
        "No se puede comprobar la correspondencia porque falta la sección de referencias.",
        "Verifica cada cita contra la lista de referencias." if missing or reference_start is None else "",
    ))

    quoted_fragments = re.findall(r"[«“\"]([^«»“”\"]{12,})[»”\"]", full_text)
    quoted_without_page = []
    for fragment in quoted_fragments:
        position = full_text.find(fragment)
        nearby = full_text[position:position + len(fragment) + 180]
        if not re.search(r"\b(?:p|pp)\.\s*\d", nearby, re.IGNORECASE):
            quoted_without_page.append(fragment[:45])
    checks.append(Check(
        "Citas", "Citas textuales y páginas", "warning" if quoted_without_page else "ok",
        f"Se detectaron {len(quoted_without_page)} posible(s) cita(s) textual(es) sin página." if quoted_without_page else
        "No se detectaron citas textuales sin número de página en el patrón revisado.",
        "Añade p. o pp. con la página de la cita textual." if quoted_without_page else "",
    ))

    author_patterns = []
    if any(re.search(r"\bet\s+al\.?", c["autor"], re.IGNORECASE) for c in citations):
        author_patterns.append("et al.")
    if any(re.search(r"(?:&|\by\b|\be\b|\band\b)", c["autor"], re.IGNORECASE) for c in citations):
        author_patterns.append("múltiples autores")
    institutional = [c["autor"] for c in citations if len(c["autor"].split()) >= 3 and not re.search(r"(?:&|\by\b|\bet\s+al\.?)", c["autor"], re.IGNORECASE)]
    if institutional:
        author_patterns.append("posible autor institucional")
    checks.append(Check(
        "Citas", "Patrones de autoría", "ok",
        "Se reconocieron: " + ", ".join(dict.fromkeys(author_patterns)) + "." if author_patterns else
        "No se detectaron patrones complejos de autoría; se revisaron citas autor-año estándar.",
        "Confirma manualmente los autores institucionales o las citas complejas." if institutional else "",
    ))

    duplicate_keys = [key for key in {r["clave"] for r in references} if sum(r["clave"] == key for r in references) > 1]
    checks.append(Check(
        "Referencias", "Referencias duplicadas", "warning" if duplicate_keys else "ok",
        f"Se detectaron {len(duplicate_keys)} posible(s) duplicado(s)." if duplicate_keys else
        "No se detectaron referencias duplicadas por autor y año.",
        "Revisa y elimina las entradas repetidas." if duplicate_keys else "",
    ))

    ordered_keys = [r["surname"] for r in references if r["surname"]]
    is_ordered = ordered_keys == sorted(ordered_keys, key=lambda value: value.casefold())
    checks.append(Check(
        "Referencias", "Orden alfabético", "ok" if is_ordered else "warning",
        "Las referencias parecen estar en orden alfabético." if is_ordered else
        "Las referencias no parecen estar ordenadas alfabéticamente por autor.",
        "Ordena la lista de referencias alfabéticamente por el primer autor." if not is_ordered else "",
    ))

    missing_year = [r for r in references if not r["anio"]]
    checks.append(Check(
        "Referencias", "Año de publicación", "warning" if missing_year else "ok",
        f"{len(missing_year)} referencia(s) no muestran un año reconocible." if missing_year else
        "Todas las referencias tienen un año reconocible.",
        "Revisa el año de cada referencia." if missing_year else "",
    ))

    checks.append(Check(
        "Referencias", "Entradas bibliográficas", "ok" if references else "warning",
        f"Se detectaron {len(references)} entrada(s) en Referencias." if references else
        "La lista de referencias está vacía o no pudo identificarse.",
        "Incluye una entrada bibliográfica completa por cada fuente citada." if not references else "",
    ))

    incomplete = [r for r in references if r["elementos_faltantes"]]
    checks.append(Check(
        "Referencias", "Elementos de las referencias", "warning" if incomplete else "ok",
        (f"{len(incomplete)} referencia(s) parecen tener elementos faltantes: " +
         "; ".join(f"{r['autor']} ({', '.join(r['elementos_faltantes'])})" for r in incomplete[:4]) +
         ("." if len(incomplete) <= 4 else "; ….")) if incomplete else
        "Las referencias contienen autor, año y una obra identificable en el patrón revisado.",
        "Completa autor, año y título u obra; confirma además el tipo de fuente." if incomplete else "",
    ))

    malformed_doi = [r for r in references if re.search(r"\bdoi(?::|\.org)", r["texto"], re.IGNORECASE) and not r["tiene_doi"]]
    checks.append(Check(
        "Referencias", "DOI y URL", "warning" if malformed_doi else "ok",
        f"Se detectaron {len(malformed_doi)} DOI aparentemente incompleto(s) o inválido(s)." if malformed_doi else
        "No se detectaron DOI o URL con formato evidentemente inválido.",
        "Revisa que el DOI use https://doi.org/10.xxxx/xxxxx o que la URL esté completa." if malformed_doi else "",
    ))

    no_hanging = [r for r in references if not r["sangria_francesa"]]
    checks.append(Check(
        "Referencias", "Sangría francesa", "warning" if no_hanging else "ok",
        f"{len(no_hanging)} de {len(references)} referencia(s) no muestran sangría francesa de 1,27 cm." if no_hanging else
        "Todas las referencias muestran una sangría francesa aproximada de 1,27 cm.",
        "Selecciona las referencias y aplica sangría francesa de 1,27 cm en Párrafo." if no_hanging else "",
    ))

    journal_refs = [r for r in references if r["patron_publicacion"] == "revista"]
    journal_without_italics = [r for r in journal_refs if not r["tiene_cursiva"]]
    checks.append(Check(
        "Referencias", "Cursivas bibliográficas", "warning" if journal_without_italics else "ok",
        f"{len(journal_without_italics)} posible(s) referencia(s) de revista no contiene(n) texto en cursiva." if journal_without_italics else
        "Las referencias de revista reconocidas contienen texto en cursiva.",
        "Aplica cursiva al nombre de la revista y al volumen cuando corresponda." if journal_without_italics else "",
    ))

    journal_without_pages = [r for r in journal_refs if not r["tiene_rango_paginas"] and not r["tiene_doi"] and not r["tiene_url"]]
    checks.append(Check(
        "Referencias", "Datos de publicación", "warning" if journal_without_pages else "ok",
        f"{len(journal_without_pages)} referencia(s) con patrón de revista no muestran un rango de páginas ni DOI/URL." if journal_without_pages else
        "Las referencias de revista reconocidas incluyen páginas o DOI/URL.",
        "Revisa volumen, número y páginas; añade DOI o URL cuando corresponda." if journal_without_pages else "",
    ))

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
        "instructions_active": custom_rules.active,
        "instructions_summary": custom_rules.summary(),
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


def correct_document(data: bytes, filename: str = "", instructions_text: str = "") -> bytes:
    if not filename.lower().endswith(".docx"):
        raise ValueError("Solo se pueden corregir archivos .docx.")
    document = Document(io.BytesIO(data))
    custom_rules = parse_instructions(instructions_text)

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
        is_title = index == 0
        is_heading = text.lower() in HEADING_WORDS or (
            paragraph.style and paragraph.style.name and paragraph.style.name.lower().startswith("heading")
        )
        if is_title:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Inches(0)
            paragraph.paragraph_format.left_indent = Inches(0)
        elif is_heading:
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

    apply_custom_corrections(document, custom_rules)
    output = io.BytesIO()
    document.save(output)
    return output.getvalue()
