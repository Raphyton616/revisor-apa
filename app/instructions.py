"""Reglas adicionales proporcionadas por el docente para una revisión puntual."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from docx.shared import Inches, Pt


@dataclass
class WorkInstructions:
    raw_text: str = ""
    font_name: str | None = None
    font_size: float | None = None
    line_spacing: float | None = None
    margin_cm: float | None = None
    reference_heading: str | None = None
    first_line_indent_cm: float | None = None
    required_sections: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def active(self) -> bool:
        return bool(self.raw_text.strip())

    def summary(self) -> list[str]:
        values = []
        if self.font_name and self.font_size:
            values.append(f"Fuente: {self.font_name} {self.font_size:g} pt")
        elif self.font_name:
            values.append(f"Fuente: {self.font_name}")
        if self.line_spacing:
            values.append(f"Interlineado: {self.line_spacing:g}")
        if self.margin_cm:
            values.append(f"Márgenes: {self.margin_cm:g} cm")
        if self.reference_heading:
            values.append(f"Título final: {self.reference_heading}")
        if self.first_line_indent_cm is not None:
            values.append(f"Sangría: {self.first_line_indent_cm:g} cm")
        if self.required_sections:
            values.append("Secciones: " + ", ".join(self.required_sections))
        return values


def _number(value: str) -> float:
    return float(value.replace(",", "."))


def parse_instructions(text: str | None) -> WorkInstructions:
    raw = (text or "").strip()
    result = WorkInstructions(raw_text=raw)
    if not raw:
        return result

    lower = raw.lower()

    font_patterns = [
        (r"times\s+new\s+roman", "Times New Roman"),
        (r"arial", "Arial"),
        (r"calibri", "Calibri"),
        (r"georgia", "Georgia"),
        (r"lucida\s+sans\s+unicode", "Lucida Sans Unicode"),
    ]
    for pattern, name in font_patterns:
        if re.search(pattern, lower):
            result.font_name = name
            break

    size_match = re.search(r"(?:fuente|letra|tamaño|tamano).{0,35}?(\d{1,2}(?:[,.]\d+)?)\s*(?:puntos?|pt)\b", lower)
    if not size_match:
        size_match = re.search(r"(?:times\s+new\s+roman|arial|calibri|georgia|lucida\s+sans\s+unicode)\s+(\d{1,2}(?:[,.]\d+)?)\b", lower)
    if size_match:
        result.font_size = _number(size_match.group(1))

    spacing_match = re.search(r"(?:interlineado|espaciado).{0,35}?(\d(?:[,.]\d+)?)", lower)
    if spacing_match:
        result.line_spacing = _number(spacing_match.group(1))
    elif "interlineado doble" in lower or "doble espacio" in lower:
        result.line_spacing = 2.0
    elif "interlineado sencillo" in lower or "interlineado simple" in lower:
        result.line_spacing = 1.0

    margin_match = re.search(r"márgenes?.{0,35}?(\d+(?:[,.]\d+)?)\s*cm", lower)
    if margin_match:
        result.margin_cm = _number(margin_match.group(1))

    indent_match = re.search(r"sangr[ií]a(?: de primera l[ií]nea)?.{0,35}?(\d+(?:[,.]\d+)?)\s*cm", lower)
    if indent_match:
        result.first_line_indent_cm = _number(indent_match.group(1))
    elif "sin sangría" in lower or "sin sangria" in lower:
        result.first_line_indent_cm = 0.0

    if re.search(r"(?:usar|use|llamar|titular).{0,30}(?:bibliografía|bibliografia)", lower):
        result.reference_heading = "Bibliografía"
    elif re.search(r"(?:usar|use|llamar|titular).{0,30}referencias", lower):
        result.reference_heading = "Referencias"

    for section in ("portada", "resumen", "introducción", "introduccion", "desarrollo", "método", "metodo", "resultados", "discusión", "discusion", "conclusión", "conclusion", "referencias", "bibliografía", "bibliografia"):
        if re.search(rf"\b(?:incluir|incluya|debe tener|contener|llevar|con)\b[^.\n]*\b{re.escape(section)}\b", lower):
            canonical = {"introduccion": "Introducción", "conclusion": "Conclusión", "metodo": "Método", "discusion": "Discusión", "bibliografia": "Bibliografía"}.get(section, section.title())
            if canonical not in result.required_sections:
                result.required_sections.append(canonical)

    if "no se puede determinar" in lower or "según lo visto en clase" in lower:
        result.notes.append("Hay una instrucción ambigua que requiere revisión manual.")
    return result


def check_custom_instructions(document, rules: WorkInstructions) -> list[dict]:
    if not rules.active:
        return []
    from app.apa_logic import length_cm, non_empty_paragraphs

    checks = []
    sections = document.sections
    if rules.margin_cm is not None and sections:
        section = sections[0]
        values = [length_cm(section.top_margin), length_cm(section.bottom_margin), length_cm(section.left_margin), length_cm(section.right_margin)]
        ok = all(value is not None and abs(value - rules.margin_cm) <= 0.1 for value in values)
        checks.append({"category": "Instrucciones", "title": "Márgenes del trabajo", "status": "ok" if ok else "error", "detail": f"Los márgenes coinciden con el requisito de {rules.margin_cm:g} cm." if ok else f"El requisito indica márgenes de {rules.margin_cm:g} cm.", "recommendation": "Ajusta los cuatro márgenes según las instrucciones del trabajo." if not ok else ""})

    paragraphs = non_empty_paragraphs(document)
    if rules.font_name or rules.font_size:
        wrong = []
        for paragraph in paragraphs:
            for run in paragraph.runs:
                if not run.text.strip():
                    continue
                if rules.font_name and (run.font.name or "") != rules.font_name:
                    wrong.append(True)
                elif rules.font_size and run.font.size and abs(run.font.size.pt - rules.font_size) > 0.1:
                    wrong.append(True)
        ok = not wrong
        expected = rules.font_name or "la fuente indicada"
        if rules.font_size:
            expected += f" {rules.font_size:g} pt"
        checks.append({"category": "Instrucciones", "title": "Tipografía del trabajo", "status": "ok" if ok else "warning", "detail": f"La tipografía coincide con {expected}." if ok else f"El requisito específico indica {expected}.", "recommendation": "Aplica la fuente y tamaño indicados por el docente." if not ok else ""})

    if rules.line_spacing is not None:
        bad = [p for p in paragraphs if p.paragraph_format.line_spacing is not None and (not isinstance(p.paragraph_format.line_spacing, (int, float)) or abs(float(p.paragraph_format.line_spacing) - rules.line_spacing) > 0.05)]
        checks.append({"category": "Instrucciones", "title": "Interlineado del trabajo", "status": "ok" if not bad else "warning", "detail": f"El interlineado coincide con {rules.line_spacing:g}." if not bad else f"El requisito específico indica interlineado {rules.line_spacing:g}.", "recommendation": "Aplica el interlineado indicado por el docente." if bad else ""})

    if rules.required_sections:
        text = " ".join(p.text.strip().lower() for p in paragraphs)
        missing = [section for section in rules.required_sections if section.lower() not in text]
        checks.append({"category": "Instrucciones", "title": "Secciones requeridas", "status": "ok" if not missing else "warning", "detail": "Se encontraron todas las secciones requeridas." if not missing else "Faltan: " + ", ".join(missing) + ".", "recommendation": "Añade las secciones solicitadas por el docente." if missing else ""})

    if rules.notes:
        checks.append({"category": "Instrucciones", "title": "Revisión manual", "status": "warning", "detail": " ".join(rules.notes), "recommendation": "Confirma esta regla directamente con el docente."})
    return checks


def apply_custom_corrections(document, rules: WorkInstructions) -> None:
    if not rules.active:
        return
    for section in document.sections:
        if rules.margin_cm is not None:
            margin = Inches(rules.margin_cm / 2.54)
            section.top_margin = margin
            section.bottom_margin = margin
            section.left_margin = margin
            section.right_margin = margin
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if rules.font_name:
                run.font.name = rules.font_name
            if rules.font_size:
                run.font.size = Pt(rules.font_size)
        if rules.line_spacing is not None:
            paragraph.paragraph_format.line_spacing = rules.line_spacing
        if rules.first_line_indent_cm is not None:
            paragraph.paragraph_format.first_line_indent = Inches(rules.first_line_indent_cm / 2.54)
    if rules.reference_heading:
        for paragraph in document.paragraphs:
            if paragraph.text.strip().lower() in {"referencias", "bibliografía", "bibliografia", "references"}:
                paragraph.text = rules.reference_heading
                break
