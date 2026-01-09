from flask import Flask, request, send_file
from io import BytesIO
from datetime import datetime
from docx import Document
from docx.shared import Pt
import json
import os

from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

app = Flask(__name__)

PLACEHOLDER = "{{COMPLETAR}}"


# -------------------------
# Helpers: names
# -------------------------

def format_name_from_email(value: str) -> str:
    """
    Convierte 'nombre.apellido@empresa.com' -> 'Nombre Apellido'
    Si no es email, devuelve tal cual (strip).
    """
    if not value:
        return ""
    value = value.strip()
    if "@" not in value:
        return value

    local = value.split("@")[0]
    parts = [p.strip() for p in local.split(".") if p.strip()]
    return " ".join(p.capitalize() for p in parts)


def safe_text(value: str) -> str:
    v = (value or "").strip()
    return v if v else PLACEHOLDER


# -------------------------
# Helpers: docx manipulation
# -------------------------

def insert_paragraph_after(paragraph: Paragraph, text: str = "", style: str | None = None) -> Paragraph:
    """
    Inserta un nuevo párrafo inmediatamente después de 'paragraph'
    """
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if style:
        new_para.style = style
    if text:
        new_para.add_run(text)
    return new_para


def remove_paragraph(paragraph: Paragraph) -> None:
    p = paragraph._p
    p.getparent().remove(p)


def find_paragraph_index(doc: Document, target_text: str) -> int:
    t = (target_text or "").strip().lower()
    for i, p in enumerate(doc.paragraphs):
        if (p.text or "").strip().lower() == t:
            return i
    return -1


def style_name(p: Paragraph) -> str:
    try:
        return p.style.name or ""
    except Exception:
        return ""


def is_heading1(p: Paragraph) -> bool:
    return style_name(p).strip().lower() == "heading 1"


def set_labeled_line(doc: Document, label: str, value: str) -> bool:
    """
    Busca un párrafo que arranque con 'label' (ej: 'Nombre:') y lo reemplaza por 'label value'
    """
    label_norm = (label or "").strip()
    for p in doc.paragraphs:
        txt = (p.text or "").strip()
        if txt.startswith(label_norm):
            p.text = f"{label_norm} {value if value else PLACEHOLDER}"
            return True
    return False


def set_run_font_size(paragraph: Paragraph, size_pt: int) -> None:
    """
    Fuerza tamaño de fuente en todos los runs del párrafo.
    """
    for r in paragraph.runs:
        r.font.size = Pt(size_pt)


def rebuild_feedback_section(doc: Document, evaluaciones: list[dict]) -> None:
    """
    Reemplaza TODO el contenido entre:
      Heading 1: 'Feedback recibido'
    y el próximo Heading 1

    Formato por evaluador:
      Heading 2: Nombre
      - Aspectos positivos   (negrita)
        respuesta            (11pt)
      - Aspectos a mejorar   (negrita)
        respuesta            (11pt)
      - Algo más...          (negrita)
        respuesta            (11pt)
    """
    start_idx = find_paragraph_index(doc, "Feedback recibido")
    if start_idx == -1:
        return

    end_idx = -1
    for j in range(start_idx + 1, len(doc.paragraphs)):
        if is_heading1(doc.paragraphs[j]):
            end_idx = j
            break
    if end_idx == -1:
        end_idx = len(doc.paragraphs)

    feedback_heading = doc.paragraphs[start_idx]

    # borrar todo entre Feedback recibido y el siguiente Heading 1
    to_remove = doc.paragraphs[start_idx + 1:end_idx]
    for p in reversed(to_remove):
        remove_paragraph(p)

    cursor = feedback_heading

    if not evaluaciones:
        p = insert_paragraph_after(cursor, PLACEHOLDER, style="Normal")
        set_run_font_size(p, 11)
        return

    for idx, ev in enumerate(evaluaciones):
        evaluador = format_name_from_email((ev.get("evaluador") or "").strip()) or PLACEHOLDER

        positivos = safe_text(ev.get("positivos"))
        mejorar = safe_text(ev.get("mejorar"))
        algo_mas = safe_text(ev.get("algo_mas"))

        # Heading 2 evaluador
        p_eval = insert_paragraph_after(cursor, evaluador, style="Heading 2")
        cursor = p_eval

        # Bullet 1 (bold)
        p_b1 = insert_paragraph_after(cursor, "", style="List Paragraph")
        r = p_b1.add_run("Aspectos positivos")
        r.bold = True
        cursor = p_b1

        p_r1 = insert_paragraph_after(cursor, positivos, style="Normal")
        set_run_font_size(p_r1, 11)
        cursor = p_r1

        # Bullet 2 (bold)
        p_b2 = insert_paragraph_after(cursor, "", style="List Paragraph")
        r = p_b2.add_run("Aspectos a mejorar")
        r.bold = True
        cursor = p_b2

        p_r2 = insert_paragraph_after(cursor, mejorar, style="Normal")
        set_run_font_size(p_r2, 11)
        cursor = p_r2

        # Bullet 3 (bold)
        p_b3 = insert_paragraph_after(cursor, "", style="List Paragraph")
        r = p_b3.add_run("Algo más que quieras compartir.")
        r.bold = True
        cursor = p_b3

        p_r3 = insert_paragraph_after(cursor, algo_mas, style="Normal")
        set_run_font_size(p_r3, 11)
        cursor = p_r3

        # espacio entre evaluadores
        if idx < len(evaluaciones) - 1:
            spacer = insert_paragraph_after(cursor, "", style="Normal")
            cursor = spacer


# -------------------------
# Routes
# -------------------------

@app.get("/")
def health():
    return "OK", 200


@app.post("/generate")
def generate():
    data = request.json or {}

    raw_evaluado = (data.get("evaluado") or "").strip()
    evaluado = format_name_from_email(raw_evaluado) or PLACEHOLDER

    mes_ano = (data.get("mes_ano") or "").strip() or PLACEHOLDER

    evaluaciones = data.get("evaluaciones") or []

    # soporte si n8n manda evaluaciones como string
    if isinstance(evaluaciones, str):
        try:
            evaluaciones = json.loads(evaluaciones)
        except Exception:
            evaluaciones = []

    if not isinstance(evaluaciones, list):
        evaluaciones = []

    # template
    template_path = os.environ.get("TEMPLATE_PATH", "template.docx")
    doc = Document(template_path) if os.path.exists(template_path) else Document()

    # Info general
    set_labeled_line(doc, "Nombre:", evaluado)
    set_labeled_line(doc, "Periodo evaluado:", mes_ano)

    # Feedback recibido (sección)
    rebuild_feedback_section(doc, evaluaciones)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)

    filename = f"Performance Review – {evaluado} – {mes_ano}.docx"
    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
