from flask import Flask, request, send_file
from io import BytesIO
from datetime import datetime
from docx import Document
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


# -------------------------
# Helpers: docx manipulation
# -------------------------

def insert_paragraph_after(paragraph: Paragraph, text: str = "", style: str | None = None) -> Paragraph:
    """
    Inserta un nuevo párrafo inmediatamente después de 'paragraph',
    preservando el orden del documento (clave para mantener el template).
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
    """
    Elimina un párrafo del documento.
    """
    p = paragraph._p
    p.getparent().remove(p)


def find_paragraph_index(doc: Document, target_text: str) -> int:
    """
    Devuelve el índice del párrafo cuyo texto coincide (case-insensitive, trimmed).
    Si no lo encuentra, devuelve -1.
    """
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
    Busca un párrafo que arranque con 'label' (ej: 'Nombre:') y lo reemplaza por 'label value'.
    Mantiene el estilo del párrafo.
    """
    label_norm = (label or "").strip()
    for p in doc.paragraphs:
        txt = (p.text or "").strip()
        if txt.startswith(label_norm):
            # mantener estilo, reemplazar texto
            p.text = f"{label_norm} {value if value else PLACEHOLDER}"
            return True
    return False


def rebuild_feedback_section(doc: Document, evaluaciones: list[dict]) -> None:
    """
    Reemplaza el contenido entre:
    Heading 1: 'Feedback recibido'
    y el próximo Heading 1 (ej: 'Resumen')

    Genera bloques por evaluador como:
      Heading 2: {Nombre}
      - Aspectos positivos
        respuesta
      - Aspectos a mejorar
        respuesta
      - Algo más que quieras compartir.
        respuesta
    """
    start_idx = find_paragraph_index(doc, "Feedback recibido")
    if start_idx == -1:
        # Si el template no tiene ese título exacto, no hacemos nada para no romper.
        return

    # Encontrar fin: siguiente Heading 1 después de start
    end_idx = -1
    for j in range(start_idx + 1, len(doc.paragraphs)):
        if is_heading1(doc.paragraphs[j]):
            end_idx = j
            break
    if end_idx == -1:
        end_idx = len(doc.paragraphs)

    # Guardamos referencia al párrafo "Feedback recibido"
    feedback_heading = doc.paragraphs[start_idx]

    # Eliminar TODO lo que hay entre feedback_heading y el siguiente Heading 1
    # (removemos desde el final hacia atrás para no romper indices)
    to_remove = doc.paragraphs[start_idx + 1:end_idx]
    for p in reversed(to_remove):
        remove_paragraph(p)

    # Ahora insertamos contenido nuevo justo después del heading
    cursor = feedback_heading

    if not evaluaciones:
        # Si no hay evaluaciones, dejamos placeholder mínimo
        p = insert_paragraph_after(cursor, PLACEHOLDER, style="Normal")
        cursor = p
        return

    for idx, ev in enumerate(evaluaciones):
        evaluador = format_name_from_email((ev.get("evaluador") or "").strip()) or PLACEHOLDER

        positivos = (ev.get("positivos") or "").strip() or PLACEHOLDER
        mejorar = (ev.get("mejorar") or "").strip() or PLACEHOLDER
        algo_mas = (ev.get("algo_mas") or "").strip() or PLACEHOLDER

        # Heading 2 con nombre del evaluador (como en tu template)
        p_eval = insert_paragraph_after(cursor, evaluador, style="Heading 2")
        cursor = p_eval

        # Bullets + respuesta debajo (manteniendo estilo List Paragraph)
        p_b1 = insert_paragraph_after(cursor, "Aspectos positivos", style="List Paragraph")
        cursor = p_b1
        p_r1 = insert_paragraph_after(cursor, positivos, style="Normal")
        cursor = p_r1

        p_b2 = insert_paragraph_after(cursor, "Aspectos a mejorar", style="List Paragraph")
        cursor = p_b2
        p_r2 = insert_paragraph_after(cursor, mejorar, style="Normal")
        cursor = p_r2

        p_b3 = insert_paragraph_after(cursor, "Algo más que quieras compartir.", style="List Paragraph")
        cursor = p_b3
        p_r3 = insert_paragraph_after(cursor, algo_mas, style="Normal")
        cursor = p_r3

        # Espacio entre evaluadores (opcional, prolijo)
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

    # -------- Inputs --------
    raw_evaluado = (data.get("evaluado") or "").strip()
    evaluado = format_name_from_email(raw_evaluado) or PLACEHOLDER

    mes_ano = (data.get("mes_ano") or "").strip() or PLACEHOLDER

    evaluaciones = data.get("evaluaciones") or []

    # Soporte por si n8n manda evaluaciones como string
    if isinstance(evaluaciones, str):
        try:
            evaluaciones = json.loads(evaluaciones)
        except Exception:
            evaluaciones = []

    if not isinstance(evaluaciones, list):
        evaluaciones = []

    # -------- Load template --------
    template_path = os.environ.get("TEMPLATE_PATH", "template.docx")
    if os.path.exists(template_path):
        doc = Document(template_path)
    else:
        # fallback: documento vacío (pierde estilos del template)
        doc = Document()

    # -------- Fill general info (sin romper si no encuentra labels) --------
    # Estos labels existen en tu template en "Información general"
    set_labeled_line(doc, "Nombre:", evaluado)
    set_labeled_line(doc, "Periodo evaluado:", mes_ano)

    # Si querés más adelante: Rol, Seniority, Links, etc:
    # set_labeled_line(doc, "Rol:", data.get("rol",""))
    # set_labeled_line(doc, "Seniority actual:", data.get("seniority_actual",""))
    # set_labeled_line(doc, "Autoevaluación: Link", data.get("autoeval_link",""))

    # -------- Fill feedback section (títulos + listado, sin tablas) --------
    rebuild_feedback_section(doc, evaluaciones)

    # -------- Output --------
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

