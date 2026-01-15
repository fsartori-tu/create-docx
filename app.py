from flask import Flask, request, send_file
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
import json
import os

app = Flask(__name__)
PLACEHOLDER = "{{COMPLETAR}}"

# -------------------------
# Helpers
# -------------------------

def format_name_from_email(value: str) -> str:
    if not value: return ""
    value = value.strip()
    if "@" not in value: return value
    local = value.split("@")[0]
    parts = [p.strip() for p in local.split(".") if p.strip()]
    return " ".join(p.capitalize() for p in parts)

def safe_text(value: str) -> str:
    v = (value or "").strip()
    return v if v else PLACEHOLDER

def insert_paragraph_after(paragraph: Paragraph, text: str = "", style: str | None = None) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if style: new_para.style = style
    if text: new_para.add_run(text)
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

def is_heading1(p: Paragraph) -> bool:
    try:
        return (p.style.name or "").strip().lower() == "heading 1"
    except: return False

def set_labeled_line(doc: Document, label: str, value: str) -> bool:
    label_norm = (label or "").strip()
    for p in doc.paragraphs:
        txt = (p.text or "").strip()
        if txt.startswith(label_norm):
            p.text = f"{label_norm} {value if value else PLACEHOLDER}"
            return True
    return False

def set_run_font_size(paragraph: Paragraph, size_pt: int) -> None:
    for r in paragraph.runs:
        r.font.size = Pt(size_pt)

# -------------------------
# Secciones
# -------------------------

def rebuild_self_feedback_section(doc: Document, auto: dict | None) -> None:
    start_idx = find_paragraph_index(doc, "Autoevaluación")
    if start_idx == -1: return

    end_idx = -1
    for j in range(start_idx + 1, len(doc.paragraphs)):
        if is_heading1(doc.paragraphs[j]):
            end_idx = j
            break
    if end_idx == -1: end_idx = len(doc.paragraphs)

    for p in reversed(doc.paragraphs[start_idx + 1:end_idx]):
        remove_paragraph(p)

    cursor = doc.paragraphs[start_idx]
    
    # Verificación más flexible de si el objeto autoevaluación tiene contenido real
    has_content = auto and any(auto.get(k) for k in ["positivos", "mejorar", "algo_mas"])

    if not has_content:
        p = insert_paragraph_after(cursor, "No se registró autoevaluación.", style="Normal")
        set_run_font_size(p, 11)
        return

    mapping = [
        ("Aspectos positivos", auto.get("positivos")),
        ("Aspectos a mejorar", auto.get("mejorar")),
        ("Algo más que quieras compartir.", auto.get("algo_mas"))
    ]

    for label, content in mapping:
        p_lab = insert_paragraph_after(cursor, "", style="List Paragraph")
        r = p_lab.add_run(label)
        r.bold = True
        cursor = p_lab
        
        p_txt = insert_paragraph_after(cursor, safe_text(content), style="Normal")
        set_run_font_size(p_txt, 11)
        cursor = p_txt

def rebuild_feedback_section(doc: Document, evaluaciones: list) -> None:
    start_idx = find_paragraph_index(doc, "Feedback recibido")
    if start_idx == -1: return

    end_idx = -1
    for j in range(start_idx + 1, len(doc.paragraphs)):
        if is_heading1(doc.paragraphs[j]):
            end_idx = j
            break
    if end_idx == -1: end_idx = len(doc.paragraphs)

    for p in reversed(doc.paragraphs[start_idx + 1:end_idx]):
        remove_paragraph(p)

    cursor = doc.paragraphs[start_idx]
    if not evaluaciones:
        p = insert_paragraph_after(cursor, PLACEHOLDER, style="Normal")
        set_run_font_size(p, 11)
        return

    for idx, ev in enumerate(evaluaciones):
        evaluador = format_name_from_email(ev.get("evaluador", "")) or PLACEHOLDER
        p_eval = insert_paragraph_after(cursor, evaluador, style="Heading 2")
        cursor = p_eval

        for lab, key in [("Aspectos positivos", "positivos"), 
                         ("Aspectos a mejorar", "mejorar"), 
                         ("Algo más...", "algo_mas")]:
            p_b = insert_paragraph_after(cursor, "", style="List Paragraph")
            p_b.add_run(lab).bold = True
            cursor = p_b
            p_r = insert_paragraph_after(cursor, safe_text(ev.get(key)), style="Normal")
            set_run_font_size(p_r, 11)
            cursor = p_r
        
        if idx < len(evaluaciones) - 1:
            cursor = insert_paragraph_after(cursor, "", style="Normal")

# -------------------------
# Ruta Principal
# -------------------------

@app.post("/generate")
def generate():
    data = request.json or {}
    
    raw_evaluado = (data.get("evaluado") or "").strip()
    evaluado = format_name_from_email(raw_evaluado) or PLACEHOLDER
    mes_ano = (data.get("mes_ano") or "").strip() or PLACEHOLDER
    
    # Autoevaluacion (dict)
    auto = data.get("autoevaluacion")
    if isinstance(auto, str) and auto.strip():
        try: auto = json.loads(auto)
        except: auto = None
        
    # Evaluaciones (list)
    evs = data.get("evaluaciones") or []
    if isinstance(evs, str):
        try: evs = json.loads(evs)
        except: evs = []

    template_path = os.environ.get("TEMPLATE_PATH", "template.docx")
    doc = Document(template_path) if os.path.exists(template_path) else Document()

    set_labeled_line(doc, "Nombre:", evaluado)
    set_labeled_line(doc, "Periodo evaluado:", mes_ano)
    
    rebuild_self_feedback_section(doc, auto)
    rebuild_feedback_section(doc, evs)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    # RESTAURADO: mes_ano en el nombre del archivo
    filename = f"Performance Review – {evaluado} – {mes_ano}.docx"
    
    return send_file(bio, as_attachment=True, download_name=filename, 
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    app.run(port=5000)
