from flask import Flask, request, send_file
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_BREAK
import json
import os
import re
from unicodedata import normalize

app = Flask(__name__)
PLACEHOLDER = "{{COMPLETAR}}"

# --- HELPERS DE FORMATO Y BÚSQUEDA ---

def format_name_from_email(value: str) -> str:
    if not value: return ""
    local = value.split("@")[0]
    parts = [p.strip() for p in local.split(".") if p.strip()]
    return " ".join(p.capitalize() for p in parts)

def clean_text(t):
    """Normaliza texto: minúsculas, sin tildes y sin espacios extra."""
    if not t: return ""
    s = normalize('NFD', t.lower())
    s = re.sub(r'[\u0300-\u036f]', '', s)
    return s.strip()

def find_paragraph_index(doc, target_text):
    """Busca el índice de un párrafo ignorando tildes y mayúsculas."""
    target = clean_text(target_text)
    for i, p in enumerate(doc.paragraphs):
        if target in clean_text(p.text):
            return i
    return -1

def insert_paragraph_after(paragraph, text="", style=None):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if style: new_para.style = style
    if text: new_para.add_run(text)
    return new_para

def remove_paragraph(paragraph):
    p = paragraph._p
    p.getparent().remove(p)

def is_heading1(p):
    try:
        name = (p.style.name or "").lower()
        return "heading 1" in name or "título 1" in name
    except: return False

def set_labeled_line(doc, label, value):
    """Busca una línea que empiece con label y le concatena el valor."""
    for p in doc.paragraphs:
        if p.text.strip().startswith(label):
            p.text = f"{label} {value}"
            return

# --- LÓGICA DE SECCIONES ---

def rebuild_self_feedback_section(doc, auto):
    """Procesa la sección 'Autoevaluación'."""
    idx = find_paragraph_index(doc, "Autoevaluacion")
    if idx == -1: return

    # 1. Encontrar límites y limpiar contenido viejo
    end_idx = -1
    for j in range(idx + 1, len(doc.paragraphs)):
        if is_heading1(doc.paragraphs[j]):
            end_idx = j
            break
    if end_idx == -1: end_idx = len(doc.paragraphs)
    for p in reversed(doc.paragraphs[idx + 1:end_idx]):
        remove_paragraph(p)

    cursor = doc.paragraphs[idx]

    # 2. Manejo de data de n8n
    if isinstance(auto, list) and len(auto) > 0: auto = auto[0]
    
    if not isinstance(auto, dict):
        insert_paragraph_after(cursor, "No se registró autoevaluación.")
        return

    pos = auto.get("positivos", "").strip()
    mej = auto.get("mejorar", "").strip()
    mas = auto.get("algo_mas", "").strip()

    if not any([pos, mej, mas]):
        insert_paragraph_after(cursor, "No se registró autoevaluación.")
    else:
        mapping = [
            ("1. Aspectos positivos", pos),
            ("2. Aspectos a mejorar", mej),
            ("3. Algo más que quieras compartir.", mas)
        ]
        for label, content in mapping:
            p_l = insert_paragraph_after(cursor, "")
            p_l.add_run(label).bold = True
            txt = content if content else PLACEHOLDER
            p_t = insert_paragraph_after(p_l, txt)
            for r in p_t.runs: r.font.size = Pt(11)
            cursor = p_t

    # Salto de página
    p_br = insert_paragraph_after(cursor, "")
    p_br.add_run().add_break(WD_BREAK.PAGE)

def rebuild_feedback_section(doc, evs):
    """Procesa la sección 'Feedback Recibido' usando la lista de evaluaciones."""
    idx = find_paragraph_index(doc, "Feedback Recibido")
    if idx == -1: return

    # Limpiar hasta el final o siguiente Título 1
    end_idx = -1
    for j in range(idx + 1, len(doc.paragraphs)):
        if is_heading1(doc.paragraphs[j]):
            end_idx = j
            break
    if end_idx == -1: end_idx = len(doc.paragraphs)
    for p in reversed(doc.paragraphs[idx + 1:end_idx]):
        remove_paragraph(p)

    cursor = doc.paragraphs[idx]

    if not evs or not isinstance(evs, list):
        insert_paragraph_after(cursor, "No se recibieron evaluaciones de terceros.")
        return

    # Iterar sobre cada evaluación (naranjas, limones, etc.)
    for i, ev in enumerate(evs):
        # Separador visual entre evaluadores
        if i > 0:
            cursor = insert_paragraph_after(cursor, "-" * 30)

        p_name = insert_paragraph_after(cursor, "")
        nombre_evaluador = format_name_from_email(ev.get("evaluador"))
        p_name.add_run(f"Evaluador: {nombre_evaluador}").bold = True
        cursor = p_name

        mapping = [
            ("Aspectos positivos:", ev.get("positivos")),
            ("Aspectos a mejorar:", ev.get("mejorar")),
            ("Algo más que quieras compartir:", ev.get("algo_mas"))
        ]

        for label, content in mapping:
            p_l = insert_paragraph_after(cursor, label)
            p_l.runs[0].bold = True
            txt = str(content).strip() if content else PLACEHOLDER
            p_t = insert_paragraph_after(p_l, txt)
            for r in p_t.runs: r.font.size = Pt(11)
            cursor = p_t

# --- RUTAS ---

@app.post("/generate")
def generate():
    try:
        data = request.json or {}
        print("FULL BODY:", json.dumps(data, ensure_ascii=False, indent=2))
        print("TOP KEYS:", list(data.keys()))

        # Datos básicos
        raw_evaluado = data.get("evaluado", "").strip()
        evaluado = format_name_from_email(raw_evaluado) or PLACEHOLDER
        mes_ano = data.get("mes_ano", "").strip() or PLACEHOLDER
        
        # Objetos n8n
        auto_data = data.get("autoevaluacion")
        print("AUTOEVAL TYPE:", type(auto_data))
        print("AUTOEVAL KEYS:", list(auto_data.keys()) if isinstance(auto_data, dict) else auto_data)
        print("AUTOEVAL RAW:", json.dumps(auto_data, ensure_ascii=False, indent=2))

        evs_data = data.get("evaluaciones", [])

        # Cargar Template
        template_path = "template.docx"
        if not os.path.exists(template_path):
            return {"error": "No se encontró template.docx"}, 404
            
        doc = Document(template_path)

        # 1. Llenar encabezado
        set_labeled_line(doc, "Nombre:", evaluado)
        set_labeled_line(doc, "Periodo evaluado:", mes_ano)

        # 2. Reconstruir secciones
        rebuild_self_feedback_section(doc, auto_data)
        rebuild_feedback_section(doc, evs_data)

        # 3. Enviar archivo
        bio = BytesIO()
        doc.save(bio)
        bio.seek(0)
        
        filename = f"Performance Review - {evaluado}.docx"
        return send_file(bio, as_attachment=True, download_name=filename, 
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
