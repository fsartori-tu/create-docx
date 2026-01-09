from flask import Flask, request, send_file
from io import BytesIO
from datetime import datetime
from docx import Document

app = Flask(__name__)

def add_label_and_text(doc, label, text):
    p = doc.add_paragraph()
    run = p.add_run(label)
    run.bold = True
    doc.add_paragraph(text or "")

@app.post("/generate")
def generate():
    data = request.json or {}

    evaluado = (data.get("evaluado") or "").strip()
    mes_ano = (data.get("mes_ano") or "").strip()
    evaluaciones = data.get("evaluaciones") or []

    doc = Document()
    doc.add_heading("Documento de Feedback – Evaluación de pares", level=1)
    doc.add_paragraph(f"Evaluado: {evaluado}")
    doc.add_paragraph(f"Fecha de generación: {datetime.now().strftime('%Y-%m-%d')}")
    if mes_ano:
        doc.add_paragraph(f"Mes/Año: {mes_ano}")
    doc.add_page_break()

    for idx, ev in enumerate(evaluaciones):
        doc.add_heading(f"Evaluador: {ev.get('evaluador','')}", level=2)
        add_label_and_text(doc, "1. Aspectos positivos", ev.get("positivos",""))
        add_label_and_text(doc, "2. Aspectos a mejorar", ev.get("mejorar",""))
        add_label_and_text(doc, "3. Algo más que quieras compartir.", ev.get("algo_mas",""))
        if idx < len(evaluaciones) - 1:
            doc.add_page_break()

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)

    filename = (
        f"Performance Review – {evaluado} – {mes_ano}.docx"
        if mes_ano else
        f"Performance Review – {evaluado}.docx"
    )

    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
