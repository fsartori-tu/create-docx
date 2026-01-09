from flask import Flask, request, send_file, abort
from docx import Document
from docx.shared import Pt
import tempfile
import os
import re

app = Flask(__name__)

# -------- helpers --------

def format_name_from_email(email: str) -> str:
    """
    nombre.apellido@dominio -> Nombre Apellido
    """
    if not email or "@" not in email:
        return email

    name_part = email.split("@")[0]
    parts = name_part.split(".")
    parts = [p.capitalize() for p in parts if p]
    return " ".join(parts)


def safe_text(value):
    return value if value and str(value).strip() else "{{COMPLETAR}}"


# -------- endpoint --------

@app.route("/generate", methods=["POST"])
def generate_docx():
    try:
        data = request.json

        evaluado_raw = data.get("evaluado", "")
        evaluado = format_name_from_email(evaluado_raw)
        evaluaciones = data.get("evaluaciones", [])
        mes_ano = data.get("mes_ano", "")

        # abrir template
        doc = Document("template.docx")

        # recorrer evaluadores
        for ev in evaluaciones:
            evaluador_raw = ev.get("evaluador", "")
            evaluador = format_name_from_email(evaluador_raw)

            positivos = safe_text(ev.get("aspectos_positivos"))
            mejorar = safe_text(ev.get("aspectos_a_mejorar"))
            extra = safe_text(ev.get("algo_mas"))

            # nombre evaluador (Heading 2)
            doc.add_heading(evaluador, level=2)

            # --- Aspectos positivos ---
            p = doc.add_paragraph(style="List Paragraph")
            run = p.add_run("Aspectos positivos")
            run.bold = True

            r = doc.add_paragraph()
            run = r.add_run(positivos)
            run.font.size = Pt(11)

            # --- Aspectos a mejorar ---
            p = doc.add_paragraph(style="List Paragraph")
            run = p.add_run("Aspectos a mejorar")
            run.bold = True

            r = doc.add_paragraph()
            run = r.add_run(mejorar)
            run.font.size = Pt(11)

            # --- Algo más ---
            p = doc.add_paragraph(style="List Paragraph")
            run = p.add_run("Algo más que quieras compartir.")
            run.bold = True

            r = doc.add_paragraph()
            run = r.add_run(extra)
            run.font.size = Pt(11)

        # guardar archivo temporal
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(tmp.name)

        filename = f"Performance Review – {evaluado} – {mes_ano}.docx"

        return send_file(
            tmp.name,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("ERROR:", e)
        abort(500)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 3000)))

