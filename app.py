from flask import Flask, request, send_file
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re

app = Flask(__name__)

def parse_and_build_docx(text):
    doc = Document()

    # Estilos generales
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # Título principal
    title = doc.add_heading('Business Case — xNova International', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    doc.add_paragraph()

    # Procesar el texto: detectar secciones numeradas y contenido
    lines = text.split('\n')
    
    for line in lines:
        line = line.rstrip()
        if not line:
            doc.add_paragraph()
            continue

        # Detectar headers de sección (1. TITULO, 2. TITULO, etc.)
        section_match = re.match(r'^(\d+)\.\s+([A-ZÁÉÍÓÚÑ\s]+)$', line.strip())
        if section_match:
            h = doc.add_heading(line.strip(), level=1)
            for run in h.runs:
                run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
            continue

        # Detectar sub-headers con ##
        if line.startswith('## '):
            h = doc.add_heading(line[3:].strip(), level=2)
            continue

        # Detectar headers markdown
        if line.startswith('# '):
            h = doc.add_heading(line[2:].strip(), level=1)
            for run in h.runs:
                run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
            continue

        # Detectar bullets
        if line.strip().startswith('- ') or line.strip().startswith('• '):
            content = line.strip()[2:]
            p = doc.add_paragraph(style='List Bullet')
            # Procesar negritas dentro del bullet
            _add_formatted_run(p, content)
            continue

        # Detectar líneas con **negrita**
        if '**' in line:
            p = doc.add_paragraph()
            _add_formatted_run(p, line.strip())
            continue

        # Párrafo normal
        p = doc.add_paragraph(line.strip())

    return doc

def _add_formatted_run(paragraph, text):
    """Añade texto al párrafo procesando **negritas**"""
    parts = re.split(r'(\*\*[^*]+\*\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            paragraph.add_run(part)

@app.route('/health', methods=['GET'])
def health():
    return {'status': 'ok'}, 200

@app.route('/generate', methods=['POST'])
def generate():
    # Aceptar tanto JSON como form-urlencoded
    if request.content_type and 'application/json' in request.content_type:
        data = request.get_json(force=True, silent=True) or {}
    else:
        data = request.form.to_dict()
        if not data:
            data = request.get_json(force=True, silent=True) or {}

    text = data.get('text', '')
    filename = data.get('filename', 'BusinessCase_xNova.docx')

    if not text:
        return {'error': 'No text provided'}, 400

    doc = parse_and_build_docx(text)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
