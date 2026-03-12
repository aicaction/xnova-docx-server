from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io, os, re

app = Flask(__name__)

def parse_and_build_docx(text):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    title = doc.add_heading('Business Case - xNova International', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.color.rgb = RGBColor(0x01, 0x18, 0xD8)
    doc.add_paragraph()
    for line in text.strip().split('\n'):
        line = line.strip()
        if not line:
            doc.add_paragraph()
        elif re.match(r'^\d+\.\s+[A-Z]{3,}', line):
            h = doc.add_heading(line, level=1)
            for run in h.runs:
                run.font.color.rgb = RGBColor(0x01, 0x18, 0xD8)
        elif line.startswith('- ') or line.startswith('* '):
            doc.add_paragraph(line[2:], style='List Bullet')
        elif line.startswith('**') and line.endswith('**'):
            p = doc.add_paragraph()
            p.add_run(line.strip('*')).bold = True
        else:
            doc.add_paragraph(line)
    return doc

@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json(force=True)
    text = data.get('text', '')
    filename = data.get('filename', 'BusinessCase_xNova.docx')
    if not text:
        return jsonify({'error': 'No text provided'}), 400
    doc = parse_and_build_docx(text)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
