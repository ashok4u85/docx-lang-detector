import os, re, json
from flask import Flask, request, jsonify, send_file, render_template
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import RGBColor, Pt
from lingua import Language, LanguageDetectorBuilder
import tempfile

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB max

detector = (
    LanguageDetectorBuilder
    .from_all_languages()
    .with_minimum_relative_distance(0.15)
    .build()
)

SKIP = [
    r'^\d[\d\-T:\+\./ ]*$',
    r'^[A-Z]{1,6}[\d\-_]+$',
    r'^\([\w ]+\)$',
    r'^[A-Z0-9\-\._/]{3,}$',
    r'^\d+\.\d+$',
    r'^(NI|UNK|true|false|ichicsr)$',
    r'^(N\.|C\.|D\.|E\.|G\.|H\.)',
    r'^\d{4}-\d{2}-\d{2}',
    r'^Page \d',
    r'^(Element|Data|Number|Name|Information|Report|Summary|Narrative|Yes|No)$',
]

def should_skip(text):
    t = text.strip()
    if len(t) < 5:
        return True
    return any(re.match(p, t, re.IGNORECASE) for p in SKIP)

def detect(text):
    if should_skip(text):
        return None
    results = detector.compute_language_confidence_values(text.strip())
    if not results:
        return None
    top = results[0]
    if top.language == Language.ENGLISH or top.value < 0.55:
        return None
    return {'language': top.language.name.title().replace('_', ' '), 'confidence': round(top.value * 100, 1)}

def extract_paragraphs(doc):
    texts = []
    for para in doc.paragraphs:
        t = para.text.strip()
        if t:
            texts.append(t)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    t = para.text.strip()
                    if t:
                        texts.append(t)
    return list(dict.fromkeys(texts))

def highlight_run(run):
    rpr = run._r.get_or_add_rPr()
    hl = OxmlElement('w:highlight')
    hl.set(qn('w:val'), 'yellow')
    rpr.append(hl)

def add_lang_tag(para, lang, conf):
    run = para.add_run(f'  [\u26a0 {lang} | {conf}%]')
    run.bold = True
    run.font.color.rgb = RGBColor(0xE6, 0x00, 0x54)
    run.font.size = Pt(9)

def process_docx(filepath):
    doc = Document(filepath)
    findings = []

    def process_para(para):
        text = para.text.strip()
        if not text:
            return
        result = detect(text)
        if result:
            for run in para.runs:
                highlight_run(run)
            add_lang_tag(para, result['language'], result['confidence'])
            findings.append({'text': text[:80], 'language': result['language'], 'confidence': result['confidence']})

    for para in doc.paragraphs:
        process_para(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    process_para(para)

    # Summary page
    doc.add_page_break()
    p = doc.add_paragraph()
    run = p.add_run('Non-English Language Detection Report')
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0x3E, 0x01, 0x6F)

    if not findings:
        doc.add_paragraph('No non-English content detected.')
    else:
        doc.add_paragraph(f'Total instances found: {len(findings)}').runs[0].bold = True
        for i, f in enumerate(findings, 1):
            doc.add_paragraph(f"{i}. [{f['language']} | {f['confidence']}%]  {f['text']}")

    out = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    doc.save(out.name)
    return out.name, findings

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/scan', methods=['POST'])
def scan():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    if not f.filename.endswith('.docx'):
        return jsonify({'error': 'Only .docx files are supported'}), 400

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    f.save(tmp.name)

    try:
        out_path, findings = process_docx(tmp.name)
        os.unlink(tmp.name)
        return jsonify({'findings': findings, 'output_file': out_path})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download', methods=['POST'])
def download():
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    f = request.files['file']
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    f.save(tmp.name)
    out_path, _ = process_docx(tmp.name)
    os.unlink(tmp.name)
    original_name = f.filename.replace('.docx', '')
    return send_file(out_path, as_attachment=True,
                     download_name=f'{original_name}_LangDetected.docx',
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
