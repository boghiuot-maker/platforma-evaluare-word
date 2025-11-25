download-templatefrom flask import Flask, request, jsonify, render_template, send_from_directory, redirect, url_for, flash
from werkzeug.utils import secure_filename
from docx import Document
import os, csv, io, datetime
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

UPLOAD_FOLDER = 'uploads'
DATA_FOLDER = 'data'
RESULTS_CSV = os.path.join(DATA_FOLDER, 'results.csv')
TEMPLATE_FILE = 'test_practic_suport.docx'

ALLOWED_EXTENSIONS = {'doc', 'docx'}

app = Flask(__name__, static_folder='static', template_folder='templates')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = os.environ.get('FLASK_SECRET', 'schimba_secretul')

# Simple admin password (set in Render as ENV var ADMIN_PASS). Default: 'profesor'
ADMIN_PASS = os.environ.get('ADMIN_PASS', 'profesor')

# Scoring criteria and points
CRITERIA = {
    'titlu_corect': 2,
    'are_imagine': 1,
    'are_tabel': 1,
    'are_antet': 1,
    'are_pie_pagina': 1,
    'are_lista': 1,
    'numar_total_cuvinte': 3
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def verifica_titlu(doc):
    for p in doc.paragraphs:
        style_name = getattr(p.style, 'name', '').lower()
        if style_name.startswith('heading'):
            return True
    return False

def verifica_imagine(doc):
    return bool(getattr(doc, 'inline_shapes', []))

def verifica_tabel(doc):
    return len(doc.tables) > 0

def verifica_antet(doc):
    try:
        for sec in doc.sections:
            if sec.header and any(p.text.strip() for p in sec.header.paragraphs):
                return True
    except Exception:
        return False
    return False

def verifica_pie_pagina(doc):
    try:
        for sec in doc.sections:
            if sec.footer and any(p.text.strip() for p in sec.footer.paragraphs):
                return True
    except Exception:
        return False
    return False

def verifica_lista(doc):
    for p in doc.paragraphs:
        style = getattr(p.style, 'name', '').lower()
        if 'list' in style or 'bullet' in style or 'number' in style:
            return True
    return False

def numar_cuvinte(doc):
    total = 0
    for p in doc.paragraphs:
        total += len(p.text.split())
    return total

def genereaza_feedback(rezultate, nr_cuv):
    fb = []
    if not rezultate.get('titlu_corect'): fb.append('Titlul nu este setat ca Heading.')
    if not rezultate.get('are_imagine'): fb.append('Nu ai inserat o imagine.')
    if not rezultate.get('are_tabel'): fb.append('Nu ai inserat un tabel.')
    if not rezultate.get('are_antet'): fb.append('Antetul lipsește.')
    if not rezultate.get('are_pie_pagina'): fb.append('Nu există text în subsol.')
    if not rezultate.get('are_lista'): fb.append('Nu ai folosit liste (bullet/numere).')
    if not (nr_cuv >= 100): fb.append('Numărul de cuvinte este prea mic (<100).')
    return fb if fb else ['Excelent! Toate cerințele au fost îndeplinite.']

def calculeaza_scor(rezultate, nr_cuv):
    scor = 0
    for crit, pts in CRITERIA.items():
        if crit == 'numar_total_cuvinte':
            if nr_cuv >= 100: scor += pts
        else:
            if rezultate.get(crit): scor += pts
    return scor, sum(CRITERIA.values())

def append_result(row):
    header = ['timestamp','nume_elev','fisier','scor','punctaj_maxim','nr_cuvinte'] + list(CRITERIA.keys())
    exists = os.path.exists(RESULTS_CSV)
    with open(RESULTS_CSV, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if not exists:
            writer.writerow(header)
        writer.writerow([row.get(h,'') for h in header])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download-template')
def download_template():
    return send_from_directory('data', 'test_practic_suport.docx', as_attachment=True)

@app.route('/evaluate', methods=['POST'])
def evaluate():
    if 'file' not in request.files:
        return jsonify({'error':'Nu ai încărcat niciun fișier'}), 400
    file = request.files['file']
    name = request.form.get('nume', 'Elev fara nume').strip()
    if file.filename == '' or not allowed_file(file.filename):
        return jsonify({'error':'Fișier invalid. Încarcă .docx'}), 400
    filename = secure_filename(f"{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(save_path)
    doc = Document(save_path)
    rezultate = {}
    rezultate['titlu_corect'] = verifica_titlu(doc)
    rezultate['are_imagine'] = verifica_imagine(doc)
    rezultate['are_tabel'] = verifica_tabel(doc)
    rezultate['are_antet'] = verifica_antet(doc)
    rezultate['are_pie_pagina'] = verifica_pie_pagina(doc)
    rezultate['are_lista'] = verifica_lista(doc)
    nr = numar_cuvinte(doc)
    scor, max_total = calculeaza_scor(rezultate, nr)
    feedback = genereaza_feedback(rezultate, nr)
    row = {
        'timestamp': datetime.datetime.now().isoformat(),
        'nume_elev': name,
        'fisier': filename,
        'scor': scor,
        'punctaj_maxim': max_total,
        'nr_cuvinte': nr
    }
    row.update(rezultate)
    append_result(row)
    return jsonify({'rezultate':rezultate,'scor':scor,'punctaj_maxim':max_total,'feedback':feedback,'nr_cuvinte':nr})

@app.route('/admin', methods=['GET','POST'])
def admin():
    pw = request.args.get('pw') or request.form.get('pw')
    if request.method == 'POST' and pw == ADMIN_PASS:
        return redirect(url_for('admin'))
    if pw != ADMIN_PASS:
        return render_template('admin_login.html', error=None)
    if os.path.exists(RESULTS_CSV):
        df = pd.read_csv(RESULTS_CSV)
        records = df.to_dict(orient='records')
    else:
        records = []
    return render_template('admin.html', records=records)

@app.route('/admin/export/csv')
def export_csv():
    if request.args.get('pw') != ADMIN_PASS:
        return 'Unauthorized', 403
    if not os.path.exists(RESULTS_CSV):
        return 'No results', 404
    return send_from_directory(DATA_FOLDER, 'results.csv', as_attachment=True)

@app.route('/admin/export/xlsx')
def export_xlsx():
    if request.args.get('pw') != ADMIN_PASS:
        return 'Unauthorized', 403
    if not os.path.exists(RESULTS_CSV):
        return 'No results', 404
    df = pd.read_csv(RESULTS_CSV)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='results')
    out.seek(0)
    return (out.read(), 200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="results.xlsx"'
    })

@app.route('/admin/result/<filename>/pdf')
def result_pdf(filename):
    if request.args.get('pw') != ADMIN_PASS:
        return 'Unauthorized', 403
    if not os.path.exists(RESULTS_CSV):
        return 'No results', 404
    df = pd.read_csv(RESULTS_CSV)
    row = df[df['fisier'] == filename]
    if row.empty:
        return 'Not found', 404
    row = row.iloc[0].to_dict()
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
    width, height = A4
    y = height - 50
    c.setFont('Helvetica-Bold', 14)
    c.drawString(50, y, f"Rezultat evaluare - {row.get('nume_elev')}")
    y -= 30
    c.setFont('Helvetica', 11)
    for k, v in row.items():
        c.drawString(50, y, f"{k}: {v}")
        y -= 18
    c.showPage()
    c.save()
    out.seek(0)
    return (out.read(), 200, {
        'Content-Type': 'application/pdf',
        'Content-Disposition': f'attachment; filename="{filename}_result.pdf"'
    })

if __name__ == '__main__':
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(DATA_FOLDER, exist_ok=True)
    app.run(host='0.0.0.0', port=5000, debug=True)
