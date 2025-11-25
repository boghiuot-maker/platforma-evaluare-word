from flask import Flask, request, jsonify, render_template, send_from_directory, redirect, url_for, flash
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

ADMIN_PASS = os.environ.get('ADMIN_PASS', 'profesor')

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
        style = getattr(p.style, 'name', '').lower()
        if style.startswith('heading'):
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
    except:
        return False
    return False

def verifica_pie_pagina(doc):
    try:
        for sec in doc.sections:
            if sec.footer and any(p.text.strip() for p in sec.footer.paragraphs):
                return True
    except:
        return False
    return False

def verifica_lista(doc):
    for p in doc.paragraphs:
        style = getattr(p.style, 'name', '').lower()
        if 'list' in style or 'bullet' in style or 'number' in style:
            return True
    return False

def numar_cuvinte(doc):
    total=0
    for p in doc.paragraphs:
        total += len(p.text.split())
    return total

def genereaza_feedback(rezultate, nr):
    fb=[]
    if not rezultate.get('titlu_corect'): fb.append('Titlul nu este setat ca Heading.')
    if not rezultate.get('are_imagine'): fb.append('Nu ai inserat o imagine.')
    if not rezultate.get('are_tabel'): fb.append('Nu ai inserat un tabel.')
    if not rezultate.get('are_antet'): fb.append('Antetul lipsește.')
    if not rezultate.get('are_pie_pagina'): fb.append('Nu există text în subsol.')
    if not rezultate.get('are_lista'): fb.append('Nu ai folosit liste.')
    if nr < 100: fb.append('Numărul de cuvinte este prea mic.')
    return fb if fb else ['Excelent!']

def calculeaza_scor(rezultate, nr):
    scor=0
    for crit,pts in CRITERIA.items():
        if crit=='numar_total_cuvinte':
            if nr>=100: scor+=pts
        else:
            if rezultate.get(crit): scor+=pts
    return scor, sum(CRITERIA.values())

def append_result(row):
    header=['timestamp','nume_elev','fisier','scor','punctaj_maxim','nr_cuvinte']+list(CRITERIA.keys())
    exists=os.path.exists(RESULTS_CSV)
    with open(RESULTS_CSV,'a',newline='',encoding='utf-8') as f:
        w=csv.writer(f)
        if not exists: w.writerow(header)
        w.writerow([row.get(h,'') for h in header])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download-template')
def download_template():
    return send_from_directory(DATA_FOLDER, TEMPLATE_FILE, as_attachment=True)

@app.route('/evaluate', methods=['POST'])
def evaluate():
    if 'file' not in request.files:
        return jsonify({'error':'Nu ai încărcat fișier'}),400
    file=request.files['file']
    name=request.form.get('nume','Elev').strip()
    if file.filename=='' or not allowed_file(file.filename):
        return jsonify({'error':'Fișier invalid'}),400
    filename=secure_filename(f"{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
    save_path=os.path.join(app.config['UPLOAD_FOLDER'],filename)
    os.makedirs(UPLOAD_FOLDER,exist_ok=True)
    file.save(save_path)
    doc=Document(save_path)
    rezultate={
        'titlu_corect':verifica_titlu(doc),
        'are_imagine':verifica_imagine(doc),
        'are_tabel':verifica_tabel(doc),
        'are_antet':verifica_antet(doc),
        'are_pie_pagina':verifica_pie_pagina(doc),
        'are_lista':verifica_lista(doc)
    }
    nr=numar_cuvinte(doc)
    scor,maxp=calculeaza_scor(rezultate,nr)
    fb=genereaza_feedback(rezultate,nr)
    row={'timestamp':datetime.datetime.now().isoformat(),'nume_elev':name,'fisier':filename,'scor':scor,'punctaj_maxim':maxp,'nr_cuvinte':nr}
    row.update(rezultate)
    os.makedirs(DATA_FOLDER,exist_ok=True)
    append_result(row)
    return jsonify({'rezultate':rezultate,'scor':scor,'punctaj_maxim':maxp,'feedback':fb,'nr_cuvinte':nr})

@app.route('/admin', methods=['GET','POST'])
def admin():
    pw=request.args.get('pw') or request.form.get('pw')
    if request.method=='POST' and pw==ADMIN_PASS:
        return redirect(url_for('admin'))
    if pw!=ADMIN_PASS:
        return render_template('admin_login.html')
    if os.path.exists(RESULTS_CSV):
        df=pd.read_csv(RESULTS_CSV)
        rec=df.to_dict(orient='records')
    else:
        rec=[]
    return render_template('admin.html',records=rec)

if __name__=='__main__':
    os.makedirs(UPLOAD_FOLDER,exist_ok=True)
    os.makedirs(DATA_FOLDER,exist_ok=True)
    app.run(host='0.0.0.0',port=5000,debug=True)
