
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import os, uuid, zipfile
from pathlib import Path
from werkzeug.utils import secure_filename
from app.evaluator import evaluate_student_package, generate_report_files

BASE = Path(__file__).resolve().parent
UPLOAD = BASE/'uploads'

# ensure upload dir exists
UPLOAD.mkdir(parents=True, exist_ok=True)
FILES = BASE/'files'
REPORTS = BASE/'reports'

app = Flask(__name__)
app.config['SECRET_KEY'] = 'x'

@app.route('/')
def index():
    support = sorted([p.name for p in FILES.glob('comp_p*.pptx')])
    return render_template('index.html', support_files=support)

@app.route('/download/<f>')
def dl(f):
    return send_file(str(FILES/f), as_attachment=True)

@app.route('/download_zip')
def dlzip():
    z = BASE/'support.zip'
    with zipfile.ZipFile(z,'w') as zf:
        for p in FILES.glob('comp_p*.pptx'):
            zf.write(p, arcname=p.name)
    return send_file(str(z), as_attachment=True)

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form.get('student_name','').strip()
    klass = request.form.get('student_class','').strip()
    date = request.form.get('test_date','')
    if not name:
        flash('Numele este obligatoriu.')
        return redirect(url_for('index'))
    sid = uuid.uuid4().hex[:8]
    folder = UPLOAD / f"{name.replace(' ','_')}_{klass}_{sid}"
    folder.mkdir(parents=True, exist_ok=True)

    for i in range(1,7):
        f = request.files.get(f"ppt_upload_{i}")
        if f and f.filename:
            f.save(str(folder/f"app{i}_{secure_filename(f.filename)}"))

        c = request.files.get(f"capture_{i}")
        if c and c.filename:
            c.save(str(folder/f"cap{i}_{secure_filename(c.filename)}"))
        else:
            if i==2:
                (folder/'cap2_MISSING').write_text('MISSING')

    results = evaluate_student_package(folder)
    pdf, xlsx = generate_report_files(folder, name, klass, date, results, REPORTS)
    return send_file(str(pdf), as_attachment=True)

    # handle export PDF upload (required)
    ep = request.files.get('export_pdf')
    if ep and ep.filename:
        ep.save(str(folder / secure_filename(ep.filename)))
    else:
        (folder / 'export_missing').write_text('MISSING')
