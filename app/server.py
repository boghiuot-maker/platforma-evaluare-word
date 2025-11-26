from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from pathlib import Path
import uuid, zipfile, os
from werkzeug.utils import secure_filename
from app.evaluator import evaluate_student_package, generate_report_files

BASE = Path(__file__).resolve().parent
UPLOADS = BASE / "uploads"
FILES = BASE.parent / "files"
REPORTS = BASE / "reports"

app = Flask(__name__)
app.config['SECRET_KEY'] = 'change-me'

@app.route('/', methods=['GET'])
def index():
    support_files = sorted([p.name for p in FILES.glob('comp_p*.pptx')])
    return render_template('index.html', support_files=support_files)

@app.route('/download/<fname>')
def download_file(fname):
    fpath = FILES / fname
    if not fpath.exists():
        return 'File not found', 404
    return send_file(str(fpath), as_attachment=True)

@app.route('/download_zip')
def download_zip():
    zip_path = BASE.parent / 'support_files.zip'
    with zipfile.ZipFile(zip_path, 'w') as z:
        for p in FILES.glob('comp_p*.pptx'):
            z.write(p, arcname=p.name)
    return send_file(str(zip_path), as_attachment=True)

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form.get('student_name','').strip()
    klass = request.form.get('student_class','').strip()
    date = request.form.get('test_date','').strip()
    if not name:
        flash('Nume elev obligatoriu','danger')
        return redirect(url_for('index'))
    UPLOADS.mkdir(parents=True, exist_ok=True)
    sid = uuid.uuid4().hex[:8]
    student_folder = UPLOADS / f"{name.replace(' ','_')}_{klass}_{sid}"
    student_folder.mkdir(parents=True, exist_ok=True)
    for i in range(1,7):
        f = request.files.get(f"ppt_upload_{i}")
        if f and f.filename:
            f.save(str(student_folder / f"app{i}_{secure_filename(f.filename)}"))
        if i==2:
            c = request.files.get('capture_2')
            if c and c.filename:
                c.save(str(student_folder / f"cap2_{secure_filename(c.filename)}"))
            else:
                (student_folder / 'cap2_MISSING').write_text('MISSING')
    # app6 export pdf required
    ep = request.files.get('export_pdf')
    if ep and ep.filename:
        ep.save(str(student_folder / secure_filename(ep.filename)))
    else:
        (student_folder / 'rezolvare_missing').write_text('MISSING')
    results = evaluate_student_package(student_folder)
    pdfp, xlsxp = generate_report_files(student_folder, name, klass, date, results, REPORTS)
    return send_file(str(pdfp), as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
