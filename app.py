
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, abort, send_file
import os, shutil, zipfile
from werkzeug.utils import secure_filename
import pandas as pd
from io import BytesIO

app = Flask(__name__)
app.secret_key = 'change-me-for-prod'

BASE = os.path.dirname(os.path.abspath(__file__))
UPLOAD = os.path.join(BASE, 'uploads')
os.makedirs(UPLOAD, exist_ok=True)

PW = "prof2025"

ALLOWED_PPT = {'.pptx'}
ALLOWED_IMG = {'.gif', '.png', '.jpg', '.jpeg'}

def allowed_ext(filename, allowed):
    return '.' in filename and os.path.splitext(filename)[1].lower() in allowed

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download/<name>')
def download_sample(name):
    safe = secure_filename(name)
    path = os.path.join(BASE, 'static')
    if not os.path.exists(os.path.join(path, safe)):
        abort(404)
    return send_from_directory(path, safe, as_attachment=True)

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form.get('name','').strip()
    clasa = request.form.get('clasa','').strip()
    data_test = request.form.get('data','').strip()
    # files
    f_ap1 = request.files.get('ap1_file')
    f_ap2 = request.files.get('ap2_pptx')
    f_ap2_img = request.files.get('ap2_img')  # optional image/gif
    # validation (logical, not blocking for optional)
    if not name or not clasa or not data_test:
        flash('Nume, clasa si data sunt obligatorii.', 'error')
        return redirect(url_for('index'))
    if not f_ap1 or f_ap1.filename == '' or not allowed_ext(f_ap1.filename, ALLOWED_PPT):
        flash('Trebuie sa incarci fisierul PPTX pentru Aplicația 1 (format .pptx).', 'error')
        return redirect(url_for('index'))
    if not f_ap2 or f_ap2.filename == '' or not allowed_ext(f_ap2.filename, ALLOWED_PPT):
        flash('Trebuie sa incarci fisierul PPTX pentru Aplicația 2 (format .pptx).', 'error')
        return redirect(url_for('index'))
    # create student dir
    sd = os.path.join(UPLOAD, secure_filename(name))
    os.makedirs(sd, exist_ok=True)
    # save files
    ap1_name = 'ap1_' + secure_filename(f_ap1.filename)
    ap2_name = 'ap2_' + secure_filename(f_ap2.filename)
    f_ap1.save(os.path.join(sd, ap1_name))
    f_ap2.save(os.path.join(sd, ap2_name))
    img_name = ''
    if f_ap2_img and f_ap2_img.filename != '':
        img_name = 'ap2img_' + secure_filename(f_ap2_img.filename)
        f_ap2_img.save(os.path.join(sd, img_name))
    # save meta
    meta = os.path.join(sd, 'meta.txt')
    with open(meta, 'w', encoding='utf-8') as m:
        m.write(f"Nume={name}\\nClasa={clasa}\\nData={data_test}\\nAP1={ap1_name}\\nAP2={ap2_name}\\nAP2IMG={img_name}\\n")
    return render_template('submitted.html', name=name, files=[ap1_name, ap2_name] + ([img_name] if img_name else []), student_dir=os.path.basename(sd))

@app.route('/uploads/<student>/<filename>')
def uploaded_file(student, filename):
    sd = os.path.join(UPLOAD, secure_filename(student))
    return send_from_directory(sd, secure_filename(filename), as_attachment=True)

# ---- admin ----
@app.route('/admin', methods=['GET','POST'])
def admin_login():
    pw = request.args.get('pw')
    if not pw:
        return render_template('admin_login.html', error=None)
    if pw != PW:
        return render_template('admin_login.html', error='Parola greșită')
    # list students
    rows = []
    for s in os.listdir(UPLOAD):
        sd = os.path.join(UPLOAD, s)
        if os.path.isdir(sd):
            row = {'student': s, 'files': []}
            meta = os.path.join(sd, 'meta.txt')
            if os.path.exists(meta):
                with open(meta, encoding='utf-8') as f:
                    for line in f:
                        if '=' in line:
                            k,v = line.strip().split('=',1)
                            row[k]=v
            row['files'] = os.listdir(sd)
            rows.append(row)
    return render_template('admin.html', rows=rows, pw=PW)

@app.route('/admin/download_zip/<student>')
def admin_download_zip(student):
    sd = os.path.join(UPLOAD, secure_filename(student))
    mem = BytesIO()
    with zipfile.ZipFile(mem, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in os.listdir(sd):
            z.write(os.path.join(sd,f), arcname=f)
    mem.seek(0)
    return send_file(mem, download_name=f"{student}.zip", as_attachment=True)

@app.route('/admin/export_xlsx')
def admin_export_xlsx():
    if request.args.get('pw') != PW:
        abort(403)
    data = []
    for s in os.listdir(UPLOAD):
        sd = os.path.join(UPLOAD, s)
        if os.path.isdir(sd):
            row = {'student': s}
            meta = os.path.join(sd, 'meta.txt')
            if os.path.exists(meta):
                with open(meta, encoding='utf-8') as f:
                    for line in f:
                        if '=' in line:
                            k,v = line.strip().split('=',1)
                            row[k]=v
            data.append(row)
    df = pd.DataFrame(data)
    mem = BytesIO()
    df.to_excel(mem, index=False)
    mem.seek(0)
    return send_file(mem, download_name='rezultate.xlsx', as_attachment=True)

@app.route('/admin/delete_all', methods=['GET'])
def admin_delete_all():
    if request.args.get('pw') != PW:
        abort(403)
    shutil.rmtree(UPLOAD)
    os.makedirs(UPLOAD, exist_ok=True)
    return redirect(url_for('admin', pw=PW))

if __name__=='__main__':
    app.run(debug=True)
