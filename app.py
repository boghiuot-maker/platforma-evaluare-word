from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, abort, send_file
import os, shutil, zipfile, json
from werkzeug.utils import secure_filename
import pandas as pd
from io import BytesIO

app = Flask(__name__)
app.secret_key = 'change-me'

BASE = os.path.dirname(os.path.abspath(__file__))
UPLOAD = os.path.join(BASE, 'uploads')
os.makedirs(UPLOAD, exist_ok=True)

PW = "prof2025"
ALLOWED_PPT = {'.pptx'}
ALLOWED_IMG = {'.gif','.png','.jpg','.jpeg'}

def allowed_ext(filename, allowed):
    return '.' in filename and os.path.splitext(filename)[1].lower() in allowed

def write_meta(student_dir, meta):
    with open(os.path.join(student_dir, 'meta.json'),'w',encoding='utf-8') as f:
        json.dump(meta,f,ensure_ascii=False,indent=2)

def read_meta(student_dir):
    path = os.path.join(student_dir, 'meta.json')
    if not os.path.exists(path):
        return None
    with open(path,encoding='utf-8') as f:
        return json.load(f)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download/<name>')
def download_sample(name):
    safe = secure_filename(name)
    p = os.path.join(BASE,'static')
    if not os.path.exists(os.path.join(p,safe)):
        abort(404)
    return send_from_directory(p, safe, as_attachment=True)

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form.get('name','').strip()
    clasa = request.form.get('clasa','').strip()
    data_test = request.form.get('data','').strip()
    # files for 6 apps
    files = {}
    for i in range(1,7):
        files[f'ap{i}'] = []
    if not name or not clasa or not data_test:
        flash('Nume, clasa si data sunt obligatorii.', 'error')
        return redirect(url_for('index'))
    # save student dir
    sd = os.path.join(UPLOAD, secure_filename(name))
    os.makedirs(sd, exist_ok=True)
    def save_file(field, prefix, allowed=None, optional=False):
        up = request.files.get(field)
        if not up or up.filename=='':
            if optional:
                return []
            else:
                return None
        if allowed and not allowed_ext(up.filename, allowed):
            pass
        fname = prefix + secure_filename(up.filename)
        up.save(os.path.join(sd, fname))
        return [fname]
    # ap1 required
    r = save_file('ap1_file','ap1_', allowed=ALLOWED_PPT, optional=False)
    if r is None:
        flash('Incarca PPTX Aplicația 1','error'); return redirect(url_for('index'))
    files['ap1'].extend(r)
    # ap2 pptx required
    r = save_file('ap2_pptx','ap2_', allowed=ALLOWED_PPT, optional=False)
    if r is None:
        flash('Incarca PPTX Aplicația 2','error'); return redirect(url_for('index'))
    files['ap2'].extend(r)
    # ap2 image optional
    r = save_file('ap2_img','ap2img_', allowed=None, optional=True)
    if r:
        files['ap2'].extend(r)
    # ap3..ap6 optional
    for i in [3,4,5,6]:
        r = save_file(f'ap{i}_pptx', f'ap{i}_', allowed=ALLOWED_PPT, optional=True)
        if r:
            files[f'ap{i}'].extend(r)
    meta = {'name': name, 'class': clasa, 'date': data_test, 'files': files}
    write_meta(sd, meta)
    return render_template('submitted.html', name=name, files=sum(files.values(),[]), student_dir=os.path.basename(sd))

@app.route('/uploads/<student>/<filename>')
def uploaded_file(student, filename):
    sd = os.path.join(UPLOAD, secure_filename(student))
    return send_from_directory(sd, secure_filename(filename), as_attachment=True)

@app.route('/admin')
def admin():
    pw = request.args.get('pw','')
    if pw != PW:
        return render_template('admin_login.html', error=None)
    # filters
    fclass = request.args.get('class','')
    fname = request.args.get('name','')
    fdate = request.args.get('date','')
    rows = []
    for s in sorted(os.listdir(UPLOAD)):
        sd = os.path.join(UPLOAD,s)
        if os.path.isdir(sd):
            meta = read_meta(sd)
            if meta is None:
                continue
            if fclass and meta.get('class','')!=fclass:
                continue
            if fname and fname.lower() not in meta.get('name','').lower():
                continue
            if fdate and meta.get('date','')!=fdate:
                continue
            rows.append({'student': s, 'meta': meta, 'files': os.listdir(sd)})
    return render_template('admin_v3.html', rows=rows, pw=PW, filters={'class':fclass,'name':fname,'date':fdate})

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
    if request.args.get('pw')!=PW:
        abort(403)
    fclass = request.args.get('class','')
    fname = request.args.get('name','')
    fdate = request.args.get('date','')
    data = []
    for s in sorted(os.listdir(UPLOAD)):
        sd = os.path.join(UPLOAD,s)
        if os.path.isdir(sd):
            meta = read_meta(sd)
            if meta is None: continue
            if fclass and meta.get('class','')!=fclass: continue
            if fname and fname.lower() not in meta.get('name','').lower(): continue
            if fdate and meta.get('date','')!=fdate: continue
            row = {'name': meta.get('name',''), 'class': meta.get('class',''), 'date': meta.get('date','')}
            for i in range(1,7):
                row[f'ap{i}'] = ','.join(meta.get('files',{}).get(f'ap{i}',[]))
            data.append(row)
    df = pd.DataFrame(data)
    mem = BytesIO()
    df.to_excel(mem, index=False)
    mem.seek(0)
    return send_file(mem, download_name='rezultate.xlsx', as_attachment=True)

@app.route('/admin/delete_all')
def admin_delete_all():
    if request.args.get('pw')!=PW:
        abort(403)
    shutil.rmtree(UPLOAD)
    os.makedirs(UPLOAD, exist_ok=True)
    return redirect(url_for('admin', pw=PW))

if __name__=='__main__':
    app.run(debug=True)
