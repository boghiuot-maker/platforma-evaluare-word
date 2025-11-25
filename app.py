
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, abort
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'dev-key-change-me'

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_PPT = {'.pptx'}
ALLOWED_IMG = {'.gif', '.png', '.jpg', '.jpeg'}

def allowed_file(filename, allowed_exts):
    return '.' in filename and os.path.splitext(filename)[1].lower() in allowed_exts

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download/<name>')
def download_sample(name):
    safe = secure_filename(name)
    path = os.path.join(BASE_DIR, 'static')
    if not os.path.exists(os.path.join(path, safe)):
        abort(404)
    return send_from_directory(path, safe, as_attachment=True)

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form.get('name','').strip()
    clasa = request.form.get('clasa','').strip()
    data_test = request.form.get('data','').strip()
    f1 = request.files.get('ap1_file')
    f2 = request.files.get('ap2_pptx')
    f2gif = request.files.get('ap2_gif')
    if not name or not clasa or not data_test:
        flash('Nume, clasa si data sunt obligatorii.', 'error')
        return redirect(url_for('index'))
    if not f1 or f1.filename == '' or not allowed_file(f1.filename, ALLOWED_PPT):
        flash('Trebuie sa incarci fisierul PPTX pentru Aplicația 1.', 'error')
        return redirect(url_for('index'))
    if not f2 or f2.filename == '' or not allowed_file(f2.filename, ALLOWED_PPT):
        flash('Trebuie sa incarci fisierul PPTX pentru Aplicația 2.', 'error')
        return redirect(url_for('index'))
    if not f2gif or f2gif.filename == '' or not allowed_file(f2gif.filename, ALLOWED_IMG):
        flash('Trebuie sa incarci captura GIF pentru Aplicația 2.', 'error')
        return redirect(url_for('index'))
    student_dir = os.path.join(UPLOAD_FOLDER, secure_filename(name))
    os.makedirs(student_dir, exist_ok=True)
    ap1_name = 'aplicatia1_' + secure_filename(f1.filename)
    ap2_name = 'aplicatia2_' + secure_filename(f2.filename)
    gif_name = 'aplicatia2_' + secure_filename(f2gif.filename)
    f1.save(os.path.join(student_dir, ap1_name))
    f2.save(os.path.join(student_dir, ap2_name))
    f2gif.save(os.path.join(student_dir, gif_name))
    meta_path = os.path.join(student_dir, 'meta.txt')
    with open(meta_path, 'w', encoding='utf-8') as mf:
        mf.write(f'Nume={name}\\nClasa={clasa}\\nData={data_test}\\nAP1={ap1_name}\\nAP2={ap2_name}\\nAP2GIF={gif_name}\\n')
    return render_template('submitted.html', name=name, files=[ap1_name, ap2_name, gif_name], student_dir=os.path.basename(student_dir))

@app.route('/uploads/<student>/<filename>')
def uploaded_file(student, filename):
    student_dir = os.path.join(UPLOAD_FOLDER, secure_filename(student))
    safe = secure_filename(filename)
    return send_from_directory(student_dir, safe, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
