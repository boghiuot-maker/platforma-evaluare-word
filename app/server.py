from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, send_file
import os, uuid, zipfile
from pathlib import Path
from evaluator import evaluate_student_package, generate_report_files
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
FILES_DIR = BASE_DIR / "files"
REPORTS_DIR = BASE_DIR / "reports"

ALLOWED_EXT = set(["pptx","ppt","pdf","png","jpg","jpeg","docx","zip","txt"])

app = Flask(__name__)
app.config['SECRET_KEY'] = 'change-me-in-production'
app.config['UPLOAD_FOLDER'] = str(UPLOAD_DIR)
UPLOAD_DIR.mkdir(exist_ok=True)
REPORTS_DIR.mkdir(exist_ok=True)

def allowed_file(filename):
    return "." in filename and filename.rsplit(".",1)[1].lower() in ALLOWED_EXT

@app.route("/", methods=["GET"])
def index():
    # list support files
    support_files = sorted([p.name for p in FILES_DIR.glob("comp_p*.pptx")])
    return render_template("index.html", support_files=support_files)

@app.route("/download/<fname>")
def download_file(fname):
    fpath = FILES_DIR / fname
    if not fpath.exists():
        return "File not found", 404
    return send_file(str(fpath), as_attachment=True)

@app.route("/download_zip")
def download_zip():
    # create zip of all comp_p*.pptx
    zip_path = BASE_DIR / "support_files.zip"
    with zipfile.ZipFile(zip_path, "w") as z:
        for p in FILES_DIR.glob("comp_p*.pptx"):
            z.write(p, arcname=p.name)
    return send_file(str(zip_path), as_attachment=True)

@app.route("/submit", methods=["POST"])
def submit():
    # collect student metadata
    student_name = request.form.get("student_name","").strip()
    student_class = request.form.get("student_class","").strip()
    test_date = request.form.get("test_date","").strip()
    if not student_name:
        flash("Nume elev obligatoriu", "danger")
        return redirect(url_for("index"))
    # create unique folder for this submission
    sid = uuid.uuid4().hex[:10]
    folder = UPLOAD_DIR / f"{student_name.replace(' ','_')}_{student_class}_{sid}"
    folder.mkdir(parents=True, exist_ok=True)
    # handle 6 pptx uploads and optional screenshots
    uploaded = []
    for i in range(1,7):
        # pptx file input name: ppt_upload_1 ... ppt_upload_6
        f = request.files.get(f"ppt_upload_{i}")
        if f and f.filename:
            filename = secure_filename(f.filename)
            f.save(str(folder / f"app{i}_" + filename))
            uploaded.append(folder / ("app%d_"%i + filename))
        # optional capture upload
        c = request.files.get(f"capture_{i}")
        if c and c.filename:
            cname = secure_filename(c.filename)
            c.save(str(folder / f"cap{i}_" + cname))
    # run evaluation
    results = evaluate_student_package(folder)
    # generate report files (pdf & xlsx)
    pdf_path, xlsx_path = generate_report_files(folder, student_name, student_class, test_date, results, REPORTS_DIR)
    return send_file(str(pdf_path), as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
