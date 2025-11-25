
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, abort, send_file
import os, shutil
from werkzeug.utils import secure_filename
import pandas as pd
import zipfile
from io import BytesIO

app=Flask(__name__)
app.secret_key="key"

BASE=os.path.dirname(os.path.abspath(__file__))
UPLOAD=os.path.join(BASE,"uploads")
os.makedirs(UPLOAD,exist_ok=True)

# ---- Student area ----

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download/<name>')
def download_sample(name):
    f=secure_filename(name)
    p=os.path.join(BASE,"static",f)
    if not os.path.exists(p): abort(404)
    return send_from_directory(os.path.join(BASE,"static"), f, as_attachment=True)

def allowed(fn, exts):
    return "." in fn and os.path.splitext(fn)[1].lower() in exts

@app.route('/submit',methods=['POST'])
def submit():
    name=request.form.get("name","").strip()
    cl=request.form.get("clasa","").strip()
    dt=request.form.get("data","").strip()
    f1=request.files.get("ap1_file")
    f2=request.files.get("ap2_pptx")
    f3=request.files.get("ap2_gif")
    if not name or not cl or not dt:
        flash("Completeaza nume/clasa/data","error"); return redirect('/')
    if not f1 or not allowed(f1.filename,{".pptx"}):
        flash("Incarca PPTX aplicatia 1","error"); return redirect('/')
    if not f2 or not allowed(f2.filename,{".pptx"}):
        flash("Incarca PPTX aplicatia 2","error"); return redirect('/')
    if not f3 or not allowed(f3.filename,{".gif"}):
        flash("Incarca GIF aplicatia 2","error"); return redirect('/')

    sd=os.path.join(UPLOAD,secure_filename(name))
    os.makedirs(sd,exist_ok=True)
    ap1="ap1_"+secure_filename(f1.filename)
    ap2="ap2_"+secure_filename(f2.filename)
    gif="gif_"+secure_filename(f3.filename)
    f1.save(os.path.join(sd,ap1))
    f2.save(os.path.join(sd,ap2))
    f3.save(os.path.join(sd,gif))
    with open(os.path.join(sd,"meta.txt"),"w",encoding="utf8") as m:
        m.write(f"Nume={name}\nClasa={cl}\nData={dt}\nAP1={ap1}\nAP2={ap2}\nGIF={gif}\n")
    return render_template("submitted.html",name=name,files=[ap1,ap2,gif],student_dir=os.path.basename(sd))

@app.route('/uploads/<student>/<fn>')
def up(student,fn):
    sd=os.path.join(UPLOAD,secure_filename(student))
    return send_from_directory(sd, secure_filename(fn), as_attachment=True)

# ---- Admin ----

PW="prof2025"

@app.route('/admin')
def admin():
    if request.args.get("pw")!=PW:
        return render_template("admin_login.html",error="Parola greșită" if "pw" in request.args else None)
    # load students
    rows=[]
    for s in os.listdir(UPLOAD):
        sd=os.path.join(UPLOAD,s)
        if os.path.isdir(sd):
            meta=os.path.join(sd,"meta.txt")
            row={"student":s}
            if os.path.exists(meta):
                with open(meta,encoding="utf8") as f: 
                    for line in f:
                        if "=" in line:
                            k,v=line.strip().split("=",1)
                            row[k]=v
            row["files"]=os.listdir(sd)
            rows.append(row)
    return render_template("admin.html",rows=rows,pw=PW)

@app.route('/admin/download_zip/<student>')
def admin_zip(student):
    sd=os.path.join(UPLOAD,secure_filename(student))
    mem=BytesIO()
    with zipfile.ZipFile(mem,'w',zipfile.ZIP_DEFLATED) as z:
        for f in os.listdir(sd):
            z.write(os.path.join(sd,f), arcname=f)
    mem.seek(0)
    return send_file(mem, download_name=f"{student}.zip", as_attachment=True)

@app.route('/admin/export_xlsx')
def export_xlsx():
    if request.args.get("pw")!=PW: abort(403)
    data=[]
    for s in os.listdir(UPLOAD):
        sd=os.path.join(UPLOAD,s)
        if os.path.isdir(sd):
            row={"student":s}
            meta=os.path.join(sd,"meta.txt")
            if os.path.exists(meta):
                with open(meta,encoding="utf8") as f:
                    for line in f:
                        if "=" in line:
                            k,v=line.strip().split("=",1)
                            row[k]=v
            data.append(row)
    df=pd.DataFrame(data)
    mem=BytesIO()
    df.to_excel(mem,index=False)
    mem.seek(0)
    return send_file(mem,download_name="rezultate.xlsx",as_attachment=True)

@app.route('/admin/delete_all')
def delete_all():
    if request.args.get("pw")!=PW: abort(403)
    shutil.rmtree(UPLOAD)
    os.makedirs(UPLOAD)
    return redirect(url_for('admin',pw=PW))

if __name__=="__main__":
    app.run(debug=True)
