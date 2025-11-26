
from pptx import Presentation
from pathlib import Path
from fpdf import FPDF
from openpyxl import Workbook

def evaluate_pptx(p):
    prs = Presentation(str(p))
    info = {'slides':len(prs.slides),'pics':0,'texts':[]}
    for s in prs.slides:
        txt=''
        for sh in s.shapes:
            if hasattr(sh,'text') and sh.text:
                txt=sh.text.strip()
        pics=sum(1 for sh in s.shapes if getattr(sh,'shape_type',0)==13)
        info['texts'].append(txt)
        info['pics']+=pics
    return info

def evaluate_student_package(folder):
    res={}
    for i in range(1,7):
        key=f"app{i}"
        res[key]={'ppt':None,'capture':None,'report':None}
        for f in folder.iterdir():
            if f.name.startswith(f'app{i}_'):
                res[key]['ppt']=f
                try:
                    res[key]['report']=evaluate_pptx(f)
                except Exception as e:
                    res[key]['report']={'error':str(e)}
            if f.name.startswith(f'cap{i}_'):
                res[key]['capture']=f
        if (folder/f"cap{i}_MISSING").exists():
            res[key]['capture']=None
    return res

def generate_report_files(folder,name,klass,date,res,out):
    out.mkdir(exist_ok=True)
    base = out/f"{name}_{klass}_{folder.name}"
    txt = base.with_suffix('.txt')
    pdf = base.with_suffix('.pdf')
    xlsx= base.with_suffix('.xlsx')

    lines=[f"Evaluare PowerPoint pentru {name}, clasa {klass}, data {date}",""]
    for i in range(1,7):
        k=f"app{i}"
        lines.append(f"--- Aplicația {i} ---")
        if res[k]['ppt'] is None:
            lines.append("NU A FOST ÎNCĂRCATĂ")
        else:
            lines.append(str(res[k]['report']))
            if i==2:
                lines.append("Captură: "+("DA" if res[k]['capture'] else "NU"))
        lines.append("")

    txt.write_text("\n".join(lines),encoding='utf8')

    p=FPDF()
    p.add_page()
    p.add_font("DejaVu","", "static/fonts/DejaVuSans.ttf", uni=True)
    p.set_font("DejaVu", size=11)
    for ln in lines:
        p.multi_cell(0,6,ln)
    p.output(str(pdf))

    wb=Workbook()
    ws=wb.active
    ws.append(["App","Slides","Pics","Text"])
    for i in range(1,7):
        k=f"app{i}"
        r=res[k]['report']
        if isinstance(r,dict):
            ws.append([i,r.get('slides',0),r.get('pics',0),r.get('texts',[''])[0]])
        else:
            ws.append([i,0,0,""])
    wb.save(str(xlsx))
    return pdf,xlsx
