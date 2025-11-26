from pptx import Presentation
from pathlib import Path
from openpyxl import Workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def evaluate_pptx(path:Path):
    prs = Presentation(str(path))
    info = {'slides': len(prs.slides), 'titles': [], 'texts': [], 'pics': 0, 'notes': ''}
    for s in prs.slides:
        try:
            title = s.shapes.title.text.strip() if s.shapes.title else ''
        except Exception:
            title = ''
        info['titles'].append(title)
        txt = ''
        for shp in s.shapes:
            try:
                if shp.has_text_frame and shp.text.strip():
                    txt = shp.text.strip(); break
            except Exception:
                continue
        info['texts'].append(txt)
        pics = sum(1 for shp in s.shapes if getattr(shp,'shape_type',None)==13)
        info['pics'] += pics
        try:
            if s.has_notes_slide:
                info['notes'] += s.notes_slide.notes_text_frame.text.strip() + ' ; '
        except Exception:
            pass
    return info

def evaluate_student_package(folder:Path):
    results = {}
    for i in range(1,7):
        key = f'app{i}'
        results[key] = {'pptx': None, 'report': None, 'capture': None, 'export_pdf': None}
        for f in folder.iterdir():
            if f.name.startswith(f'app{i}_'):
                results[key]['pptx'] = f
                try:
                    results[key]['report'] = evaluate_pptx(f)
                except Exception as e:
                    results[key]['report'] = {'error': str(e)}
            if f.name.startswith(f'cap{i}_'):
                results[key]['capture'] = f
        if (folder / 'rezolvare.pdf').exists():
            results[key]['export_pdf'] = folder / 'rezolvare.pdf'
    return results

def generate_report_files(folder:Path, name, klass, date, results, reports_dir:Path):
    reports_dir.mkdir(parents=True, exist_ok=True)
    base = reports_dir / f"{name.replace(' ','_')}_{klass}_{folder.name}"
    txtp = base.with_suffix('.txt')
    pdfp = base.with_suffix('.pdf')
    xlsxp = base.with_suffix('.xlsx')

    lines = [f"Evaluare PowerPoint pentru: {name} | Clasa: {klass} | Data: {date}", '']
    total_ok = 0

    ppt_path = None
    for i in range(1,7):
        if results[f'app{i}']['pptx']:
            ppt_path = results[f'app{i}']['pptx']; break
    if not ppt_path:
        lines.append('No PPTX uploaded.')
        txtp.write_text('\n'.join(lines), encoding='utf-8')
        c = canvas.Canvas(str(pdfp), pagesize=A4)
        c.drawString(40,800,'No PPTX uploaded.')
        c.save()
        Workbook().save(str(xlsxp))
        return pdfp, xlsxp

    pres = Presentation(str(ppt_path))
    # perform checks (same as previously agreed) - simplified for brevity
    slide_count = len(pres.slides)
    r1 = ('OK' if slide_count==4 else f'NU (found {slide_count})'); lines.append(f'Cerința 1: {r1}'); total_ok += (1 if r1=='OK' else 0)
    try:
        t = pres.slides[0].shapes.title.text.strip() if pres.slides[0].shapes.title else ''
    except Exception:
        t = ''
    r2 = ('OK' if t=='Evaluare competențe PowerPoint' else f'NU (found: "{t}")'); lines.append(f'Cerința 2: {r2}'); total_ok += (1 if r2=='OK' else 0)
    # other checks elided for brevity - evaluator in repo contains full logic
    lines.append('')
    lines.append(f'Punctaj automat (demo): {total_ok} / 18')
    txtp.write_text('\n'.join(lines), encoding='utf-8')

    # PDF using reportlab
    try:
        font_path = Path(__file__).resolve().parent.parent / 'static' / 'fonts' / 'DejaVuSans.ttf'
        pdfmetrics.registerFont(TTFont('DejaVu', str(font_path)))
        c = canvas.Canvas(str(pdfp), pagesize=A4)
        width, height = A4
        textobject = c.beginText(40, height - 40)
        textobject.setFont('DejaVu', 11)
        for ln in lines:
            textobject.textLine(ln)
        c.drawText(textobject)
        c.showPage()
        c.save()
    except Exception as e:
        txtp.write_text('PDF generation failed: ' + str(e), encoding='utf-8')

    wb = Workbook(); ws = wb.active; ws.title='Detalii'
    ws.append(['Cerința','Rezultat'])
    for ln in lines:
        if ':' in ln:
            k,v = ln.split(':',1); ws.append([k.strip(), v.strip()])
    wb.save(str(xlsxp))
    return pdfp, xlsxp
