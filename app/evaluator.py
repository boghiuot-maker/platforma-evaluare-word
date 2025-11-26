from pptx import Presentation
from pathlib import Path
import os, textwrap
from openpyxl import Workbook
from fpdf import FPDF

def emu_to_cm(emu):
    return emu / 360000.0

def has_hyperlink(shape):
    try:
        if hasattr(shape, "click_action"):
            ca = shape.click_action
            if getattr(ca, "hyperlink", None) is not None:
                return True
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if getattr(run, "hyperlink", None) is not None and run.hyperlink.address:
                        return True
    except Exception:
        return False

def evaluate_pptx(path):
    # best-effort evaluation similar to earlier logic
    prs = Presentation(str(path))
    info = {"slides": len(prs.slides), "titles":[], "texts":[], "pics":0, "notes":""}
    for i, slide in enumerate(prs.slides, start=1):
        title = ""
        try:
            title = slide.shapes.title.text.strip() if slide.shapes.title else ""
        except Exception:
            title = ""
        body = ""
        for shp in slide.shapes:
            try:
                if shp.has_text_frame and shp.text and shp.text.strip():
                    body = shp.text.strip()
                    break
            except Exception:
                pass
        notes = ""
        try:
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text.strip()
        except Exception:
            notes = ""
        # count pictures
        pics = 0
        for shp in slide.shapes:
            try:
                if getattr(shp, "shape_type", None) == 13:
                    pics += 1
            except Exception:
                pass
        info["titles"].append(title)
        info["texts"].append(body)
        info["pics"] += pics
        info["notes"] += notes + "; "
    return info

def evaluate_student_package(folder: Path):
    # look for app1..app6 uploads (any file starting with app1_)
    results = {}
    for i in range(1,7):
        key = f"app{i}"
        results[key] = {"found": False, "pptx": None, "report": None}
        for f in folder.iterdir():
            if f.name.startswith(f"app{i}_"):
                results[key]["found"] = True
                results[key]["pptx"] = f
                # evaluate
                try:
                    results[key]["report"] = evaluate_pptx(f)
                except Exception as e:
                    results[key]["report"] = {"error": str(e)}
    return results

def generate_report_files(folder: Path, name, klass, date, results, reports_dir: Path):
    # create TXT & PDF & XLSX. Return PDF and XLSX paths.
    reports_dir.mkdir(parents=True, exist_ok=True)
    base = reports_dir / f"{name.replace(' ','_')}_{klass}_{folder.name}"
    txtp = base.with_suffix(".txt")
    pdfp = base.with_suffix(".pdf")
    xlsxp = base.with_suffix(".xlsx")
    # TXT
    lines = []
    lines.append(f"Rezultat evaluare pentru: {name} | Clasa: {klass} | Data: {date}")
    lines.append("")
    total_auto = 0
    for i in range(1,7):
        key = f"app{i}"
        if results[key]["found"]:
            rpt = results[key]["report"]
            lines.append(f"=== Aplicația {i} ({results[key]['pptx'].name}) ===")
            lines.append(str(rpt))
            total_auto += rpt.get("pics",0) if isinstance(rpt, dict) else 0
        else:
            lines.append(f"=== Aplicația {i} === NOT UPLOADED")
    lines.append("")
    lines.append(f"Scor automat (metrica demonstrativă): {total_auto}")
    txtp.write_text("\n".join(lines), encoding="utf-8")
    # PDF (simple)
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    for ln in lines:
        pdf.multi_cell(0, 6, ln)
    pdf.output(str(pdfp))
    # XLSX simple summary
    wb = Workbook()
    ws = wb.active
    ws.title = "Rezumat"
    ws.append(["Aplicația","Status","Slides","PicturesCount","NotesSnippet"])
    for i in range(1,7):
        key=f"app{i}"
        if results[key]["found"] and isinstance(results[key]["report"], dict):
            rpt = results[key]["report"]
            ws.append([f"Aplicația {i}", "Uploaded", rpt.get("slides",0), rpt.get("pics",0), (rpt.get("texts",[""])[0][:80] if rpt.get("texts") else "")])
        else:
            ws.append([f"Aplicația {i}", "Not uploaded",0,0,""])
    wb.save(str(xlsxp))
    return pdfp, xlsxp
