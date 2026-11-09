"""Renders one sample agreement per template for the README (docs/img).

This mirrors what the VBA does (same {{placeholder}} rules and formats) so the
README can show output without Windows/Word. The real tool is the VBA in src/vba.
"""
import re, subprocess, sys
from copy import deepcopy
from datetime import date, datetime
from pathlib import Path
from docx import Document
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
TOKEN = re.compile(r"\{\{([A-Za-z0-9_]+)\}\}")

wb = load_workbook(ROOT / "AgreementGenerator.xlsx", data_only=True)
settings = {r[0]: (r[1] or "") for r in wb["Settings"].iter_rows(min_row=5, max_col=2, values_only=True) if r[0]}
tpls = {r[0]: r for r in wb["Templates"].iter_rows(min_row=5, max_col=4, values_only=True) if r[0]}
ws = wb["Agreements"]
hdr = [c.value for c in ws[5]]
rows = [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=6, values_only=True) if r[0]]

def fmt(v):
    if v is None: return ""
    if isinstance(v, (datetime, date)): return v.strftime("%d.%m.%Y")
    if isinstance(v, (int, float)): return f"{v:,.0f}".replace(",", " ")
    return str(v).strip()

def fill_par(p, vals, problems):
    for run in p.runs:
        def sub(m):
            k = m.group(1)
            if k not in vals: problems.add(f"Unknown {{{{{k}}}}}"); return m.group(0)
            if not vals[k]: problems.add(f"No value for {{{{{k}}}}}"); return m.group(0)
            return vals[k]
        run.text = TOKEN.sub(sub, run.text)

def all_pars(doc):
    yield from doc.paragraphs
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells: yield from c.paragraphs
    for s in doc.sections:
        yield from s.header.paragraphs; yield from s.footer.paragraphs

out = ROOT / "docs" / "img"; tmp = ROOT / "docs" / "_tmp"; tmp.mkdir(exist_ok=True)
done = set()
for row in rows:
    t = row["AgreementType"]
    if t in done: continue
    vals = {**settings, **{k: fmt(v) for k, v in row.items()},
            "DocumentDate": date.today().strftime("%d.%m.%Y"), "TemplateVersion": tpls[t][3]}
    doc = Document(ROOT / "templates" / tpls[t][1]); problems = set()
    for p in all_pars(doc): fill_par(p, vals, problems)
    if problems: print(row["AgreementID"], problems); continue
    name = f"{tpls[t][2]}_{row['AgreementID']}_sample"
    doc.save(tmp / f"{name}.docx")
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp), str(tmp / f"{name}.docx")],
                   check=True, capture_output=True)
    subprocess.run(["pdftoppm", "-png", "-r", "70", "-f", "1", "-l", "1", str(tmp / f"{name}.pdf"), str(out / name)], check=True)
    done.add(t); print("rendered", name)
