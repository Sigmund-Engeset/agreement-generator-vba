"""Builds AgreementGenerator.xlsx: the tracker workbook with mock data.

All companies, people and organisation numbers are fictional and randomly
generated. Norwegian org numbers are made MOD11-valid so the VBA validation has
something real to check; two rows are deliberately broken to demo validation.
"""
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.hyperlink import Hyperlink

random.seed(42)
ROOT = Path(__file__).resolve().parents[1]

# "Autumn Harvest" palette (same as the LinkedIn carousel)
NAVY, NAVY2, SAND, GOLD, ORANGE = "202C39", "283845", "B8B08D", "F2D492", "F29559"
TEAL, SLATE, SAGE, BLUE, RED = NAVY, NAVY2, SAND, GOLD, ORANGE  # role aliases used below
INK, MUTED, TEAL50, BG = "202C39", "4A5866", "E7E1CF", "FAF7EE"
FONT = "Segoe UI"  # Windows default UI font; renders the same on any Windows PC

STATUSES = ["Not created", "Created", "Sent for signing", "Signed", "Processed"]
TYPES = [  # AgreementType, TemplateFile, FilePrefix, TemplateVersion
    ("Service Agreement", "Service_Agreement.docx", "SA", "1.0"),
    ("Mutual NDA", "Mutual_NDA.docx", "NDA", "1.0"),
    ("Consultancy Agreement", "Consultancy_Agreement.docx", "CA", "1.0"),
]

COMPANIES = [  # name, address, postcode, city, country
    ("Fjellheim Logistikk AS", "Industriveien 14", "2008", "Fjerdingby", "Norway"),
    ("Havbris Sjømat AS", "Kaigata 3", "6003", "Ålesund", "Norway"),
    ("Nordlysa Energi AS", "Storgata 88", "9008", "Tromsø", "Norway"),
    ("Bjørkly Eiendom AS", "Kirkeveien 52", "0368", "Oslo", "Norway"),
    ("Vestrand Bygg AS", "Sandviksveien 21", "5035", "Bergen", "Norway"),
    ("Granlund Helse AS", "Olav Tryggvasons gate 9", "7011", "Trondheim", "Norway"),
    ("Solvik Retail AS", "Markveien 40", "0554", "Oslo", "Norway"),
    ("Elvetun Software AS", "Kongens gate 17", "4610", "Kristiansand", "Norway"),
    ("Steinsland Maskin AS", "Verkstedveien 6", "4313", "Sandnes", "Norway"),
    ("Kvitfjord Media AS", "Prinsens gate 2", "0157", "Oslo", "Norway"),
    ("Lauvås Regnskap AS", "Torget 5", "3015", "Drammen", "Norway"),
    ("Mjøsli Agro AS", "Gårdsvegen 11", "2312", "Ottestad", "Norway"),
    ("Sjøberg Analys AB", "Drottninggatan 61", "111 21", "Stockholm", "Sweden"),
    ("Ekdal Industri AB", "Hamngatan 9", "411 06", "Göteborg", "Sweden"),
    ("Lindqvist Design AB", "Södra Förstadsgatan 30", "211 43", "Malmö", "Sweden"),
    ("Holmgaard Data ApS", "Vesterbrogade 74", "1620", "København V", "Denmark"),
    ("Rasmussen Transport A/S", "Havnegade 12", "8000", "Aarhus C", "Denmark"),
    ("Tervola Digital Oy", "Mannerheimintie 20", "00100", "Helsinki", "Finland"),
]
FIRST = ["Kari", "Ola", "Ingrid", "Lars", "Maja", "Erik", "Sofie", "Anders", "Nora", "Jonas",
         "Elin", "Magnus", "Astrid", "Henrik", "Linnea", "Mikkel", "Freja", "Aino"]
LAST = ["Haugen", "Berg", "Lund", "Nilsen", "Strand", "Moe", "Lindberg", "Johansson",
        "Sørensen", "Holm", "Virtanen", "Bakke", "Eriksen", "Dahl"]
TITLES = ["CEO", "CFO", "COO", "Head of Procurement", "Head of IT", "Finance Manager",
          "Operations Director", "Head of Data"]
SIGNERS = [("Ingrid Solberg", "Chief Executive Officer"), ("Henrik Aas", "Head of Sales"),
           ("Marte Lie", "Commercial Director")]
OWNERS = ["Henrik Aas", "Marte Lie", "Jonas Berge"]
SCOPE = {
    "Service Agreement": [
        "Monthly management reporting in Power BI, including data model maintenance and up to 8 hours of change requests per month.",
        "Operation and support of the Customer's data platform (Snowflake), including monitoring, user administration and quarterly health checks.",
        "Hosting, maintenance and second-line support of the Customer's sales dashboard suite.",
        "Automated month-end reconciliation of accounts receivable, delivered as a managed service.",
    ],
    "Mutual NDA": [
        "Evaluating a potential data and analytics partnership between the Parties.",
        "Discussing a pilot project for automated customer onboarding.",
        "Exchanging information to scope a migration of reporting from Excel to Power BI.",
    ],
    "Consultancy Agreement": [
        "Designing and building a sales pipeline dashboard in Power BI, including data model and DAX measures.",
        "Analysis of customer churn drivers and a recommendation report for management.",
        "Automating document production from Excel (VBA) and training two key users.",
        "Requirements analysis and prototype of a resource allocation app in Power Apps.",
    ],
}


def mod11_orgno():
    """Random 9-digit Norwegian org number starting with 9, MOD11-valid."""
    w = [3, 2, 7, 6, 5, 4, 3, 2]
    while True:
        d = [9] + [random.randint(0, 9) for _ in range(7)]
        r = 11 - sum(a * b for a, b in zip(d, w)) % 11
        r = 0 if r == 11 else r
        if r != 10:
            s = "".join(map(str, d + [r]))
            return f"{s[:3]} {s[3:6]} {s[6:]}"


def foreign_orgno(country):
    if country == "Sweden":
        return f"556{random.randint(100, 999)}-{random.randint(1000, 9999)}"
    if country == "Denmark":
        return f"DK{random.randint(10_000_000, 99_999_999)}"
    return f"FI{random.randint(1_000_000, 9_999_999)}-{random.randint(0, 9)}"


def build_rows(n=30):
    rows, orgnos = [], {}
    start0 = date(2026, 11, 1)
    for i in range(1, n + 1):
        comp = random.choice(COMPANIES)
        name, addr, pc, city, country = comp
        orgnos.setdefault(name, mod11_orgno() if country == "Norway" else foreign_orgno(country))
        atype = random.choices([t[0] for t in TYPES], weights=[45, 25, 30])[0]
        fn, ln = random.choice(FIRST), random.choice(LAST)
        domain = name.split()[0].lower().translate(str.maketrans("æøåäöé", "aoaaoe"))
        tld = {"Norway": "no", "Sweden": "se", "Denmark": "dk", "Finland": "fi"}[country]
        signer = random.choice(SIGNERS)
        currency = {"Norway": "NOK", "Sweden": "SEK", "Denmark": "DKK", "Finland": "EUR"}[country]
        start = start0 + timedelta(days=random.choice([0, 14, 30, 31, 45, 61, 75, 92]))
        start = start.replace(day=1) if atype != "Mutual NDA" else start
        row = {
            "AgreementID": f"AGR-2026-{i:03d}",
            "AgreementType": atype,
            "Status": "Not created",
            "CounterpartyName": name,
            "CounterpartyOrgNo": orgnos[name],
            "CounterpartyAddress": addr,
            "CounterpartyPostcode": pc,
            "CounterpartyCity": city,
            "CounterpartyCountry": country,
            "ContactName": f"{fn} {ln}",
            "ContactTitle": random.choice(TITLES),
            "ContactEmail": f"{fn.lower()}.{ln.lower().translate(str.maketrans('øæå', 'oea'))}@{domain}.{tld}",
            "StartDate": start,
            "TermMonths": {"Service Agreement": random.choice([12, 24, 36]),
                           "Mutual NDA": random.choice([12, 24]),
                           "Consultancy Agreement": random.choice([2, 3, 4, 6])}[atype],
            "EndDate": None,  # formula
            "ContractValue": None,
            "HourlyRate": None,
            "Currency": currency if atype != "Mutual NDA" else None,
            "PaymentTermsDays": random.choice([14, 30]) if atype != "Mutual NDA" else None,
            "NoticePeriodDays": {"Service Agreement": random.choice([60, 90]),
                                 "Consultancy Agreement": random.choice([14, 30])}.get(atype),
            "Scope": random.choice(SCOPE[atype]),
            "OurSignatory": signer[0],
            "OurSignatoryTitle": signer[1],
            "AccountOwner": random.choice(OWNERS),
        }
        fx = {"NOK": 1, "SEK": 1, "DKK": 0.65, "EUR": 0.087}[currency]
        if atype == "Service Agreement":
            row["ContractValue"] = round(random.randint(120, 900) * 1000 * fx, -3)
        elif atype == "Consultancy Agreement":
            rate = round(random.choice([1250, 1350, 1450, 1550]) * fx, 0)
            row["HourlyRate"] = rate
            row["ContractValue"] = round(rate * random.choice([80, 120, 160, 240, 320]), -3)
        rows.append(row)

    # Two deliberately broken rows so validation has something to catch
    no_rows = [r for r in rows if r["CounterpartyCountry"] == "Norway"]
    bad = no_rows[3]
    digits = bad["CounterpartyOrgNo"].replace(" ", "")
    last = (int(digits[-1]) + 1) % 10
    bad["CounterpartyOrgNo"] = f"{digits[:3]} {digits[3:6]} {digits[6:8]}{last}"
    rows[11]["ContactEmail"] = rows[11]["ContactEmail"].split("@")[0] + "@"
    return rows


AUDIT_COLS = ["DocumentPath", "SignedDocumentPath", "CreatedAt", "SentAt", "SignedAt", "ProcessedAt", "LastError"]

thin = Side(style="thin", color=TEAL50)
HDR_FILL = PatternFill("solid", fgColor=TEAL)
AUDIT_FILL = PatternFill("solid", fgColor=SLATE)
HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
BODY = Font(name=FONT, size=10, color=INK)


def style_header(ws, row, ncols, audit_from=None):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = AUDIT_FILL if audit_from and c >= audit_from else HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=False)
    ws.row_dimensions[row].height = 28


def add_table(ws, ref, name):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(t)


# Uniform grid used on every sheet
COL_W = 20          # every column has the same width
ROW_H = 22          # every data row has the same height
NAV_ROW, ACCENT_ROW, BAND_ROW, HINT_ROW, HDR_ROW = 1, 2, 3, 4, 5
GRID_COLS = 12      # columns styled on sheets with few columns, so every sheet looks the same width

# Sheet -> (icon glyph in "Segoe MDL2 Assets", tooltip). The font ships with Windows 10/11.
NAV = [
    ("Dashboard", "\uE80F", "Dashboard"),      # Home
    ("Agreements", "\uE8FD", "Agreements"),    # BulletedList
    ("Templates", "\uE8A5", "Templates"),      # Document
    ("Settings", "\uE713", "Settings"),        # Setting
    ("Log", "\uE81C", "Audit log"),            # History
    ("How to use", "\uE897", "How to use"),    # Help
]
NAV_FIRST_COL = 5   # icons sit in columns E..J on every sheet
ICON_FONT = "Segoe MDL2 Assets"


def add_nav(ws, title, hint, ncols):
    """Same frame on every sheet: navy nav bar with icon links (row 1), orange accent (row 2),
    page title (row 3), hint (row 4). The icons are cell hyperlinks, so they work without macros."""
    ncols = max(ncols, GRID_COLS)
    for c in range(1, ncols + 1):
        ws.column_dimensions[get_column_letter(c)].width = COL_W
        ws.cell(row=NAV_ROW, column=c).fill = PatternFill("solid", fgColor=NAVY)
        ws.cell(row=ACCENT_ROW, column=c).fill = PatternFill("solid", fgColor=ORANGE)
    ws.row_dimensions[NAV_ROW].height = 36
    ws.row_dimensions[ACCENT_ROW].height = 4
    ws.row_dimensions[BAND_ROW].height = 40
    ws.row_dimensions[HINT_ROW].height = ROW_H
    ws.sheet_view.showGridLines = False

    name = ws.cell(row=NAV_ROW, column=1, value="\u03a3  Agreement Generator")
    name.font = Font(name=FONT, size=13, bold=True, color="FFFFFF")
    name.alignment = Alignment(vertical="center", indent=1)
    for i, (sheet, glyph, tip) in enumerate(NAV):
        cell = ws.cell(row=NAV_ROW, column=NAV_FIRST_COL + i, value=glyph)
        cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!A1", tooltip=tip, display=glyph)
        active = sheet == ws.title
        cell.font = Font(name=ICON_FONT, size=16, color=GOLD if active else "FFFFFF", underline=None)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if active:
            cell.border = Border(bottom=Side(style="thick", color=ORANGE))
    dc = ws.cell(row=NAV_ROW, column=NAV_FIRST_COL + len(NAV) + 1, value="=TODAY()")
    dc.number_format = "dd.mm.yyyy"
    dc.font = Font(name=FONT, size=10, color=SAND)
    dc.alignment = Alignment(horizontal="right", vertical="center")

    if title:
        t = ws.cell(row=BAND_ROW, column=1, value=title)
        t.font = Font(name=FONT, size=18, bold=True, color=NAVY)
        t.alignment = Alignment(vertical="center")
    h = ws.cell(row=HINT_ROW, column=1, value=hint)
    h.font = Font(name=FONT, size=9, italic=True, color=MUTED)
    h.alignment = Alignment(vertical="center")


def style_body(ws, first_row, last_row, ncols):
    line = Border(bottom=Side(style="thin", color=TEAL50))
    for r in range(first_row, last_row + 1):
        ws.row_dimensions[r].height = ROW_H
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = line
            cell.alignment = Alignment(vertical="center", horizontal=cell.alignment.horizontal)


def main():
    rows = build_rows()
    wb = Workbook()

    # ---------- Agreements ----------
    ws = wb.active
    ws.title = "Agreements"
    ws.sheet_properties.tabColor = TEAL
    HR = HDR_ROW
    data_cols = list(rows[0].keys())
    cols = data_cols + AUDIT_COLS
    for c, h in enumerate(cols, 1):
        ws.cell(row=HR, column=c, value=h)
    style_header(ws, HR, len(cols), audit_from=len(data_cols) + 1)
    idx = {h: i + 1 for i, h in enumerate(cols)}
    for r, row in enumerate(rows, HR + 1):
        for h, v in row.items():
            cell = ws.cell(row=r, column=idx[h], value=v)
            cell.font = BODY
        s, t = get_column_letter(idx["StartDate"]), get_column_letter(idx["TermMonths"])
        ws.cell(row=r, column=idx["EndDate"],
                value=f'=IF(OR({s}{r}="",{t}{r}=""),"",EDATE({s}{r},{t}{r})-1)').font = BODY
    last = HR + len(rows)
    for h in ("StartDate", "EndDate"):
        for r in range(HR + 1, last + 1):
            ws.cell(row=r, column=idx[h]).number_format = "dd.mm.yyyy"
    for h in ("ContractValue", "HourlyRate"):
        for r in range(HR + 1, last + 1):
            ws.cell(row=r, column=idx[h]).number_format = "#,##0"
    for h in ("CreatedAt", "SentAt", "SignedAt", "ProcessedAt"):
        for r in range(HR + 1, last + 1):
            ws.cell(row=r, column=idx[h]).number_format = "yyyy-mm-dd hh:mm"
    for h in ("CounterpartyPostcode", "CounterpartyOrgNo"):
        for r in range(HR + 1, last + 1):
            ws.cell(row=r, column=idx[h]).number_format = "@"
    style_body(ws, HR + 1, last, len(cols))
    add_table(ws, f"A{HR}:{get_column_letter(len(cols))}{last}", "tblAgreements")

    ws.freeze_panes = ws.cell(row=HR + 1, column=idx["CounterpartyName"])

    st_col = get_column_letter(idx["Status"])
    rng = f"{st_col}{HR+1}:{st_col}{last}"
    for status, fill, font in (
        ("Not created", "ECE8DA", INK), ("Created", SAND, INK), ("Sent for signing", GOLD, INK),
        ("Signed", ORANGE, INK), ("Processed", NAVY, "FFFFFF"),
    ):
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'${st_col}{HR+1}="{status}"'], fill=PatternFill("solid", fgColor=fill),
            font=Font(color=font, bold=True)))
    err_col = get_column_letter(idx["LastError"])
    ws.conditional_formatting.add(f"A{HR+1}:{get_column_letter(len(cols))}{last}", FormulaRule(
        formula=[f'${err_col}{HR+1}<>""'], fill=PatternFill("solid", fgColor="FCE3D2")))

    dv_status = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=False,
                               showErrorMessage=True, errorTitle="Status",
                               error="Use the buttons to change status, so timestamps and the log stay correct.")
    dv_type = DataValidation(type="list", formula1="=Templates!$A$6:$A$30", allow_blank=False)
    dv_cur = DataValidation(type="list", formula1='"NOK,SEK,DKK,EUR"', allow_blank=True)
    for dv, h in ((dv_status, "Status"), (dv_type, "AgreementType"), (dv_cur, "Currency")):
        col = get_column_letter(idx[h])
        dv.add(f"{col}{HR+1}:{col}{HR+500}")
        ws.add_data_validation(dv)

    # ---------- Settings ----------
    s = wb.create_sheet("Settings")
    s.sheet_properties.tabColor = SLATE
    settings = [
        ("OurCompanyName", "Nordvik Analytics AS", "Your company (fictional)"),
        ("OurOrgNo", "998 877 664", "Fictional"),
        ("OurAddress", "Dronning Eufemias gate 16", ""),
        ("OurPostcode", "0191", ""),
        ("OurCity", "Oslo", ""),
        ("GoverningLaw", "Norwegian law", ""),
        ("Venue", "Oslo District Court", ""),
        ("TemplateFolder", "", "Blank = .\\templates next to this workbook"),
        ("OutputFolder", "", "Blank = .\\output next to this workbook"),
        ("DateFormat", "dd.mm.yyyy", "VBA Format() pattern for dates in documents"),
        ("NumberFormat", "#,##0", "VBA Format() pattern for numbers in documents"),
        ("KeepDocx", "Yes", "Also save the filled .docx next to the PDF"),
        ("PdfA", "Yes", "Export as PDF/A (ISO 19005-1) for archiving"),
        ("OverwriteExisting", "No", "Regenerate even if the PDF already exists"),
        ("CreateOutlookDrafts", "No", "Yes = Send for signing creates Outlook drafts with the PDF attached"),
        ("ValidateNorwegianOrgNo", "Yes", "MOD11 check on Norwegian organisation numbers"),
    ]
    for c, h in enumerate(("Key", "Value", "Note"), 1):
        s.cell(row=HDR_ROW, column=c, value=h)
    style_header(s, HDR_ROW, 3)
    for r, (k, v, n) in enumerate(settings, HDR_ROW + 1):
        s.cell(row=r, column=1, value=k).font = BODY
        vc = s.cell(row=r, column=2, value=v)
        vc.font = Font(name=FONT, size=10, bold=True, color=NAVY)
        vc.fill = PatternFill("solid", fgColor="FBF3DF")   # light gold = input cell
        vc.number_format = "@"
        s.cell(row=r, column=3, value=n).font = Font(name=FONT, size=9, color=MUTED)
    last_s = HDR_ROW + len(settings)
    style_body(s, HDR_ROW + 1, last_s, 3)
    add_table(s, f"A{HDR_ROW}:C{last_s}", "tblSettings")

    # ---------- Templates ----------
    t = wb.create_sheet("Templates")
    t.sheet_properties.tabColor = SAGE
    for c, h in enumerate(("AgreementType", "TemplateFile", "FilePrefix", "TemplateVersion"), 1):
        t.cell(row=HDR_ROW, column=c, value=h)
    style_header(t, HDR_ROW, 4)
    for r, row in enumerate(TYPES, HDR_ROW + 1):
        for c, v in enumerate(row, 1):
            cell = t.cell(row=r, column=c, value=v); cell.font = BODY; cell.number_format = "@"
    last_t = HDR_ROW + len(TYPES)
    style_body(t, HDR_ROW + 1, last_t, 4)
    add_table(t, f"A{HDR_ROW}:D{last_t}", "tblTemplates")

    # ---------- Log ----------
    lg = wb.create_sheet("Log")
    lg.sheet_properties.tabColor = MUTED
    hdr = ("Timestamp", "AgreementID", "Action", "FromStatus", "ToStatus", "User", "Message")
    for c, h in enumerate(hdr, 1):
        lg.cell(row=HDR_ROW, column=c, value=h)
    style_header(lg, HDR_ROW, len(hdr))
    style_body(lg, HDR_ROW + 1, HDR_ROW + 1, len(hdr))
    add_table(lg, f"A{HDR_ROW}:G{HDR_ROW + 1}", "tblLog")

    # ---------- Dashboard ----------
    d = wb.create_sheet("Dashboard", 0)
    d.sheet_properties.tabColor = RED
    stc = f"Agreements!${st_col}:${st_col}"
    typ = f"Agreements!${get_column_letter(idx['AgreementType'])}:${get_column_letter(idx['AgreementType'])}"
    val = f"Agreements!${get_column_letter(idx['ContractValue'])}:${get_column_letter(idx['ContractValue'])}"
    errc = f"Agreements!${err_col}${HR+1}:${err_col}$1000"
    idc = f"Agreements!$A${HR+1}:$A$1000"
    for c, h in enumerate(("Status", "Agreements", "Contract value"), 1):
        d.cell(row=HDR_ROW, column=c, value=h)
    style_header(d, HDR_ROW, 3)
    r1 = HDR_ROW + 1
    for r, stt in enumerate(STATUSES, r1):
        d.cell(row=r, column=1, value=stt)
        d.cell(row=r, column=2, value=f'=COUNTIF({stc},A{r})')
        d.cell(row=r, column=3, value=f'=SUMIF({stc},A{r},{val})').number_format = "#,##0"
    tr = r1 + len(STATUSES)
    d.cell(row=tr, column=1, value="Total")
    d.cell(row=tr, column=2, value=f"=SUM(B{r1}:B{tr-1})")
    d.cell(row=tr, column=3, value=f"=SUM(C{r1}:C{tr-1})").number_format = "#,##0"
    d.cell(row=tr + 1, column=1, value="Rows with errors")
    d.cell(row=tr + 1, column=2, value=f'=SUMPRODUCT(({errc}<>"")*({idc}<>""))')
    d.conditional_formatting.add(f"B{tr+1}", FormulaRule(formula=[f"B{tr+1}>0"],
                                 fill=PatternFill("solid", fgColor=ORANGE)))
    for r in range(r1, tr + 2):
        for c in range(1, 4):
            d.cell(row=r, column=c).font = Font(name=FONT, size=10, bold=(r >= tr), color=INK)
    style_body(d, r1, tr + 1, 3)

    r0 = tr + 3
    for c, h in enumerate(("Agreement type", "Agreements", "Signed or processed"), 1):
        d.cell(row=r0, column=c, value=h)
    style_header(d, r0, 3)
    for i, (atype, *_r) in enumerate(TYPES, 1):
        r = r0 + i
        d.cell(row=r, column=1, value=atype)
        d.cell(row=r, column=2, value=f'=COUNTIF({typ},A{r})')
        d.cell(row=r, column=3, value=f'=COUNTIFS({typ},A{r},{stc},"Signed")+COUNTIFS({typ},A{r},{stc},"Processed")')
        for c in range(1, 4):
            d.cell(row=r, column=c).font = Font(name=FONT, size=10, color=INK)
    style_body(d, r0 + 1, r0 + len(TYPES), 3)

    ch = BarChart(); ch.type = "bar"; ch.style = 10
    ch.title = "Agreements by status"; ch.legend = None
    ch.add_data(Reference(d, min_col=2, min_row=HDR_ROW, max_row=HDR_ROW + len(STATUSES)), titles_from_data=True)
    ch.set_categories(Reference(d, min_col=1, min_row=r1, max_row=HDR_ROW + len(STATUSES)))
    ch.series[0].graphicalProperties.solidFill = NAVY
    ch.series[0].graphicalProperties.line.solidFill = NAVY
    ch.y_axis.majorGridlines = None
    ch.x_axis.scaling.orientation = "maxMin"
    ch.y_axis.delete = False; ch.x_axis.delete = False
    ch.height, ch.width = 7.6, 14.5
    d.add_chart(ch, f"E{HDR_ROW}")

    # ---------- Instructions ----------
    ins = wb.create_sheet("How to use")
    ins.sheet_properties.tabColor = BLUE
    lines = [
        ("One-time setup", 12, True, TEAL),
        ("1. Save this workbook as AgreementGenerator.xlsm (Excel Macro-Enabled Workbook) in the repo folder.", 10, False, INK),
        ("2. Alt+F11 > File > Import File... and import the four .bas files from src\\vba.", 10, False, INK),
        ("3. Run the macro Setup (Alt+F8). It creates the output folders and draws the buttons on the Agreements sheet.", 10, False, INK),
        ("", 10, False, INK),
        ("Daily flow (each button acts on selected rows, or on every eligible row if nothing in the table is selected)", 12, True, TEAL),
        ("Validate: checks required fields, e-mail, duplicate IDs, template files and Norwegian org numbers (MOD11). Errors go to LastError.", 10, False, INK),
        ("Generate documents: Not created > Created. Fills the template, saves PDF (and .docx) to output\\Generated.", 10, False, INK),
        ("Send for signing: Created > Sent for signing. Optionally creates an Outlook draft per agreement with the PDF attached.", 10, False, INK),
        ("Import signed: scans output\\Signed for files containing the AgreementID and moves matching rows to Signed.", 10, False, INK),
        ("Mark signed: manual alternative for selected rows (e.g. signed on paper).", 10, False, INK),
        ("Process: Signed > Processed. Moves the signed file to output\\Archive\\<year>.", 10, False, INK),
        ("Reset: puts selected rows back to Not created and renames the old PDF *_VOID_<timestamp>. Use after fixing data.", 10, False, INK),
        ("Demo: simulate signing: copies the PDFs of rows in Sent for signing into output\\Signed, as if they came back from e-signing.", 10, False, INK),
        ("", 10, False, INK),
        ("Rules", 12, True, TEAL),
        ("Status only moves one step forward. Change it with the buttons, not the dropdown, so timestamps and the log stay correct.", 10, False, INK),
        ("Placeholders in templates are {{ColumnHeader}} or {{SettingsKey}}, plus {{DocumentDate}} and {{TemplateVersion}}.", 10, False, INK),
        ("A document is never exported if a placeholder is unknown or its value is blank; the row gets an error instead.", 10, False, INK),
        ("Light gold cells on the Settings sheet are inputs. Columns from DocumentPath onwards (dark header) are written by the macros.", 10, False, INK),
        ("Navigation: the icons in the top bar link to the sheets (hover for the name).", 10, False, INK),
        ("All companies, people and organisation numbers are fictional.", 10, False, MUTED),
    ]
    for r, (txt, size, bold, color) in enumerate(lines, HDR_ROW + 1):
        cell = ins.cell(row=r, column=1, value=txt)
        cell.font = Font(name=FONT, size=size, bold=bold, color=color)
        cell.alignment = Alignment(vertical="center")
        ins.row_dimensions[r].height = ROW_H

    add_nav(ws, None, "Fill in one row per agreement. Select rows to act on them, or select nothing "
                      "to process every eligible row.", len(cols))
    add_nav(d, "Agreement status", "Live counts from the Agreements sheet.", 3)
    add_nav(s, "Settings", "Every Key can be used as a {{Key}} placeholder in templates. Edit the light gold cells only.", 3)
    add_nav(t, "Templates", "Maps each AgreementType to a .docx in the template folder. Add a row to add a new agreement type.", 4)
    add_nav(lg, "Audit log", "Written by the macros. One line per action; do not edit.", 7)
    add_nav(ins, "How to use", "Setup, daily flow and rules.", 1)
    for sheet in (d, s, t, lg, ins):
        sheet.freeze_panes = f"A{HDR_ROW + 1}"
    wb.move_sheet("How to use", offset=-(len(wb.sheetnames) - 1))
    wb.active = wb.sheetnames.index("Agreements")
    wb.calculation.fullCalcOnLoad = True  # Excel calculates all formulas when the file opens
    out = ROOT / "AgreementGenerator.xlsx"
    wb.save(out)
    print("Saved", out, len(rows), "rows")


if __name__ == "__main__":
    main()
