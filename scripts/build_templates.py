"""Builds the three Word agreement templates used by the VBA generator.

Placeholders use the syntax {{FieldName}}. FieldName must match a column header
in tblAgreements, a key in tblSettings, or a computed field (DocumentDate,
TemplateVersion). Each placeholder is written in a single Word run so Find can
match it; when you edit a template in Word, type placeholders in one go.
"""
from pathlib import Path
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm

TEAL = RGBColor(0x20, 0x2C, 0x39)   # navy, "Autumn Harvest" palette
MUTED = RGBColor(0x4A, 0x58, 0x66)
INK = RGBColor(0x20, 0x2C, 0x39)
OUT = Path(__file__).resolve().parents[1] / "templates"


def base_doc(title):
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(2.3)
    sec.top_margin, sec.bottom_margin = Cm(2.2), Cm(2.0)

    st = doc.styles["Normal"]
    st.font.name, st.font.size, st.font.color.rgb = "Calibri", Pt(10.5), INK
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.15
    for name, size in (("Heading 1", 11.5), ("Title", 20)):
        h = doc.styles[name]
        h.font.name, h.font.size, h.font.bold, h.font.color.rgb = "Calibri", Pt(size), True, TEAL
        rf = h.element.rPr.rFonts
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            rf.attrib.pop(qn(a), None)
        # Title style ships with a bottom border; we draw our own rule
        pPr = h.element.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:pBdr")) is not None:
            pPr.remove(pPr.find(qn("w:pBdr")))
    doc.styles["Heading 1"].paragraph_format.space_before = Pt(12)
    doc.styles["Heading 1"].paragraph_format.space_after = Pt(4)

    # Header / footer placeholders prove that every story range is processed
    hp = sec.header.paragraphs[0]
    r = hp.add_run("{{OurCompanyName}}  |  {{AgreementID}}")
    r.font.size, r.font.color.rgb = Pt(8.5), MUTED
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    fp = sec.footer.paragraphs[0]
    r = fp.add_run("{{AgreementType}}  ·  Template {{TemplateVersion}}  ·  Generated {{DocumentDate}}  ·  Page ")
    r.font.size, r.font.color.rgb = Pt(8), MUTED
    add_page_field(fp)

    t = doc.add_paragraph(title, style="Title")
    t.paragraph_format.space_after = Pt(2)
    sub = doc.add_paragraph()
    r = sub.add_run("Agreement no. {{AgreementID}}  ·  Dated {{DocumentDate}}")
    r.font.color.rgb, r.font.size = MUTED, Pt(10)
    rule(sub)
    return doc


def add_page_field(p):
    r = p.add_run()
    r.font.size, r.font.color.rgb = Pt(8), MUTED
    for tag, txt in (("begin", None), (None, " PAGE "), ("separate", None), ("text", "1"), ("end", None)):
        if tag == "text":
            el = OxmlElement("w:t"); el.text = txt
        elif tag:
            el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), tag)
        else:
            el = OxmlElement("w:instrText"); el.set(qn("xml:space"), "preserve"); el.text = txt
        r._r.append(el)


def rule(p):
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "8"), ("w:space", "6"), ("w:color", "F29559")):
        b.set(qn(k), v)
    bdr.append(b); pPr.append(bdr)


def parties(doc, our_role, their_role):
    doc.add_paragraph("This agreement is entered into between:")
    for name, org, addr, pc, city, country, role in (
        ("{{OurCompanyName}}", "{{OurOrgNo}}", "{{OurAddress}}", "{{OurPostcode}}", "{{OurCity}}", "Norway", our_role),
        ("{{CounterpartyName}}", "{{CounterpartyOrgNo}}", "{{CounterpartyAddress}}", "{{CounterpartyPostcode}}",
         "{{CounterpartyCity}}", "{{CounterpartyCountry}}", their_role),
    ):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.add_run(name).bold = True
        p.add_run(f", org. no. {org}, {addr}, {pc} {city}, {country} (the “{role}”)")
    doc.add_paragraph("each a “Party” and together the “Parties”.")


def clause(doc, heading, *paras):
    doc.add_paragraph(heading, style="Heading 1")
    for txt in paras:
        doc.add_paragraph(txt)


def signatures(doc, left_role, right_role):
    doc.add_paragraph("Signatures", style="Heading 1")
    intro = doc.add_paragraph("This agreement is signed electronically in two (2) originals, one for each Party.")
    intro.paragraph_format.keep_with_next = True
    tbl = doc.add_table(rows=1, cols=2)
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))  # keep the signature block on one page
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    blocks = (
        (left_role, "{{OurCompanyName}}", "{{OurSignatory}}", "{{OurSignatoryTitle}}"),
        (right_role, "{{CounterpartyName}}", "{{ContactName}}", "{{ContactTitle}}"),
    )
    for cell, (role, comp, name, title) in zip(tbl.rows[0].cells, blocks):
        cell.paragraphs[0].add_run(f"For the {role}").font.color.rgb = MUTED
        cell.add_paragraph().add_run(comp).bold = True
        cell.add_paragraph("\n\n______________________________")
        cell.add_paragraph(name)
        cell.add_paragraph(title).runs[0].font.color.rgb = MUTED
        cell.add_paragraph("Place/date: ____________________")


def service_agreement():
    doc = base_doc("Service Agreement")
    parties(doc, "Supplier", "Customer")
    clause(doc, "1. Scope of services",
           "The Supplier shall deliver the following services to the Customer (the “Services”): {{Scope}}",
           "The Services shall be performed with due care and in accordance with good industry practice.")
    clause(doc, "2. Term",
           "This agreement takes effect on {{StartDate}} and runs for {{TermMonths}} months until {{EndDate}}. "
           "Thereafter it renews automatically for successive twelve (12) month periods unless terminated "
           "by either Party with {{NoticePeriodDays}} days’ written notice.")
    clause(doc, "3. Fees and payment",
           "The annual fee for the Services is {{Currency}} {{ContractValue}} excluding VAT, invoiced quarterly in advance.",
           "Invoices are payable within {{PaymentTermsDays}} days of the invoice date. Late payment interest accrues "
           "in accordance with the Norwegian Late Payment Interest Act.")
    clause(doc, "4. Customer responsibilities",
           "The Customer shall provide timely access to information, systems and personnel reasonably required "
           "for the Supplier to perform the Services. {{ContactName}} ({{ContactEmail}}) is the Customer’s "
           "primary contact.")
    clause(doc, "5. Confidentiality",
           "Each Party shall keep confidential all non-public information received from the other Party and use it "
           "only for the purpose of this agreement. This obligation survives termination for three (3) years.")
    clause(doc, "6. Limitation of liability",
           "Neither Party is liable for indirect or consequential loss. Each Party’s total liability under this "
           "agreement is limited to the fees paid in the twelve (12) months preceding the claim, except in cases of "
           "gross negligence or wilful misconduct.")
    clause(doc, "7. Governing law and disputes",
           "This agreement is governed by {{GoverningLaw}}. Disputes that cannot be resolved amicably shall be "
           "settled by the ordinary courts, with {{Venue}} as agreed venue.")
    signatures(doc, "Supplier", "Customer")
    doc.save(OUT / "Service_Agreement.docx")


def mutual_nda():
    doc = base_doc("Mutual Non-Disclosure Agreement")
    parties(doc, "First Party", "Second Party")
    clause(doc, "1. Purpose",
           "The Parties wish to exchange confidential information for the following purpose (the “Purpose”): {{Scope}}")
    clause(doc, "2. Confidential information",
           "“Confidential Information” means any information, in any form, disclosed by one Party to the other "
           "in connection with the Purpose that is marked as confidential or should reasonably be understood to be "
           "confidential, including business plans, customer data, pricing, source code and know-how.")
    clause(doc, "3. Obligations",
           "The receiving Party shall (a) use the Confidential Information only for the Purpose, (b) not disclose it to "
           "third parties without prior written consent, and (c) protect it with at least the same degree of care it "
           "uses for its own confidential information, and no less than reasonable care.")
    clause(doc, "4. Exceptions",
           "The obligations do not apply to information that is or becomes publicly available through no fault of the "
           "receiving Party, was lawfully known before disclosure, is independently developed, or must be disclosed "
           "by law or court order.")
    clause(doc, "5. Term",
           "This agreement applies to information disclosed from {{StartDate}} for a period of {{TermMonths}} months. "
           "The confidentiality obligations survive for five (5) years after the end of that period.")
    clause(doc, "6. Return of information",
           "On request, the receiving Party shall promptly return or delete all Confidential Information, except "
           "copies it is required to keep by law.")
    clause(doc, "7. Governing law and disputes",
           "This agreement is governed by {{GoverningLaw}}, with {{Venue}} as agreed venue.")
    signatures(doc, "First Party", "Second Party")
    doc.save(OUT / "Mutual_NDA.docx")


def consultancy_agreement():
    doc = base_doc("Consultancy Agreement")
    parties(doc, "Consultant", "Client")
    clause(doc, "1. Assignment",
           "The Consultant shall carry out the following assignment for the Client (the “Assignment”): {{Scope}}",
           "The Consultant works independently and is free to decide how the Assignment is carried out, within the "
           "Client’s reasonable instructions.")
    clause(doc, "2. Period",
           "The Assignment starts on {{StartDate}} and is expected to be completed by {{EndDate}}. Either Party may "
           "terminate the Assignment with {{NoticePeriodDays}} days’ written notice.")
    clause(doc, "3. Fees",
           "The Consultant invoices hours worked at an hourly rate of {{Currency}} {{HourlyRate}} excluding VAT. "
           "Total fees shall not exceed {{Currency}} {{ContractValue}} without the Client’s prior written approval.",
           "The Consultant invoices monthly in arrears. Invoices are payable within {{PaymentTermsDays}} days.")
    clause(doc, "4. Intellectual property",
           "On payment, ownership of deliverables produced specifically for the Client under this agreement passes to "
           "the Client. The Consultant retains its pre-existing tools, methods and know-how.")
    clause(doc, "5. Confidentiality",
           "The Consultant shall keep confidential all information about the Client’s business received during the "
           "Assignment, during and after the term of this agreement.")
    clause(doc, "6. Governing law and disputes",
           "This agreement is governed by {{GoverningLaw}}, with {{Venue}} as agreed venue.")
    signatures(doc, "Consultant", "Client")
    doc.save(OUT / "Consultancy_Agreement.docx")


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    service_agreement(); mutual_nda(); consultancy_agreement()
    print("Templates written to", OUT)
