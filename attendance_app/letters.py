# letters.py
"""
AttendPro — Letter Generation Engine
Builds Employee Request Letters (Salary Hike Request, Leave Request) and
HR Issued Letters (Hike Approval, Experience Letter, NOC) as both
.docx (python-docx) and .pdf (reportlab), from the same content spec so
both files always match.

Design:
  - Employee Request Letters: NO company letterhead (the employee is writing
    TO the company), clean Ref/Date bar, To/From blocks, a shaded details
    box for structured data, a reason block, and a simple signature line.
  - HR Issued Letters: full company letterhead (the company is writing TO
    the employee), same modern details-box treatment for figures, and a
    "For Company" signature block.
"""

import io
import re
import datetime


# ─────────────────────────────────────────────────────────────────────────────
# Shared helpers
# ─────────────────────────────────────────────────────────────────────────────

LIGHT_BORDER_HEX = 'D9DEE7'
TINT_HEX = 'FFF3E8'
BOX_TINT_HEX = 'F1F5F9'
BOX_BORDER_HEX = 'E2E8F0'


def fmt_date(d, fmt='%d %B %Y'):
    if not d:
        return '—'
    return d.strftime(fmt)


def fmt_money(v, currency='AED'):
    if v is None or v == '':
        return '—'
    return f"{currency} {float(v):,.2f}"


def parse_date(s):
    """Parse an ISO date string or date/datetime object -> date or None."""
    if not s:
        return None
    if isinstance(s, datetime.datetime):
        return s.date()
    if isinstance(s, datetime.date):
        return s
    try:
        return datetime.date.fromisoformat(str(s))
    except (ValueError, TypeError):
        return None


def parse_inline(text):
    """Split a string on **bold** markers -> list of (segment_text, is_bold)."""
    parts = re.split(r'(\*\*.*?\*\*)', text)
    out = []
    for p in parts:
        if not p:
            continue
        if p.startswith('**') and p.endswith('**'):
            out.append((p[2:-2], True))
        else:
            out.append((p, False))
    return out


def number_to_words_int(n):
    ones = ['', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine',
            'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen',
            'seventeen', 'eighteen', 'nineteen']
    tens = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety']

    def under_thousand(num):
        if num < 20:
            return ones[num]
        if num < 100:
            return tens[num // 10] + (f"-{ones[num % 10]}" if num % 10 else '')
        return ones[num // 100] + " hundred" + (f" {under_thousand(num % 100)}" if num % 100 else '')

    if n == 0:
        return 'zero'
    parts = []
    for div, name in [(1_000_000_000, 'billion'), (1_000_000, 'million'), (1_000, 'thousand')]:
        if n >= div:
            parts.append(f"{under_thousand(n // div)} {name}")
            n %= div
    if n:
        parts.append(under_thousand(n))
    return ' '.join(parts)


def years_months_between(d1, d2):
    if not d1:
        return (0, 0)
    months = (d2.year - d1.year) * 12 + (d2.month - d1.month)
    if d2.day < d1.day:
        months -= 1
    months = max(months, 0)
    return months // 12, months % 12


def service_duration_phrase(joining_date, as_of=None):
    as_of = as_of or datetime.date.today()
    years, months = years_months_between(joining_date, as_of)
    parts = []
    if years:
        parts.append(f"{number_to_words_int(years)} ({years}) year{'s' if years != 1 else ''}")
    if months:
        parts.append(f"{number_to_words_int(months)} ({months}) month{'s' if months != 1 else ''}")
    if not parts:
        return "less than a month"
    return ' and '.join(parts)


# ─────────────────────────────────────────────────────────────────────────────
# DOCX builder — modern design
# ─────────────────────────────────────────────────────────────────────────────

def build_letter_docx(ctx):
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    cs = ctx['company']
    DARK = RGBColor(0x1E, 0x29, 0x3B)
    ORANGE = RGBColor(0xE8, 0x65, 0x0A)
    GREY = RGBColor(0x64, 0x74, 0x8B)

    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10.5)
    style.font.color.rgb = DARK
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), 'Calibri')

    def set_cell_bg(cell, hex_color):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), hex_color)
        tcPr.append(shd)

    def set_table_borders(table, color=LIGHT_BORDER_HEX, sz=6, sides=('top', 'left', 'bottom', 'right', 'insideH', 'insideV')):
        tbl = table._tbl
        tblPr = tbl.tblPr
        borders = OxmlElement('w:tblBorders')
        for edge in sides:
            el = OxmlElement(f'w:{edge}')
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), color)
            borders.append(el)
        tblPr.append(borders)

    def set_cell_margins(cell, top=90, bottom=90, left=140, right=140):
        tcPr = cell._tc.get_or_add_tcPr()
        mar = OxmlElement('w:tcMar')
        for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
            node = OxmlElement(f'w:{side}')
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            mar.append(node)
        tcPr.append(mar)

    def hr_rule(color_rgb=(0xE8, 0x65, 0x0A), size_pt=2.2, space_after=14, space_before=0):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.space_before = Pt(space_before)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), str(int(size_pt * 8)))
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), '%02X%02X%02X' % color_rgb)
        pBdr.append(bottom)
        pPr.append(pBdr)
        return p

    def add_para(text='', bold=False, size=10.5, align=None, space_after=8, space_before=0,
                 italic=False, color=None, justify=False, underline=False, letter_spacing=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.line_spacing = 1.32
        if justify:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        elif align is not None:
            p.alignment = align
        if text:
            for seg, seg_bold in parse_inline(text):
                run = p.add_run(seg)
                run.bold = bold or seg_bold
                run.italic = italic
                run.underline = underline
                run.font.size = Pt(size)
                run.font.color.rgb = color if color else DARK
        return p

    # ── LETTERHEAD (HR letters only) ─────────────────────────────────────────
    if ctx.get('letterhead'):
        add_para(cs.company_name, bold=True, size=17, align=WD_ALIGN_PARAGRAPH.CENTER,
                  space_after=3, color=DARK)
        add_para(cs.company_address.replace('\n', ', '), size=9, align=WD_ALIGN_PARAGRAPH.CENTER,
                  space_after=1, color=GREY)
        contact_bits = []
        if cs.company_phone:
            contact_bits.append(f"Tel: {cs.company_phone}")
        if cs.company_email:
            contact_bits.append(f"Email: {cs.company_email}")
        if cs.company_trn:
            contact_bits.append(f"TRN: {cs.company_trn}")
        if contact_bits:
            add_para('   |   '.join(contact_bits), size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER,
                      space_after=8, color=GREY)
        hr_rule(space_after=16)

    # ── REF / DATE BAR ────────────────────────────────────────────────────────
    date_p = add_para('', space_after=2)
    date_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1 = date_p.add_run('Date: ')
    r1.bold = True; r1.font.size = Pt(9); r1.font.color.rgb = GREY
    r2 = date_p.add_run(fmt_date(ctx['date']))
    r2.font.size = Pt(9); r2.font.color.rgb = GREY
    hr_rule(color_rgb=(0xCB, 0xD5, 0xE1), size_pt=0.75, space_after=2, space_before=2)

    # ── TITLE ────────────────────────────────────────────────────────────────
    add_para(ctx['title'].upper(), bold=True, size=15, align=WD_ALIGN_PARAGRAPH.CENTER,
              space_after=4, space_before=10, color=ORANGE if ctx.get('letterhead') else DARK)
    underline_p = doc.add_paragraph()
    underline_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    underline_p.paragraph_format.space_after = Pt(16)
    pPr = underline_p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '10')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '%02X%02X%02X' % ((0xE8, 0x65, 0x0A) if ctx.get('letterhead') else (0x1E, 0x29, 0x3B)))
    pBdr.append(bottom)
    pPr.append(pBdr)
    underline_p.add_run(' ' * 18)

    # ── TO / FROM BLOCKS ─────────────────────────────────────────────────────
    def block(label, lines, bold_first=1):
        add_para(f"{label},", bold=True, size=9.5, space_after=3, color=GREY)
        for i, line in enumerate(lines):
            add_para(line, bold=(i < bold_first), size=10.5, space_after=1)
        add_para('', space_after=6)

    if ctx.get('to_block'):
        block('To', ctx['to_block'], bold_first=1)
    if ctx.get('from_block'):
        block('From', ctx['from_block'], bold_first=1)

    # ── SUBJECT ──────────────────────────────────────────────────────────────
    sub_p = add_para(f"**Subject: {ctx['subject']}**", size=10.5, space_after=2, color=DARK)
    hr_rule(color_rgb=(0xCB, 0xD5, 0xE1), size_pt=0.75, space_after=12)

    # ── SALUTATION ───────────────────────────────────────────────────────────
    if ctx.get('salutation'):
        add_para(ctx['salutation'], size=10.5, space_after=10)

    # ── INTRO ────────────────────────────────────────────────────────────────
    for para in ctx.get('intro', []):
        add_para(para, size=10.5, space_after=12, justify=True)

    # ── DETAILS BOX ──────────────────────────────────────────────────────────
    if ctx.get('details_box'):
        db = ctx['details_box']
        add_para(db['title'].upper(), bold=True, size=9, space_after=6, space_before=2, color=ORANGE)
        rows = db['rows']
        tbl = doc.add_table(rows=len(rows), cols=2)
        tbl.autofit = True
        set_table_borders(tbl, color=BOX_BORDER_HEX, sz=4, sides=('top', 'left', 'bottom', 'right', 'insideH'))
        for ri, (label, value) in enumerate(rows):
            lcell, vcell = tbl.rows[ri].cells
            set_cell_margins(lcell, top=70, bottom=70)
            set_cell_margins(vcell, top=70, bottom=70)
            set_cell_bg(lcell, BOX_TINT_HEX)
            set_cell_bg(vcell, BOX_TINT_HEX)
            lcell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            vcell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            lp2 = lcell.paragraphs[0]
            run = lp2.add_run(label)
            run.font.size = Pt(9.5); run.font.color.rgb = GREY
            vp2 = vcell.paragraphs[0]
            run = vp2.add_run(str(value))
            run.bold = True; run.font.size = Pt(9.5); run.font.color.rgb = DARK
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # ── REASON / NOTE BLOCK ──────────────────────────────────────────────────
    if ctx.get('reason_block'):
        rb = ctx['reason_block']
        add_para(rb['title'].upper(), bold=True, size=9, space_after=4, space_before=2, color=ORANGE)
        add_para(rb['text'], size=10.5, space_after=12, justify=True)

    # ── BODY (closing paragraphs) ────────────────────────────────────────────
    for para in ctx['body']:
        add_para(para, size=10.5, space_after=12, justify=True)

    # ── SIGNATURE ────────────────────────────────────────────────────────────
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    if ctx['signature_mode'] == 'employee_request':
        add_para(ctx.get('closing_phrase', 'Yours sincerely,'), size=10.5, space_after=40)
        hr_rule(color_rgb=(0x33, 0x33, 0x33), size_pt=0.75, space_after=3)
        sig = ctx['signature']
        add_para(sig['name'], bold=True, size=10.5, space_after=1)
        for line in sig.get('lines', []):
            add_para(line, size=9.5, space_after=1, color=GREY)

    elif ctx['signature_mode'] == 'warning':
        add_para(f"For **{cs.company_name}**", size=10.5, space_after=30, color=DARK)
        sign_tbl = doc.add_table(rows=2, cols=2)
        set_table_borders(sign_tbl, color=LIGHT_BORDER_HEX, sz=5)
        headers = [ctx.get('signatory_title', 'HR Manager'), 'Employee Acknowledgement (Signature & Date)']
        for ci, h in enumerate(headers):
            cell = sign_tbl.cell(0, ci)
            set_cell_margins(cell, top=80, bottom=80)
            set_cell_bg(cell, TINT_HEX)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(h)
            run.bold = True; run.font.size = Pt(9.5); run.font.color.rgb = ORANGE
        for ci in range(2):
            cell = sign_tbl.cell(1, ci)
            set_cell_margins(cell, top=280, bottom=100)
            cell.paragraphs[0].add_run(' ')

    else:
        add_para(f"For **{cs.company_name}**", size=10.5, space_after=44, color=DARK)
        sign_tbl = doc.add_table(rows=2, cols=3)
        set_table_borders(sign_tbl, color=LIGHT_BORDER_HEX, sz=5)
        headers = [ctx.get('signatory_title', 'HR Manager'), 'Approved By (In-Charge)', 'Approved By (Manager)']
        for ci, h in enumerate(headers):
            cell = sign_tbl.cell(0, ci)
            set_cell_margins(cell, top=80, bottom=80)
            set_cell_bg(cell, TINT_HEX)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(h)
            run.bold = True; run.font.size = Pt(9.5); run.font.color.rgb = ORANGE
        for ci in range(3):
            cell = sign_tbl.cell(1, ci)
            set_cell_margins(cell, top=280, bottom=100)
            cell.paragraphs[0].add_run(' ')

    # ── FOOTER ───────────────────────────────────────────────────────────────
    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = footer_p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '4')
    top.set(qn('w:space'), '6')
    top.set(qn('w:color'), LIGHT_BORDER_HEX)
    pBdr.append(top)
    pPr.append(pBdr)
    footer_text = ctx.get('footer_note') or (
        f"{cs.company_name}  •  {cs.company_address.splitlines()[0] if cs.company_address else ''}"
    )
    run = footer_p.add_run(footer_text)
    run.font.size = Pt(7.5)
    run.font.color.rgb = GREY
    run.italic = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# PDF builder (reportlab) — mirrors the DOCX layout
# ─────────────────────────────────────────────────────────────────────────────

def build_letter_pdf(ctx):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_JUSTIFY
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    from reportlab.lib.styles import ParagraphStyle

    cs = ctx['company']
    ORANGE = colors.HexColor('#E8650A')
    DARK = colors.HexColor('#1E293B')
    GREY = colors.HexColor('#64748B')
    BORDER = colors.HexColor('#D9DEE7')
    BOX_BORDER = colors.HexColor('#E2E8F0')
    BOX_TINT = colors.HexColor('#F1F5F9')
    TINT = colors.HexColor('#FFF3E8')
    TITLE_COLOR = ORANGE if ctx.get('letterhead') else DARK

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=1.8 * cm, rightMargin=1.8 * cm,
        topMargin=1.6 * cm, bottomMargin=1.6 * cm,
    )

    def rl_markup(text):
        text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        text = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', text)
        return text

    normal = ParagraphStyle('normal', fontName='Helvetica', fontSize=10, leading=15.5, textColor=DARK)
    justified = ParagraphStyle('justified', parent=normal, alignment=TA_JUSTIFY, spaceAfter=10)
    bold_style = ParagraphStyle('bold', parent=normal, fontName='Helvetica-Bold')
    grey_small = ParagraphStyle('greysmall', parent=normal, fontSize=9, textColor=GREY, fontName='Helvetica-Bold')
    right_grey = ParagraphStyle('right', parent=grey_small, alignment=TA_RIGHT)
    company_style = ParagraphStyle('company', parent=normal, alignment=TA_CENTER, fontName='Helvetica-Bold',
                                    fontSize=16, textColor=DARK, spaceAfter=4)
    addr_style = ParagraphStyle('addr', parent=normal, alignment=TA_CENTER, fontSize=8.5, textColor=GREY)
    title_style = ParagraphStyle('title', parent=normal, alignment=TA_CENTER, fontName='Helvetica-Bold',
                                  fontSize=14, textColor=TITLE_COLOR)
    label_style = ParagraphStyle('label', parent=normal, fontName='Helvetica-Bold', fontSize=8.5, textColor=GREY)
    section_head = ParagraphStyle('section', parent=normal, fontName='Helvetica-Bold', fontSize=9,
                                   textColor=ORANGE, spaceAfter=6)
    footer_style = ParagraphStyle('footer', parent=normal, alignment=TA_CENTER, fontSize=7.5,
                                   textColor=GREY, fontName='Helvetica-Oblique')

    story = []

    # ── LETTERHEAD (HR letters only) ─────────────────────────────────────────
    if ctx.get('letterhead'):
        story.append(Paragraph(cs.company_name, company_style))
        story.append(Paragraph(cs.company_address.replace('\n', ', '), addr_style))
        contact_bits = []
        if cs.company_phone:
            contact_bits.append(f"Tel: {cs.company_phone}")
        if cs.company_email:
            contact_bits.append(f"Email: {cs.company_email}")
        if cs.company_trn:
            contact_bits.append(f"TRN: {cs.company_trn}")
        if contact_bits:
            story.append(Paragraph('   |   '.join(contact_bits), addr_style))
        story.append(Spacer(1, 8))
        story.append(HRFlowable(width='100%', thickness=2, color=ORANGE))
        story.append(Spacer(1, 14))

    # ── REF / DATE BAR ────────────────────────────────────────────────────────
    date_text = f"Date: {fmt_date(ctx['date'])}"
    story.append(Paragraph(date_text, right_grey))
    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#CBD5E1')))
    story.append(Spacer(1, 10))

    # ── TITLE ─────────────────────────────────────────────────────────────────
    story.append(Paragraph(ctx['title'].upper(), title_style))
    story.append(Spacer(1, 2))
    story.append(HRFlowable(width='30%', thickness=1.4, color=TITLE_COLOR, hAlign='CENTER'))
    story.append(Spacer(1, 14))

    # ── TO / FROM BLOCKS ─────────────────────────────────────────────────────
    def block(label, lines, bold_first=1):
        story.append(Paragraph(f"{label},", grey_small))
        for i, line in enumerate(lines):
            st = bold_style if i < bold_first else normal
            story.append(Paragraph(line.replace('&', '&amp;'), ParagraphStyle('bl', parent=st, spaceBefore=2, fontSize=10.5)))
        story.append(Spacer(1, 8))

    if ctx.get('to_block'):
        block('To', ctx['to_block'])
    if ctx.get('from_block'):
        block('From', ctx['from_block'])

    # ── SUBJECT ───────────────────────────────────────────────────────────────
    story.append(Paragraph(f"<b>Subject: {ctx['subject']}</b>", ParagraphStyle('subj', parent=normal, fontSize=10.5)))
    story.append(Spacer(1, 3))
    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#CBD5E1')))
    story.append(Spacer(1, 12))

    # ── SALUTATION ────────────────────────────────────────────────────────────
    if ctx.get('salutation'):
        story.append(Paragraph(ctx['salutation'], normal))
        story.append(Spacer(1, 8))

    # ── INTRO ─────────────────────────────────────────────────────────────────
    for para in ctx.get('intro', []):
        story.append(Paragraph(rl_markup(para), justified))

    # ── DETAILS BOX ───────────────────────────────────────────────────────────
    if ctx.get('details_box'):
        db = ctx['details_box']
        story.append(Paragraph(db['title'].upper(), section_head))
        rows = [[Paragraph(lbl, ParagraphStyle('l', parent=normal, fontSize=9.5, textColor=GREY)),
                 Paragraph(f"<b>{val}</b>", ParagraphStyle('v', parent=normal, fontSize=9.5))]
                for lbl, val in db['rows']]
        box = Table(rows, colWidths=['38%', '62%'])
        box.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.75, BOX_BORDER),
            ('INNERGRID', (0, 0), (-1, -1), 0.5, BOX_BORDER),
            ('BACKGROUND', (0, 0), (-1, -1), BOX_TINT),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10), ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 6), ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(box)
        story.append(Spacer(1, 12))

    # ── REASON / NOTE BLOCK ───────────────────────────────────────────────────
    if ctx.get('reason_block'):
        rb = ctx['reason_block']
        story.append(Paragraph(rb['title'].upper(), section_head))
        story.append(Paragraph(rl_markup(rb['text']), justified))
        story.append(Spacer(1, 4))

    # ── BODY (closing paragraphs) ────────────────────────────────────────────
    for para in ctx['body']:
        story.append(Paragraph(rl_markup(para), justified))

    story.append(Spacer(1, 22))

    # ── SIGNATURE ─────────────────────────────────────────────────────────────
    if ctx['signature_mode'] == 'employee_request':
        story.append(Paragraph(ctx.get('closing_phrase', 'Yours sincerely,'), normal))
        story.append(Spacer(1, 42))
        story.append(HRFlowable(width='40%', thickness=0.75, color=colors.HexColor('#333333'), hAlign='LEFT'))
        story.append(Spacer(1, 3))
        sig = ctx['signature']
        story.append(Paragraph(f"<b>{sig['name']}</b>", normal))
        for line in sig.get('lines', []):
            story.append(Paragraph(line, ParagraphStyle('sigline', parent=normal, fontSize=9.5, textColor=GREY)))
    elif ctx['signature_mode'] == 'warning':
        story.append(Paragraph(f"For <b>{cs.company_name}</b>", normal))
        story.append(Spacer(1, 10))
        t = Table(
            [[ctx.get('signatory_title', 'HR Manager'), 'Employee Acknowledgement (Signature & Date)'],
             ['', '']],
            colWidths=[8.4 * cm, 8.4 * cm], rowHeights=[22, 40],
        )
        t.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.75, BORDER),
            ('INNERGRID', (0, 0), (-1, -1), 0.75, BORDER),
            ('BACKGROUND', (0, 0), (-1, 0), TINT),
            ('TEXTCOLOR', (0, 0), (-1, 0), ORANGE),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9.5),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(t)

    else:
        story.append(Paragraph(f"For <b>{cs.company_name}</b>", normal))
        story.append(Spacer(1, 10))
        t = Table(
            [[ctx.get('signatory_title', 'HR Manager'), 'Approved By (In-Charge)', 'Approved By (Manager)'],
             ['', '', '']],
            colWidths=[5.6 * cm, 5.6 * cm, 5.6 * cm], rowHeights=[22, 40],
        )
        t.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.75, BORDER),
            ('INNERGRID', (0, 0), (-1, -1), 0.75, BORDER),
            ('BACKGROUND', (0, 0), (-1, 0), TINT),
            ('TEXTCOLOR', (0, 0), (-1, 0), ORANGE),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9.5),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(t)

    # ── FOOTER ────────────────────────────────────────────────────────────────
    story.append(Spacer(1, 28))
    story.append(HRFlowable(width='100%', thickness=0.5, color=BORDER))
    story.append(Spacer(1, 4))
    footer_text = ctx.get('footer_note') or (
        f"{cs.company_name}  •  {cs.company_address.splitlines()[0] if cs.company_address else ''}"
    )
    story.append(Paragraph(footer_text, footer_style))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# Letter content builders — one function per letter type.
# Each returns a ctx dict ready for build_letter_docx / build_letter_pdf.
# ─────────────────────────────────────────────────────────────────────────────

def letter_ctx_salary_hike_request(employee, cs, current_salary, requested_salary=None, extra_note='',
                                    ref_no=None, today=None):
    today = today or datetime.date.today()
    duration = service_duration_phrase(employee.joining_date, today)

    from_block = [
        f"{employee.name} (ID: {employee.emirates_id or '—'})",
        employee.job_title or '—',
        f"Date of Joining: {fmt_date(employee.joining_date)}",
    ]
    to_block = ["The Manager", cs.company_name]

    intro = [
        f"I am writing to formally request a review of my current salary in light of my "
        f"**{duration}** of service with **{cs.company_name}**."
    ]

    rows = [
        ('Current Monthly Salary', fmt_money(current_salary)),
    ]
    if requested_salary:
        rows.append(('Requested Monthly Salary', fmt_money(requested_salary)))
    rows.append(('Length of Service', duration.capitalize()))
    if employee.mobile:
        rows.append(('Contact Phone', employee.mobile))

    body = [
        "During my time here, I have performed my duties sincerely and fulfilled my responsibilities to the "
        "best of my abilities. Considering my experience and contribution, I kindly request you to consider "
        "granting a suitable salary increment.",
        "I assure you of my continued commitment and dedication to the company. Thank you for your time and "
        "consideration.",
    ]

    reason_block = None
    if extra_note:
        reason_block = {'title': 'Additional Note', 'text': extra_note}

    return {
        'letterhead': False,
        'title': 'Salary Increment Request',
        'ref_no': ref_no,
        'date': today,
        'from_block': from_block,
        'to_block': to_block,
        'subject': 'Request for Salary Increment',
        'salutation': 'Dear Sir/Madam,',
        'intro': intro,
        'details_box': {'title': 'Salary Details', 'rows': rows},
        'reason_block': reason_block,
        'body': body,
        'signature_mode': 'employee_request',
        'closing_phrase': 'Yours sincerely,',
        'signature': {
            'name': employee.name,
            'lines': [employee.job_title or '', f"ID: {employee.emirates_id or '—'}"],
        },
        'footer_note': 'Employee salary request — For official use only',
        'company': cs,
    }


def letter_ctx_leave_request(employee, cs, leave_type_name, start_date, end_date, total_days, reason='',
                              ref_no=None, today=None):
    today = today or datetime.date.today()

    from_block = [
        f"{employee.name} (ID: {employee.emirates_id or '—'})",
        employee.job_title or '—',
        f"Date of Joining: {fmt_date(employee.joining_date)}",
    ]
    to_block = ["The Manager", cs.company_name]

    intro = [
        f"I am writing to formally request **{leave_type_name}** for a period of **{total_days} "
        f"day{'s' if total_days != 1 else ''}**, from **{fmt_date(start_date)}** to **{fmt_date(end_date)}**."
    ]

    rows = [
        ('Leave Type', leave_type_name),
        ('From Date', fmt_date(start_date)),
        ('To Date', fmt_date(end_date)),
        ('Total Days', f"{total_days} Day{'s' if total_days != 1 else ''}"),
    ]
    if employee.mobile:
        rows.append(('Contact Phone', employee.mobile))

    body = [
        "I kindly request your approval for this leave. I will ensure all my pending work is completed and "
        "properly handed over before my departure, and I will be reachable during my leave period if needed.",
        "Thank you for your consideration.",
    ]

    return {
        'letterhead': False,
        'title': f'{leave_type_name} Request',
        'ref_no': ref_no,
        'date': today,
        'from_block': from_block,
        'to_block': to_block,
        'subject': f'Request for {leave_type_name} from {fmt_date(start_date)} to {fmt_date(end_date)}',
        'salutation': 'Dear Sir/Madam,',
        'intro': intro,
        'details_box': {'title': 'Leave Details', 'rows': rows},
        'reason_block': {'title': 'Reason for Leave', 'text': reason or leave_type_name},
        'body': body,
        'signature_mode': 'employee_request',
        'closing_phrase': 'Yours sincerely,',
        'signature': {
            'name': employee.name,
            'lines': [employee.job_title or '', f"ID: {employee.emirates_id or '—'}"],
        },
        'footer_note': 'Employee leave request — For official use only',
        'company': cs,
    }


def letter_ctx_hike_approval(employee, cs, last_salary, new_salary, hike_percent, effective_date, ref_no, today=None):
    today = today or datetime.date.today()

    to_block = [
        f"{employee.name} (ID: {employee.emirates_id or '—'})",
        employee.job_title or '—',
    ]

    intro = [
        f"We are pleased to inform you that, based on your performance and contribution to "
        f"**{cs.company_name}**, the management has approved a revision to your salary as detailed below."
    ]

    rows = [
        ('Previous Monthly Salary', fmt_money(last_salary)),
        ('New Monthly Salary', fmt_money(new_salary)),
        ('Increment', f"{hike_percent:.1f}%"),
        ('Effective From', fmt_date(effective_date)),
    ]

    body = [
        "This revision reflects our appreciation of your continued dedication and hard work. We look forward "
        "to your ongoing contribution to the growth of the company.",
        "Please treat this letter as confirmation of the above salary revision for your records.",
    ]

    return {
        'letterhead': True,
        'title': 'Salary Increment Approval',
        'ref_no': ref_no,
        'date': today,
        'from_block': None,
        'to_block': to_block,
        'subject': 'Salary Increment Approval',
        'salutation': f"Dear {employee.name},",
        'intro': intro,
        'details_box': {'title': 'Salary Revision Details', 'rows': rows},
        'reason_block': None,
        'body': body,
        'signature_mode': 'hr',
        'signatory_title': 'HR Manager',
        'company': cs,
    }


def letter_ctx_experience(employee, cs, last_working_date, still_employed, job_title_override, ref_no, today=None):
    today = today or datetime.date.today()
    position = job_title_override or employee.job_title or '—'

    if still_employed:
        period = f"{fmt_date(employee.joining_date)} to date"
        intro = [
            f"This is to certify that **{employee.name}** has been working with **{cs.company_name}** "
            f"since **{fmt_date(employee.joining_date)}** to date, in the position of **{position}**."
        ]
    else:
        period = f"{fmt_date(employee.joining_date)} to {fmt_date(last_working_date)}"
        intro = [
            f"This is to certify that **{employee.name}** worked with **{cs.company_name}** from "
            f"**{fmt_date(employee.joining_date)}** to **{fmt_date(last_working_date)}**, in the position "
            f"of **{position}**."
        ]

    rows = [
        ('Employee Name', employee.name),
        ('Emirates ID', employee.emirates_id or '—'),
        ('Position', position),
        ('Period of Service', period),
    ]

    body = [
        "During the tenure with our organization, we found the employee to be sincere, hardworking, and "
        "professional in conduct. We wish them continued success in their future endeavours.",
        "This certificate is issued upon the employee's request for whatever purpose it may serve best.",
    ]

    return {
        'letterhead': True,
        'title': 'Experience Certificate',
        'ref_no': ref_no,
        'date': today,
        'from_block': None,
        'to_block': ["To Whomsoever It May Concern,"],
        'subject': 'Experience Certificate',
        'salutation': '',
        'intro': intro,
        'details_box': {'title': 'Employment Details', 'rows': rows},
        'reason_block': None,
        'body': body,
        'signature_mode': 'hr',
        'signatory_title': 'HR Manager',
        'company': cs,
    }


def letter_ctx_noc(employee, cs, purpose, addressed_to, ref_no, today=None):
    today = today or datetime.date.today()

    intro = [
        f"This is to certify that **{employee.name}** is/was employed with **{cs.company_name}** since "
        f"**{fmt_date(employee.joining_date)}**, working as **{employee.job_title or '—'}**.",

        f"We, **{cs.company_name}**, have no objection to the above-named employee proceeding with "
        f"**{purpose}**.",
    ]

    rows = [
        ('Employee Name', employee.name),
        ('Emirates ID', employee.emirates_id or '—'),
        ('Position', employee.job_title or '—'),
        ('Purpose', purpose),
    ]

    body = [
        "This No Objection Certificate is issued upon the employee's request for whatever legal purpose it "
        "may serve.",
    ]

    return {
        'letterhead': True,
        'title': 'No Objection Certificate',
        'ref_no': ref_no,
        'date': today,
        'from_block': None,
        'to_block': [addressed_to or "To Whomsoever It May Concern,"],
        'subject': 'No Objection Certificate (NOC)',
        'salutation': '',
        'intro': intro,
        'details_box': {'title': 'Certificate Details', 'rows': rows},
        'reason_block': None,
        'body': body,
        'signature_mode': 'hr',
        'signatory_title': 'HR Manager',
        'company': cs,
    }


def letter_ctx_warning(employee, cs, violation_type, incident_date, warning_level, description,
                        corrective_action, ref_no, today=None):
    today = today or datetime.date.today()

    to_block = [
        f"{employee.name} (ID: {employee.emirates_id or '—'})",
        employee.job_title or '—',
    ]

    intro = [
        f"This letter serves as formal notice regarding a **{violation_type}** incident on "
        f"**{fmt_date(incident_date)}**, as described below."
    ]

    rows = [
        ('Warning Level', warning_level),
        ('Violation Type', violation_type),
        ('Incident Date', fmt_date(incident_date)),
    ]

    reason_block = {'title': 'Description of Incident', 'text': description} if description else None

    body = [
        "Please treat this matter with the seriousness it deserves. Repeated or continued violations of "
        "company policy may result in further disciplinary action, up to and including termination of "
        "employment, in accordance with company policy and UAE labour law.",
    ]
    if corrective_action:
        body.insert(0, f"**Required corrective action:** {corrective_action}")

    body.append(
        "Please sign below to acknowledge that you have received and understood this letter. Your signature "
        "does not necessarily indicate agreement, only that the letter has been received and discussed with you."
    )

    return {
        'letterhead': True,
        'title': f'{warning_level} — Disciplinary Notice',
        'ref_no': ref_no,
        'date': today,
        'from_block': None,
        'to_block': to_block,
        'subject': f'{warning_level}: {violation_type}',
        'salutation': f"Dear {employee.name},",
        'intro': intro,
        'details_box': {'title': 'Warning Details', 'rows': rows},
        'reason_block': reason_block,
        'body': body,
        'signature_mode': 'warning',
        'signatory_title': 'HR Manager',
        'company': cs,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Dispatcher — rebuilds a ctx from a saved GeneratedLetter record (for
# edit-prefill, list/detail downloads, and regeneration after edits).
# ─────────────────────────────────────────────────────────────────────────────

def build_ctx_for_record(letter_type, employee, cs, details):
    d = details or {}
    today = parse_date(d.get('letter_date')) or datetime.date.today()
    ref_no = d.get('ref_no')

    if letter_type == 'salary_hike_request':
        return letter_ctx_salary_hike_request(
            employee, cs,
            current_salary=d.get('current_salary'),
            requested_salary=d.get('requested_salary'),
            extra_note=d.get('extra_note', ''),
            ref_no=ref_no,
            today=today,
        )
    if letter_type == 'leave_request':
        return letter_ctx_leave_request(
            employee, cs,
            leave_type_name=d.get('leave_type', 'Leave'),
            start_date=parse_date(d.get('start_date')),
            end_date=parse_date(d.get('end_date')),
            total_days=d.get('total_days', 0),
            reason=d.get('reason', ''),
            ref_no=ref_no,
            today=today,
        )
    if letter_type == 'hike_approval':
        return letter_ctx_hike_approval(
            employee, cs,
            last_salary=d.get('last_salary'),
            new_salary=d.get('new_salary'),
            hike_percent=d.get('hike_percent', 0),
            effective_date=parse_date(d.get('effective_date')),
            ref_no=ref_no,
            today=today,
        )
    if letter_type == 'experience':
        return letter_ctx_experience(
            employee, cs,
            last_working_date=parse_date(d.get('last_working_date')),
            still_employed=d.get('still_employed', True),
            job_title_override=d.get('job_title_override', ''),
            ref_no=ref_no,
            today=today,
        )
    if letter_type == 'noc':
        return letter_ctx_noc(
            employee, cs,
            purpose=d.get('purpose', ''),
            addressed_to=d.get('addressed_to', ''),
            ref_no=ref_no,
            today=today,
        )
    if letter_type == 'warning':
        return letter_ctx_warning(
            employee, cs,
            violation_type=d.get('violation_type', ''),
            incident_date=parse_date(d.get('incident_date')),
            warning_level=d.get('warning_level', 'First Warning'),
            description=d.get('description', ''),
            corrective_action=d.get('corrective_action', ''),
            ref_no=ref_no,
            today=today,
        )
    raise ValueError(f"Unknown letter_type: {letter_type}")
