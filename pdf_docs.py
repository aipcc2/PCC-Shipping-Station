"""
PDF versions of the shipping documents, drawn directly by the app (no Excel needed).

Each writer takes the .xlsx path the app just wrote and saves the PDF next to it with the same name
(".../4099-11-5-26-COC.xlsx" -> ".../4099-11-5-26-COC.pdf"), using the same data, so the two always
match. Layouts follow the .xlsx forms: same headings, fields, serial tables, Code128 barcodes, totals.

Needs: reportlab (pure Python; build_exe.bat installs it and bundles it with --collect-all reportlab).
"""

from pathlib import Path

from reportlab.graphics.barcode.code128 import Code128
from reportlab.lib.colors import HexColor, black
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen.canvas import Canvas

GREEN = HexColor("#2A9D5C")
RED = HexColor("#C0392B")
GREY = HexColor("#666666")
PAGE_W, PAGE_H = letter

COMPANY = "PRINTED CIRCUITS CORP."
TAGLINE = "Design, Manufacturing and Assembly of Printed Circuit Boards"
ADDRESS = "5295 Webb Parkway, Lilburn, GA 30047"
PHONE = "Phone: (770) 638 8658   Fax: (770) 638 8659"
EMAIL = "Email: contact@pcc-i.com, Website: www.pcc-i.com"


def pdf_path_for(xlsx_path):
    return Path(xlsx_path).with_suffix(".pdf")


def _canvas(path, title):
    c = Canvas(str(path), pagesize=letter)
    c.setTitle(title)
    c.setAuthor(COMPANY)
    c.setCreator("PCC Shipping Station")
    return c


def _s(v):
    return "" if v is None else str(v)


def _barcode(c, value, x, y, max_w, h):
    """Draws a Code128 barcode for value, centred in [x, x+max_w], bottom at y. Skips empty values."""
    value = _s(value).strip()
    if not value:
        return
    try:
        probe = Code128(value, barWidth=1, barHeight=h, quiet=False)
        bw = min(1.2, max_w / max(probe.width, 1))
        bc = Code128(value, barWidth=bw, barHeight=h, quiet=False)
        bc.drawOn(c, x + (max_w - bc.width) / 2, y)
    except Exception:
        pass  # a value Code128 can't encode just gets no barcode, like the .xlsx


def _fit(c, text, font, size, max_w, min_size=6):
    """Largest font size (down to min_size) at which text fits max_w."""
    while size > min_size and stringWidth(text, font, size) > max_w:
        size -= 0.5
    return size


def _text_in_box(c, text, x, y_mid, w, font="Helvetica", size=11, align="left", pad=6):
    text = _s(text)
    avail = w - 2 * pad
    size = _fit(c, text, font, size, avail, min_size=9)
    lines = [text]
    if stringWidth(text, font, size) > avail:  # still too long: two lines, smaller only if needed
        lines = simpleSplit(text, font, size, avail)
        while len(lines) > 2 and size > 6:
            size -= 0.5
            lines = simpleSplit(text, font, size, avail)
        lines = lines[:2]
    c.setFont(font, size)
    top = y_mid + (len(lines) - 1) * size * 0.6 - size * 0.35
    for i, line in enumerate(lines):
        yy = top - i * size * 1.2
        if align == "center":
            c.drawCentredString(x + w / 2, yy, line)
        elif align == "right":
            c.drawRightString(x + w - pad, yy, line)
        else:
            c.drawString(x + pad, yy, line)


class _Table:
    """A bordered table that starts a new page (and repeats its header) when it runs out of room."""

    def __init__(self, c, x, widths, header, header_h, row_h, bottom, on_new_page):
        self.c, self.x, self.widths, self.header = c, x, widths, header
        self.header_h, self.row_h, self.bottom, self.on_new_page = header_h, row_h, bottom, on_new_page

    def draw_header(self, y):
        c = self.c
        c.setFont("Helvetica-Bold", 11)
        x = self.x
        for (label, align), w in zip(self.header, self.widths):
            c.rect(x, y - self.header_h, w, self.header_h)
            _text_in_box(c, label, x, y - self.header_h / 2, w, "Helvetica-Bold", 11, align)
            x += w
        return y - self.header_h

    def rows(self, y, items, draw_cell):
        y = self.draw_header(y)
        for i, item in enumerate(items):
            if y - self.row_h < self.bottom:
                self.c.showPage()
                y = self.draw_header(self.on_new_page())
            x = self.x
            for col, w in enumerate(self.widths):
                self.c.rect(x, y - self.row_h, w, self.row_h)
                draw_cell(col, i, item, x, y - self.row_h, w, self.row_h)
                x += w
            y -= self.row_h
        return y


###############################################################################
# CERTIFICATE OF COMPLIANCE
###############################################################################

def write_coc_pdf(xlsx_path, coc_data, serials):
    out = pdf_path_for(xlsx_path)
    c = _canvas(out, f"Certificate of Compliance {Path(xlsx_path).stem}")
    m = 36
    w = PAGE_W - 2 * m
    y = PAGE_H - m - 20

    c.setFillColor(GREEN); c.setFont("Helvetica-Bold", 20); c.drawCentredString(PAGE_W / 2, y, COMPANY)
    y -= 18; c.setFillColor(RED); c.setFont("Helvetica", 11); c.drawCentredString(PAGE_W / 2, y, TAGLINE)
    c.setFillColor(black)
    for line in (ADDRESS, PHONE, EMAIL):
        y -= 15; c.drawCentredString(PAGE_W / 2, y, line)
    y -= 36; c.setFont("Helvetica-Bold", 14); c.drawCentredString(PAGE_W / 2, y, "CERTIFICATE OF COMPLIANCE")

    pn = _s(coc_data.get("pn")); rv = _s(coc_data.get("rev"))
    rs = f" Rev.- {rv}" if rv else " Rev.-"
    cert = (f"This is to certify that the following part number was manufactured and assembled as per "
            f"drawings and specifications of the {pn}{rs} drawings.")
    y -= 30; c.setFont("Helvetica", 11)
    for line in simpleSplit(cert, "Helvetica", 11, w):
        c.drawString(m, y, line); y -= 14

    y -= 12
    label_right = m + 125
    for label, value in [
        ("Ship Date:", coc_data.get("ship_date")),
        ("PO #:", coc_data.get("po")),
        ("P/N:", pn),
        ("QTY:", f"{_s(coc_data.get('qty'))}ea."),
        ("DC:", f"{_s(coc_data.get('datecode'))} (YYWW)"),
        ("PCC Job#:", coc_data.get("job")),
        ("Type:", coc_data.get("type")),
        ("Assembly process:", coc_data.get("assembly_process")),
    ]:
        c.setFont("Helvetica-Bold", 11); c.drawRightString(label_right, y, label)
        c.setFont("Helvetica", 11); c.drawString(label_right + 10, y, _s(value))
        y -= 17

    y -= 10
    widths = [125, 256, w - 381]

    def cell(col, i, sn, x, yb, cw, ch):
        if col == 0:
            _text_in_box(c, i + 1, x, yb + ch / 2, cw, align="center")
        elif col == 1:
            _text_in_box(c, sn, x, yb + ch / 2, cw, align="center")
        else:
            _barcode(c, sn, x + 8, yb + 6, cw - 16, ch - 12)

    table = _Table(c, m, widths, [("S.No", "center"), ("Serial Number", "center"), ("Barcode", "center")],
                   22, 36, m + 70, lambda: PAGE_H - m)
    y = table.rows(y, list(serials or []), cell)

    if y - 60 < m:
        c.showPage(); y = PAGE_H - m
    c.setFont("Helvetica", 11)
    y -= 10
    for line in ("Hung Ly", "QC", "Printed Circuits Corp."):
        y -= 16; c.drawString(m, y, line)
    c.save()
    return str(out)


###############################################################################
# PACKING LIST
###############################################################################

def write_packing_list_pdf(xlsx_path, pl_data, serials, logo_path=None):
    out = pdf_path_for(xlsx_path)
    c = _canvas(out, f"Packing List {Path(xlsx_path).stem}")
    m = 29
    w = PAGE_W - 2 * m
    col = w / 12
    top = PAGE_H - m

    def header():
        if logo_path and Path(logo_path).exists():
            try:
                c.drawImage(ImageReader(str(logo_path)), m + 4, top - 92, width=80, height=88,
                            preserveAspectRatio=True, anchor="sw", mask="auto")
            except Exception:
                pass
        right = m + w
        c.setFillColor(GREEN); c.setFont("Helvetica-Bold", 22); c.drawRightString(right, top - 24, COMPANY)
        c.setFillColor(RED); c.setFont("Helvetica-Bold", 14); c.drawRightString(right, top - 44, "PACKING LIST")
        c.setFillColor(black); c.setFont("Helvetica", 10)
        c.drawRightString(right, top - 64, ADDRESS)
        c.drawRightString(right, top - 79, PHONE)
        c.setLineWidth(2.2); c.line(m, top - 98, m + w, top - 98); c.setLineWidth(1)
        return top - 108

    y = header()

    # Info table: two rows of four boxes (label, value, barcode of the value).
    box_h = 58
    for row in ([("CUSTOMER", "customer"), ("PACKING LIST #", "packing_no"), ("P.O. #", "po"), ("P/N", "pn")],
                [("WO / LOT #", "wo"), ("PCC ENGINEERING #", "eng"), ("DATE", "date"), ("DATE CODE", "datecode")]):
        for i, (label, key) in enumerate(row):
            x = m + i * 3 * col
            bw = 3 * col
            c.rect(x, y - box_h, bw, box_h)
            c.setFillColor(GREY); c.setFont("Helvetica-Bold", 7); c.drawString(x + 6, y - 10, label); c.setFillColor(black)
            _text_in_box(c, pl_data.get(key), x, y - 26, bw, "Helvetica-Bold", 12)
            _barcode(c, pl_data.get(key), x + 10, y - box_h + 4, bw - 20, 14)
        y -= box_h
    y -= 12

    widths = [col, 4 * col, 4 * col, 3 * col]

    def cell(ci, i, sn, x, yb, cw, ch):
        if ci == 0:
            _text_in_box(c, i + 1, x, yb + ch / 2, cw, align="center")
        elif ci == 1:
            _text_in_box(c, sn, x, yb + ch / 2, cw, "Helvetica-Bold")
        elif ci == 2:
            _barcode(c, sn, x + 10, yb + 5, cw - 20, ch - 10)
        else:
            _text_in_box(c, 1, x, yb + ch / 2, cw, "Helvetica-Bold", align="center")

    table = _Table(c, m, widths, [("#", "center"), ("SERIAL", "left"), ("SERIAL BARCODE", "center"), ("QTY", "center")],
                   22, 34, m + 60, header)
    y = table.rows(y, list(serials or []), cell)

    y -= 8
    if y - 50 < m:
        c.showPage(); y = header()
    c.rect(m, y - 46, 6 * col, 46)
    c.rect(m + 6 * col, y - 46, 6 * col, 46)
    c.setFont("Helvetica-Bold", 12); c.drawRightString(m + w - 8, y - 16, "TOTAL UNITS")
    c.setFont("Helvetica-Bold", 18); c.drawRightString(m + w - 8, y - 39, str(len(serials or [])))
    c.save()
    return str(out)


###############################################################################
# SERIAL LIST
###############################################################################

def write_serials_pdf(xlsx_path, eng, wo, serials, date="", time="", operator=""):
    out = pdf_path_for(xlsx_path)
    c = _canvas(out, f"Serial Numbers {Path(xlsx_path).stem}")
    m = 36
    w = PAGE_W - 2 * m

    def header():
        y = PAGE_H - m - 14
        c.setFillColor(GREEN); c.setFont("Helvetica-Bold", 16)
        c.drawString(m, y, f"Serial Numbers - WO {_s(wo)} (Eng #{_s(eng)})")
        c.setFillColor(black); c.setFont("Helvetica", 10)
        details = [f"Date: {date}", f"Time: {time}", f"Operator: {operator}", f"Total serials: {len(serials or [])}"]
        c.drawString(m, y - 18, "     ".join(d for d in details if not d.endswith(": ")))
        return y - 32

    y = header()
    widths = [60, 220, w - 280]

    def cell(ci, i, sn, x, yb, cw, ch):
        if ci == 0:
            _text_in_box(c, i + 1, x, yb + ch / 2, cw, align="center")
        elif ci == 1:
            _text_in_box(c, sn, x, yb + ch / 2, cw)
        else:
            _barcode(c, sn, x + 10, yb + 5, cw - 20, ch - 10)

    _Table(c, m, widths, [("#", "center"), ("Serial Number", "left"), ("Barcode", "center")],
           22, 32, m, header).rows(y, list(serials or []), cell)
    c.save()
    return str(out)
