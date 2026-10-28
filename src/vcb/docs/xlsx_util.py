from openpyxl.styles import Alignment, Border, Font, Side

THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD = Font(name="Times New Roman", bold=True, size=11)
NORMAL = Font(name="Times New Roman", size=11)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
MONEY = "#,##0.00"


def put(ws, ref, value, *, bold=False, border=False, align=None, fmt=None):
    c = ws[ref]
    c.value = value
    c.font = BOLD if bold else NORMAL
    if border:
        c.border = BORDER
    if align:
        c.alignment = align
    if fmt:
        c.number_format = fmt
    return c


def widths(ws, spec: dict):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def printable(ws, landscape=True):
    """Печать на страницу по ширине."""
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
