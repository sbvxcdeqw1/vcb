"""Шахматка начислений и оплат по услугам: всё на одном листе, печать на A4 вертикально.

Месяцы идут слева направо по MONTHS_PER_BLOCK штук в блоке; когда ширина страницы заканчивается, следующие
месяцы переносятся «на новую строку» - в следующий блок под предыдущим. Каждый блок имеет собственную шапку
(месяцы, «Начислено» / «Оплата») и строку «Итог». В конце - сводный блок «Итого за период» по услугам.
Оплаты показываются отрицательными и стоят в тех периодах, которые они закрывают (разноска по реестру оплат).
"""
from __future__ import annotations

from collections import OrderedDict, defaultdict
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.pagebreak import Break

from .. import __version__
from ..models import CHARGE, PAYMENT
from .context import Case
from .xlsx_util import CENTER, HEADER_FILL, LEFT, MONEY, TOTAL_FILL, setup_print, signature_block, title_block, widths
from .xlsx_util import put as _put

MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь",
          "Декабрь"]
MONTHS_PER_BLOCK = 4  # 4 месяца x 2 столбца + «Услуга» = ширина A4 в вертикальной ориентации
NCOLS = 2 + MONTHS_PER_BLOCK * 2  # A:B - услуга, дальше пары «Начислено / Оплата»
PAGE_HEIGHT = 760  # пунктов листа на страницу при печати (с запасом на масштаб по ширине)


def write_checkerboard(path, case: Case):
    grid = [e for res in case.results.values() for e in res.grid_entries]
    entries = [e for e in (grid or case.account.entries) if e.kind in (CHARGE, PAYMENT) and e.date <= case.as_of]
    if not entries:
        return
    services = list(OrderedDict.fromkeys(e.purpose for e in entries))
    charges, paid = defaultdict(Decimal), defaultdict(Decimal)
    for e in entries:
        (paid if e.kind == PAYMENT else charges)[(e.purpose, e.date.year, e.date.month)] += e.amount
    first, last = min(e.date for e in entries), max(e.date for e in entries)
    months, y, m = [], first.year, first.month
    while (y, m) <= (last.year, last.month):
        months.append((y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)

    wb = Workbook()
    ws = wb.active
    ws.title = "Шахматка"
    put = lambda *a, **k: _put(*a, **{"size": 9, **k})  # таблица набрана 9 пт

    org = case.claimant.short_name or case.claimant.name
    lines = ([(org, "org")] if org else []) + [
        ("РАСЧЁТ НАЧИСЛЕНИЙ И ОПЛАТ ПО УСЛУГАМ", "title"),
        (f"за период {MONTHS[months[0][1] - 1].lower()} {months[0][0]} г. – {MONTHS[months[-1][1] - 1].lower()} "
         f"{months[-1][0]} г.", "center"), ("", "gap"),
        (f"Лицевой счёт: {case.account.number}", "left"),
        (f"Адрес: {case.debtor.address or case.account.address}", "left")]
    if case.debtor.fio or case.account.owner:
        lines.append((f"Собственник: {case.debtor.fio or case.account.owner}", "left"))
    lines.append(("", "gap"))
    r = title_block(ws, NCOLS, lines, width_chars=100)
    used = sum(ws.row_dimensions[i].height or 15 for i in range(1, r))  # высота напечатанного на странице, пт
    breaks: list[int] = []

    def service_height(name: str) -> float:
        return 26 if len(name) > 30 else 15

    def need(height: float):
        """Начинает новую страницу, если блок высотой `height` не помещается."""
        nonlocal used
        if used + height > PAGE_HEIGHT and used > 0:
            breaks.append(r - 1)
            used = 0
        used += height

    blocks_total: dict[str, list[tuple[str, str]]] = {s: [] for s in services}  # услуга -> ячейки (начислено, оплата)
    for start in range(0, len(months), MONTHS_PER_BLOCK):
        chunk = months[start:start + MONTHS_PER_BLOCK]
        need(30 + 16 + sum(service_height(s) for s in services) + 16 + 10)
        top, sub = r, r + 1
        put(ws, f"A{top}", "Услуга", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
        put(ws, f"B{top}", None, border=True, fill=HEADER_FILL)
        put(ws, f"A{sub}", None, border=True, fill=HEADER_FILL)
        put(ws, f"B{sub}", None, border=True, fill=HEADER_FILL)
        ws.merge_cells(start_row=top, start_column=1, end_row=sub, end_column=2)
        for i in range(MONTHS_PER_BLOCK):
            c = 3 + i * 2
            if i < len(chunk):
                ym = chunk[i]
                put(ws, f"{L(c)}{top}", f"{MONTHS[ym[1] - 1]}\n{ym[0]}", bold=True, border=True, align=CENTER,
                    fill=HEADER_FILL)
                put(ws, f"{L(c + 1)}{top}", None, border=True, fill=HEADER_FILL)
                ws.merge_cells(start_row=top, start_column=c, end_row=top, end_column=c + 1)
                put(ws, f"{L(c)}{sub}", "Начислено", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
                put(ws, f"{L(c + 1)}{sub}", "Оплата", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
        ws.row_dimensions[top].height, ws.row_dimensions[sub].height = 30, 16
        r = sub + 1
        r0 = r
        for svc in services:
            put(ws, f"A{r}", svc, border=True, align=LEFT)
            put(ws, f"B{r}", None, border=True)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            ws.row_dimensions[r].height = service_height(svc)
            for i, ym in enumerate(chunk):
                c = 3 + i * 2
                ch, pa = charges.get((svc, *ym)), paid.get((svc, *ym))
                put(ws, f"{L(c)}{r}", ch or None, border=True, fmt=MONEY)
                put(ws, f"{L(c + 1)}{r}", -pa if pa else None, border=True, fmt=MONEY)
                blocks_total[svc].append((f"{L(c)}{r}", f"{L(c + 1)}{r}"))
            r += 1
        put(ws, f"A{r}", "Итог:", bold=True, border=True, fill=TOTAL_FILL)
        put(ws, f"B{r}", None, border=True, fill=TOTAL_FILL)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        for i in range(len(chunk)):
            for k in (0, 1):
                col = L(3 + i * 2 + k)
                put(ws, f"{col}{r}", f"=SUM({col}{r0}:{col}{r - 1})", bold=True, border=True, fmt=MONEY, fill=TOTAL_FILL)
        ws.row_dimensions[r].height = 16
        r += 1
        ws.row_dimensions[r].height = 10  # промежуток между блоками
        r += 1

    # сводный блок: итого по каждой услуге за весь период
    need(30 + sum(service_height(s) for s in services) + 16 + 10 + 80)
    top = r
    put(ws, f"A{top}", "Итого за период", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
    put(ws, f"B{top}", None, border=True, fill=HEADER_FILL)
    ws.merge_cells(start_row=top, start_column=1, end_row=top, end_column=2)
    spans = [(3, 5), (6, 8), (9, 10)]  # три группы на всю ширину блоков (C:E, F:H, I:J)
    for (c1, c2), title in zip(spans, ("Начислено, руб.", "Оплата, руб.", "Задолженность, руб.")):
        for c in range(c1, c2 + 1):
            put(ws, f"{L(c)}{top}", title if c == c1 else None, bold=True, border=True, align=CENTER, fill=HEADER_FILL)
        ws.merge_cells(start_row=top, start_column=c1, end_row=top, end_column=c2)
    ws.row_dimensions[top].height = 30
    r = top + 1
    r0 = r
    for svc in services:
        put(ws, f"A{r}", svc, border=True, align=LEFT)
        put(ws, f"B{r}", None, border=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.row_dimensions[r].height = service_height(svc)
        cells_c = "+".join(c for c, _ in blocks_total[svc])
        cells_p = "+".join(p for _, p in blocks_total[svc])
        for (c1, c2), formula in zip(spans, (f"={cells_c}", f"={cells_p}", f"=C{r}+F{r}")):
            for c in range(c1, c2 + 1):
                put(ws, f"{L(c)}{r}", formula if c == c1 else None, border=True, fmt=MONEY)
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        r += 1
    put(ws, f"A{r}", "Итог:", bold=True, border=True, fill=TOTAL_FILL)
    put(ws, f"B{r}", None, border=True, fill=TOTAL_FILL)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    for c1, c2 in spans:
        for c in range(c1, c2 + 1):
            put(ws, f"{L(c)}{r}", f"=SUM({L(c1)}{r0}:{L(c1)}{r - 1})" if c == c1 else None, bold=True, border=True,
                fmt=MONEY, fill=TOTAL_FILL)
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    ws.row_dimensions[r].height = 16
    signature_block(ws, r + 3, case, ncols=NCOLS)

    widths(ws, {"A": 26, "B": 2, **{L(c): 9.4 for c in range(3, NCOLS + 1)}})
    setup_print(ws, landscape=False, footer_left=f"vcb {__version__}")
    for b in breaks:
        ws.row_breaks.append(Break(id=b))
    wb.save(path)
