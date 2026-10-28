"""Шахматка по образцу 1С: услуги по строкам, месяцы по столбцам (Начислено / Оплата), остатки.

Лист «Шахматка» - весь период; затем лист на каждый год (остаток на начало года - в «Начальном остатке»).
Оплаты показываются отрицательными, столбец «Оплата» выводится только в месяцах, где были оплаты.
"""
from __future__ import annotations

from collections import OrderedDict, defaultdict
from datetime import date
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.utils import get_column_letter as L

from ..models import CHARGE, OPENING, PAYMENT
from .context import Case
from .xlsx_util import *

MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь",
          "Декабрь"]
ZERO = Decimal(0)


def write_checkerboard(path, case: Case):
    entries = [e for e in case.account.entries if e.kind in (CHARGE, PAYMENT, OPENING) and e.date <= case.as_of]
    if not entries:
        return
    services = list(OrderedDict.fromkeys(e.purpose for e in entries))
    first = min(e.date for e in entries if e.kind != OPENING) if any(e.kind != OPENING for e in entries) else None
    last = max(e.date for e in entries)
    wb = Workbook()
    ws = wb.active
    ws.title = "Шахматка"
    _grid(ws, case, entries, services, (first.year, first.month), (last.year, last.month))
    for y in range(first.year, last.year + 1):
        sheet = wb.create_sheet(str(y))
        _grid(sheet, case, entries, services, (y, 1 if y > first.year else first.month),
              (y, 12 if y < last.year else last.month))
    for w in wb.worksheets:
        printable(w, landscape=True)
    wb.save(path)


def _grid(ws, case, entries, services, start: tuple[int, int], end: tuple[int, int]):
    start_d = date(*start, 1)
    months = []
    y, m = start
    while (y, m) <= end:
        months.append((y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    # суммы
    ch, pay = defaultdict(Decimal), defaultdict(Decimal)
    opening = defaultdict(Decimal)
    for e in entries:
        if e.date < start_d:  # всё, что было до листа, уходит в начальный остаток
            opening[e.purpose] += -e.amount if e.kind == PAYMENT else e.amount
        elif e.kind == OPENING:
            opening[e.purpose] += e.amount
        elif e.kind == PAYMENT:
            pay[(e.purpose, e.date.year, e.date.month)] += e.amount
        else:
            ch[(e.purpose, e.date.year, e.date.month)] += e.amount
    has_pay = {(y, m) for (_, y, m) in pay}

    put(ws, "A2", "Шахматка", bold=True)
    put(ws, "A3", f"за период {MONTHS[months[0][1] - 1]} {months[0][0]} г. - {MONTHS[months[-1][1] - 1]} "
                  f"{months[-1][0]} г.")
    put(ws, "A4", f"{case.account.number}")
    put(ws, "A5", case.debtor.address or case.account.address)
    if case.debtor.fio or case.account.owner:
        put(ws, "A6", case.debtor.fio or case.account.owner)
    top, sub = 8, 9
    merges = []
    put(ws, f"A{top}", "Услуга", bold=True, border=True, align=CENTER)
    merges.append(dict(start_row=top, start_column=1, end_row=sub, end_column=2))
    put(ws, f"C{top}", "Начальный остаток", bold=True, border=True, align=CENTER)
    merges.append(dict(start_row=top, start_column=3, end_row=sub, end_column=3))
    col = 4
    colmap = {}  # (y, m, "c"/"p") -> номер столбца
    for ym in months:
        title = f"{MONTHS[ym[1] - 1]} {ym[0]}"
        span = 2 if ym in has_pay else 1
        put(ws, f"{L(col)}{top}", title, bold=True, border=True, align=CENTER)
        if span == 2:
            merges.append(dict(start_row=top, start_column=col, end_row=top, end_column=col + 1))
            put(ws, f"{L(col + 1)}{top}", None, border=True)
        put(ws, f"{L(col)}{sub}", "Начислено", bold=True, border=True, align=CENTER)
        colmap[(*ym, "c")] = col
        if span == 2:
            put(ws, f"{L(col + 1)}{sub}", "Оплата", bold=True, border=True, align=CENTER)
            colmap[(*ym, "p")] = col + 1
        col += span
    tot_c, tot_p, close = col, col + 1, col + 3
    put(ws, f"{L(tot_c)}{top}", "Итог", bold=True, border=True, align=CENTER)
    merges.append(dict(start_row=top, start_column=tot_c, end_row=top, end_column=tot_p))
    put(ws, f"{L(tot_p)}{top}", None, border=True)
    put(ws, f"{L(tot_c)}{sub}", "Начислено", bold=True, border=True, align=CENTER)
    put(ws, f"{L(tot_p)}{sub}", "Оплата", bold=True, border=True, align=CENTER)
    put(ws, f"{L(close)}{top}", "Конечный остаток", bold=True, border=True, align=CENTER)
    merges.append(dict(start_row=top, start_column=close, end_row=sub, end_column=close))

    r0 = sub + 1
    for i, svc in enumerate(services):
        r = r0 + i
        put(ws, f"A{r}", svc, border=True, align=LEFT)
        merges.append(dict(start_row=r, start_column=1, end_row=r, end_column=2))
        put(ws, f"C{r}", opening.get(svc) or None, border=True, fmt=MONEY)
        for ym in months:
            c = colmap[(*ym, "c")]
            put(ws, f"{L(c)}{r}", ch.get((svc, *ym)) or None, border=True, fmt=MONEY)
            if (*ym, "p") in colmap:
                p = pay.get((svc, *ym))
                put(ws, f"{L(colmap[(*ym, 'p')])}{r}", -p if p else None, border=True, fmt=MONEY)
        c_cols = [L(colmap[(*ym, "c")]) for ym in months]
        p_cols = [L(colmap[(*ym, "p")]) for ym in months if (*ym, "p") in colmap]
        put(ws, f"{L(tot_c)}{r}", "=" + "+".join(f"{c}{r}" for c in c_cols), border=True, fmt=MONEY)
        put(ws, f"{L(tot_p)}{r}", ("=" + "+".join(f"{c}{r}" for c in p_cols)) if p_cols else 0, border=True,
            fmt=MONEY)
        put(ws, f"{L(close)}{r}", f"=C{r}+{L(tot_c)}{r}+{L(tot_p)}{r}", border=True, fmt=MONEY)
    rt = r0 + len(services)
    put(ws, f"A{rt}", "Итог:", bold=True, border=True)
    merges.append(dict(start_row=rt, start_column=1, end_row=rt, end_column=2))
    for c in [3] + [v for v in colmap.values()] + [tot_c, tot_p, close]:
        put(ws, f"{L(c)}{rt}", f"=SUM({L(c)}{r0}:{L(c)}{rt - 1})", bold=True, border=True, fmt=MONEY)
    widths(ws, {"A": 30, "B": 2, "C": 14, **{L(c): 12 for c in range(4, close + 1)}})
    for mg in merges:
        ws.merge_cells(**mg)
    ws.freeze_panes = ws.cell(sub + 1, 4)
