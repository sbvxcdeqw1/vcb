"""Шахматка по образцу 1С: услуги по строкам, месяцы по столбцам (Начислено / Оплата), остатки.

Лист «Шахматка» - весь период; затем лист на каждый год (остаток на начало года - в «Начальном остатке»).
Оплаты показываются отрицательными, столбец «Оплата» выводится только в месяцах, где были оплаты.
"""
from __future__ import annotations

from collections import OrderedDict, defaultdict
from functools import partial
from datetime import date
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.utils import get_column_letter as L

from .. import __version__
from ..models import CHARGE, OPENING, PAYMENT
from .context import Case
from .xlsx_util import put as _put
from .xlsx_util import (BORDER, CENTER, HEADER_FILL, LEFT, MONEY, TOTAL_FILL, header_row, put, setup_print,
                        signature_block, title_block, widths)

MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь",
          "Декабрь"]
ZERO = Decimal(0)


def write_checkerboard(path, case: Case):
    grid = [e for res in case.results.values() for e in res.grid_entries]
    entries = [e for e in (grid or case.account.entries) if e.kind in (CHARGE, PAYMENT, OPENING) and e.date <= case.as_of]
    if not entries:
        return
    services = list(OrderedDict.fromkeys(e.purpose for e in entries))
    first = min(e.date for e in entries if e.kind != OPENING) if any(e.kind != OPENING for e in entries) else None
    last = max(e.date for e in entries)
    wb = Workbook()
    ws = wb.active
    ws.title = "Шахматка"
    _grid(ws, case, entries, services, (first.year, first.month), (last.year, last.month), kind="full")
    for y in range(first.year, last.year + 1):
        sheet = wb.create_sheet(str(y))
        _grid(sheet, case, entries, services, (y, 1 if y > first.year else first.month),
              (y, 12 if y < last.year else last.month))
    wb.save(path)


def _grid(ws, case, entries, services, start: tuple[int, int], end: tuple[int, int], kind: str = "year"):
    put = partial(_put, size=9)  # таблица шахматки набрана 9 пт: на вертикальной странице помещается больше месяцев
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

    ncols_est = 3 + sum(2 if ym in has_pay else 1 for ym in months) + 3
    org = case.claimant.short_name or case.claimant.name
    head_lines = ([(org, "org")] if org else []) + [
        ("РАСЧЁТ НАЧИСЛЕНИЙ И ОПЛАТ ПО УСЛУГАМ", "title"),
        (f"за период {MONTHS[months[0][1] - 1].lower()} {months[0][0]} г. – {MONTHS[months[-1][1] - 1].lower()} "
         f"{months[-1][0]} г.", "center"), ("", "gap"),
        (f"Лицевой счёт: {case.account.number}", "left"),
        (f"Адрес: {case.debtor.address or case.account.address}", "left")]
    if case.debtor.fio or case.account.owner:
        head_lines.append((f"Собственник: {case.debtor.fio or case.account.owner}", "left"))
    head_lines.append(("", "gap"))
    # шапка занимает столбцы A:C - они повторяются на каждой печатной странице, текст не обрезается
    top = title_block(ws, 3, head_lines, width_chars=40)
    sub = top + 1
    merges = []
    put(ws, f"A{top}", "Услуга", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
    put(ws, f"B{top}", None, border=True, fill=HEADER_FILL)
    merges.append(dict(start_row=top, start_column=1, end_row=sub, end_column=2))
    put(ws, f"C{top}", "Начальный остаток", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
    merges.append(dict(start_row=top, start_column=3, end_row=sub, end_column=3))
    col = 4
    colmap = {}  # (y, m, "c"/"p") -> номер столбца
    for ym in months:
        title = f"{MONTHS[ym[1] - 1]}\n{ym[0]}"
        span = 2 if ym in has_pay else 1
        put(ws, f"{L(col)}{top}", title, bold=True, border=True, align=CENTER, fill=HEADER_FILL)
        if span == 2:
            merges.append(dict(start_row=top, start_column=col, end_row=top, end_column=col + 1))
            put(ws, f"{L(col + 1)}{top}", None, border=True, fill=HEADER_FILL)
        put(ws, f"{L(col)}{sub}", "Начислено", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
        colmap[(*ym, "c")] = col
        if span == 2:
            put(ws, f"{L(col + 1)}{sub}", "Оплата", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
            colmap[(*ym, "p")] = col + 1
        col += span
    tot_c, tot_p, close = col, col + 1, col + 3
    put(ws, f"{L(tot_c)}{top}", "Итог", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
    merges.append(dict(start_row=top, start_column=tot_c, end_row=top, end_column=tot_p))
    put(ws, f"{L(tot_p)}{top}", None, border=True, fill=HEADER_FILL)
    put(ws, f"{L(tot_c)}{sub}", "Начислено", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
    put(ws, f"{L(tot_p)}{sub}", "Оплата", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
    put(ws, f"{L(close)}{top}", "Конечный остаток", bold=True, border=True, align=CENTER, fill=HEADER_FILL)
    merges.append(dict(start_row=top, start_column=close, end_row=sub, end_column=close))
    ws.row_dimensions[top].height = 30
    ws.row_dimensions[sub].height = 18
    for c_ in range(tot_p + 1, close):  # пустой промежуток между «Итог» и «Конечный остаток» не рисуем
        pass

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
    put(ws, f"A{rt}", "Итог:", bold=True, border=True, fill=TOTAL_FILL)
    put(ws, f"B{rt}", None, border=True, fill=TOTAL_FILL)
    merges.append(dict(start_row=rt, start_column=1, end_row=rt, end_column=2))
    for c in [3] + [v for v in colmap.values()] + [tot_c, tot_p, close]:
        put(ws, f"{L(c)}{rt}", f"=SUM({L(c)}{r0}:{L(c)}{rt - 1})", bold=True, border=True, fmt=MONEY, fill=TOTAL_FILL)
    signature_block(ws, rt + 3, case, ncols=3)
    widths(ws, {"A": 26, "B": 2, "C": 11, **{L(c): 9.2 for c in range(4, close + 1)}})
    widths(ws, {L(tot_c): 11, L(tot_p): 11, L(tot_p + 1): 1.5, L(close): 11.5})
    for mg in merges:
        ws.merge_cells(**mg)
    ws.freeze_panes = ws.cell(sub + 1, 4)
    # Печать: A4 вертикально, масштаб 85%. Слева на каждой странице повторяются «Услуга» и «Остаток на начало»,
    # месяцы (с их столбцом «Оплата») укладываются на страницу целиком, итоги - на последней странице.
    scale = 88
    page_cm = 21.0 - 1.5 - 1.0  # ширина A4 минус поля слева (1,5) и справа (1,0), см
    avail = page_cm / (scale / 100) - sum(ws.column_dimensions[L(c)].width for c in (1, 2, 3)) * 0.2
    groups = [[colmap[(y, m, k)] for k in ("c", "p") if (y, m, k) in colmap] for y, m in months]
    groups.append([tot_c, tot_p, tot_p + 1, close])
    breaks, used = [], 0.0
    for g in groups:
        w = sum(ws.column_dimensions[L(c)].width for c in g) * 0.2
        if used and used + w > avail:
            breaks.append(g[0] - 1)
            used = 0.0
        used += w
    setup_print(ws, landscape=False, title_rows=f"{top}:{sub}", title_cols="A:C", fit_width=False, scale=scale,
                col_breaks=tuple(breaks), footer_left=f"vcb {__version__}")
