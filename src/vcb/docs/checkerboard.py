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

from .. import __version__
from ..models import CHARGE, OPENING, PAYMENT
from .context import Case
from .xlsx_util import (BORDER, CENTER, HEADER_FILL, LEFT, MONEY, TOTAL_FILL, header_row, put, setup_print,
                        title_block, widths)

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
    _grid(ws, case, entries, services, (first.year, first.month), (last.year, last.month), kind="full")
    for y in range(first.year, last.year + 1):
        sheet = wb.create_sheet(str(y))
        _grid(sheet, case, entries, services, (y, 1 if y > first.year else first.month),
              (y, 12 if y < last.year else last.month))
    wb.save(path)


def _grid(ws, case, entries, services, start: tuple[int, int], end: tuple[int, int], kind: str = "year"):
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
        ("ШАХМАТКА НАЧИСЛЕНИЙ И ОПЛАТ ПО УСЛУГАМ", "title"),
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
    widths(ws, {"A": 28, "B": 2, "C": 12.5, **{L(c): 10.5 for c in range(4, close + 1)}})
    widths(ws, {L(tot_c): 12, L(tot_p): 12, L(tot_p + 1): 1.5, L(close): 13})
    for mg in merges:
        ws.merge_cells(**mg)
    ws.freeze_panes = ws.cell(sub + 1, 4)
    if kind == "full":  # весь период: страница = полугодие, слева повторяются «Услуга» и «Начальный остаток»
        halves = [(y, m) for y, m in months if m in (1, 7)]
        breaks = tuple(colmap[(y, m, "c")] - 1 for y, m in halves[1:]) + (tot_c - 1,)
        per_half: dict = {}
        for (y, m, _k), c_ in colmap.items():
            key = (y, m > 6)
            per_half[key] = per_half.get(key, 0) + 1
        # ширина столбца Excel ≈ 0,2 см; рабочая ширина A4 в альбомной ориентации ≈ 27 см
        widest = (30 + 12.5 + max(per_half.values()) * 10.5) * 0.2
        scale = max(60, min(100, int(27.0 / widest * 100) // 5 * 5))
        setup_print(ws, landscape=True, title_rows=f"{top}:{sub}", title_cols="A:C", fit_width=False, scale=scale,
                    col_breaks=breaks, footer_left=f"vcb {__version__}")
    else:  # один год: весь лист на одной странице по ширине
        setup_print(ws, landscape=True, title_rows=f"{top}:{sub}", footer_left=f"vcb {__version__}")
