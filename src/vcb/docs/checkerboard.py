"""Шахматка: начисления и оплаты по назначению платежа, месяцы по горизонтали, лист на каждый год."""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from openpyxl import Workbook

from ..models import CATEGORY_TITLES, CHARGE, PAYMENT, PENALTY
from .context import Case
from .xlsx_util import *

MONTHS = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
KINDS = [(CHARGE, "Начислено"), (PAYMENT, "Оплачено"), (PENALTY, "Пени (по данным 1С)")]


def write_checkerboard(path, case: Case):
    entries = [e for e in case.account.entries if e.date <= case.as_of]
    if not entries:
        return
    cells = defaultdict(Decimal)  # (год, категория, назначение, вид, месяц) -> сумма
    for e in entries:
        cells[(e.date.year, e.category, e.purpose, e.kind, e.date.month)] += e.amount
    years = sorted({k[0] for k in cells})
    wb = Workbook()
    summary = wb.active
    summary.title = "Итого по годам"
    for y in years:
        _year_sheet(wb.create_sheet(str(y)), case, cells, y)
    _summary(summary, case, cells, years)
    for _ws in wb.worksheets:
        printable(_ws, landscape=True)
    wb.save(path)


def _year_sheet(ws, case, cells, year):
    put(ws, "A1", f"Шахматка начислений и оплат за {year} г. — {case.debtor.fio or case.account.owner}", bold=True)
    put(ws, "A2", f"Лицевой счёт {case.account.number}, {case.debtor.address or case.account.address}")
    hdr = ["Назначение платежа", "Вид"] + MONTHS + ["Итого за год"]
    for i, t in enumerate(hdr, 1):
        put(ws, f"{ws.cell(4, i).coordinate}", t, bold=True, border=True, align=CENTER)
    r = 5
    rows_by_kind = defaultdict(list)
    keys = sorted({(k[1], k[2], k[3]) for k in cells if k[0] == year})
    for cat, purpose, kind in keys:
        label = dict(KINDS)[kind]
        put(ws, f"A{r}", f"{purpose} [{CATEGORY_TITLES[cat].split(',')[0]}]", border=True, align=LEFT)
        put(ws, f"B{r}", label, border=True)
        for m in range(1, 13):
            v = cells.get((year, cat, purpose, kind, m))
            put(ws, ws.cell(r, 2 + m).coordinate, v, border=True, fmt=MONEY)
        put(ws, ws.cell(r, 15).coordinate, f"=SUM(C{r}:N{r})", bold=True, border=True, fmt=MONEY)
        rows_by_kind[kind].append(r)
        r += 1
    r += 1
    for kind, label in KINDS:
        if not rows_by_kind[kind]:
            continue
        put(ws, f"A{r}", f"ИТОГО {label.lower()}", bold=True, border=True)
        put(ws, f"B{r}", None, border=True)
        for col in range(3, 16):
            L = ws.cell(4, col).column_letter
            f = "+".join(f"{L}{x}" for x in rows_by_kind[kind])
            put(ws, ws.cell(r, col).coordinate, f"={f}", bold=True, border=True, fmt=MONEY)
        rows_by_kind[kind + "_total"] = r
        r += 1
    if rows_by_kind.get(CHARGE + "_total") and rows_by_kind.get(PAYMENT + "_total"):
        a, b = rows_by_kind[CHARGE + "_total"], rows_by_kind[PAYMENT + "_total"]
        put(ws, f"A{r}", "Разница (начислено − оплачено)", bold=True, border=True)
        put(ws, f"B{r}", None, border=True)
        for col in range(3, 16):
            L = ws.cell(4, col).column_letter
            put(ws, ws.cell(r, col).coordinate, f"={L}{a}-{L}{b}", bold=True, border=True, fmt=MONEY)
    widths(ws, {"A": 46, "B": 20, **{ws.cell(1, c).column_letter: 11 for c in range(3, 15)}, "O": 14})
    ws.freeze_panes = "C5"


def _summary(ws, case, cells, years):
    put(ws, "A1", f"Итого по годам — {case.debtor.fio or case.account.owner}, ЛС {case.account.number}", bold=True)
    for i, t in enumerate(["Назначение платежа", "Вид"] + [str(y) for y in years] + ["Всего"], 1):
        put(ws, ws.cell(3, i).coordinate, t, bold=True, border=True, align=CENTER)
    keys = sorted({(k[1], k[2], k[3]) for k in cells})
    r = 4
    for cat, purpose, kind in keys:
        put(ws, f"A{r}", f"{purpose} [{CATEGORY_TITLES[cat].split(',')[0]}]", border=True, align=LEFT)
        put(ws, f"B{r}", dict(KINDS)[kind], border=True)
        for j, y in enumerate(years):
            tot = sum((v for k, v in cells.items() if k[:4] == (y, cat, purpose, kind)), Decimal(0))
            put(ws, ws.cell(r, 3 + j).coordinate, tot or None, border=True, fmt=MONEY)
        last = ws.cell(r, 3 + len(years) - 1).column_letter
        put(ws, ws.cell(r, 3 + len(years)).coordinate, f"=SUM(C{r}:{last}{r})", bold=True, border=True, fmt=MONEY)
        r += 1
    widths(ws, {"A": 46, "B": 20, **{ws.cell(1, 3 + j).column_letter: 14 for j in range(len(years) + 1)}})
