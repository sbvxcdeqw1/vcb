"""Отчёт о разноске оплат: что зачтено, что исключено, на какие начисления перенесено."""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from openpyxl import Workbook

from ..money import fmt
from ..parsers.payments import KIND_TITLES
from .context import Case
from .xlsx_util import *


def write_allocation(path, case: Case, res):
    skipped = {ref: why for ref, _, why in getattr(res, "ignored", [])}
    refs = defaultdict(list)  # ref -> [(начисление, сумма)]
    paid_by = defaultdict(list)  # id строки -> [(оплата, сумма)]
    for ln in res.all_lines:
        for p in ln.portions:
            if p.closed is not None and p.ref:
                refs[p.ref].append((ln.label, p.amount))
                paid_by[id(ln)].append((p.ref, p.closed, p.amount))
    wb = Workbook()
    ws = wb.active
    ws.title = "Оплаты"
    put(ws, "A1", f"Разноска оплат — {case.debtor.fio or case.account.owner}, ЛС {case.account.number}", bold=True)
    put(ws, "A2", "Оплата идёт на период из назначения платежа, остаток и оплаты без периода — на самые ранние долги; "
                  "платежи по исполнительному производству в зачёт не берутся.")
    heads = ["№", "Дата", "Сумма", "Тип платежа", "Зачтено", "Основание", "Период по назначению", "Разнесено на",
             "Примечание", "Назначение платежа"]
    for i, t in enumerate(heads, 1):
        put(ws, ws.cell(4, i).coordinate, t, bold=True, border=True, align=CENTER)
    r = 5
    for p in sorted(case.payments, key=lambda p: (p.date, p.ref)):
        parts = refs.get(p.ref, [])
        put(ws, f"A{r}", p.ref, border=True)
        put(ws, f"B{r}", p.date, border=True, fmt="DD.MM.YYYY")
        put(ws, f"C{r}", p.amount, border=True, fmt=MONEY)
        put(ws, f"D{r}", KIND_TITLES[p.kind], border=True, align=LEFT)
        put(ws, f"E{r}", "да" if p.include else "нет", bold=not p.include, border=True)
        put(ws, f"F{r}", p.reason, border=True, align=LEFT)
        put(ws, f"G{r}", f"{p.period_hint[1]:02d}.{p.period_hint[0]}" if p.period_hint else "", border=True)
        where = "; ".join(f"{lbl}: {fmt(a)}" for lbl, a in parts)
        if p.ref in skipped:
            where = f"не учтена: {skipped[p.ref]}"
        put(ws, f"H{r}", where if p.include else "—", border=True, align=LEFT)
        put(ws, f"I{r}", "; ".join(p.flags), border=True, align=LEFT)
        put(ws, f"J{r}", p.info[:300], border=True, align=LEFT)
        r += 1
    put(ws, f"A{r}", "Итого зачтено", bold=True, border=True)
    put(ws, f"C{r}", f'=SUMIF(E5:E{r - 1},"да",C5:C{r - 1})', bold=True, border=True, fmt=MONEY)
    put(ws, f"A{r + 1}", "Итого не зачтено", bold=True, border=True)
    put(ws, f"C{r + 1}", f'=SUMIF(E5:E{r - 1},"нет",C5:C{r - 1})', bold=True, border=True, fmt=MONEY)
    widths(ws, {"A": 7, "B": 12, "C": 12, "D": 24, "E": 9, "F": 34, "G": 12, "H": 48, "I": 28, "J": 70})
    ws.freeze_panes = "A5"
    printable(ws, landscape=True)

    w2 = wb.create_sheet("Начисления")
    for i, t in enumerate(["Начисление", "Срок оплаты", "Сумма", "Погашено оплатами", "Остаток долга", "В заявлении"], 1):
        put(w2, w2.cell(1, i).coordinate, t, bold=True, border=True, align=CENTER)
    included = {id(ln) for ln in res.lines}
    for r, ln in enumerate(res.all_lines, 2):
        put(w2, f"A{r}", ln.label, border=True)
        put(w2, f"B{r}", ln.due, border=True, fmt="DD.MM.YYYY")
        put(w2, f"C{r}", ln.amount, border=True, fmt=MONEY)
        put(w2, f"D{r}", "; ".join(f"{ref} {c:%d.%m.%Y} {fmt(a)}" for ref, c, a in paid_by.get(id(ln), [])),
            border=True, align=LEFT)
        put(w2, f"E{r}", ln.debt, border=True, fmt=MONEY)
        put(w2, f"F{r}", "да" if id(ln) in included else "нет (давность / вне периода / остаток)", border=True)
    widths(w2, {"A": 20, "B": 13, "C": 12, "D": 70, "E": 13, "F": 34})
    w2.freeze_panes = "A2"

    w3 = wb.create_sheet("Сценарии")
    for i, t in enumerate(["Вариант учёта оплат", "Долг", "Пени", "Итого"], 1):
        put(w3, w3.cell(1, i).coordinate, t, bold=True, border=True, align=CENTER)
    for r, (name, debt, pen) in enumerate(case.scenarios, 2):
        put(w3, f"A{r}", name, border=True)
        put(w3, f"B{r}", debt, border=True, fmt=MONEY)
        put(w3, f"C{r}", pen, border=True, fmt=MONEY)
        put(w3, f"D{r}", f"=B{r}+C{r}", border=True, fmt=MONEY)
    n = len(case.scenarios) + 3
    put(w3, f"A{n}", "Принято в заявлении (с учётом настроек и ручных решений)", bold=True, border=True)
    put(w3, f"B{n}", res.debt, bold=True, border=True, fmt=MONEY)
    put(w3, f"C{n}", res.penalty, bold=True, border=True, fmt=MONEY)
    put(w3, f"D{n}", f"=B{n}+C{n}", bold=True, border=True, fmt=MONEY)
    widths(w3, {"A": 58, "B": 14, "C": 14, "D": 14})
    wb.save(path)
