"""Карточка расчёта задолженности и расчёт пени (xlsx) - по образцам заказчика."""
from __future__ import annotations

from openpyxl import Workbook

from ..models import CATEGORY_TITLES, CAPREPAIR, CHARGE, OPENING, PAYMENT
from .context import Case, CategoryResult
from .xlsx_util import *


def _head(ws, case: Case, title: str, category: str):
    put(ws, "A1", case.claimant.short_name or case.claimant.name or "Взыскатель: ________", bold=True)
    put(ws, "B2", title, bold=True)
    put(ws, "A3", f"за период с {case.period_from:%d.%m.%Y} по {case.period_to:%d.%m.%Y}")
    put(ws, "C4", f"по состоянию на {case.as_of:%d.%m.%Y}г.")
    put(ws, "A5", f"по услугам: {CATEGORY_TITLES[category]}" if category != CAPREPAIR else
        f"по взносам на капитальный ремонт")
    put(ws, "B7", f"Адрес: {case.debtor.address or case.account.address}")
    put(ws, "A9", f"ФИО собственника:  {case.debtor.fio or case.account.owner}")
    if case.debtor.area:
        put(ws, "B11", f"Общая площадь квартиры {case.debtor.area} кв.м")


def _sign(ws, row, case):
    put(ws, f"B{row}", f"Представитель {case.claimant.short_name or case.claimant.name or '________'}")
    put(ws, f"D{row}", case.representative.sign_name or "________")


def write_debt_card(path, case: Case, res: CategoryResult):
    """Образец: «Расчет задолженности собственника жилья» (Дата, Регистратор, Операция, Начислено, Оплата)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Расчет задолженности"
    _head(ws, case, "Расчет задолженности собственника жилья", res.category)
    for col, t in zip("ABCDE", ["Дата", "Регистратор", "Операция", "Начислено", "Оплата"]):
        put(ws, f"{col}13", t, bold=True, border=True, align=CENTER)
    r = 14
    for e in sorted(res.entries, key=lambda e: e.date):
        if e.kind != OPENING and not (case.period_from <= e.date <= case.as_of):
            continue
        put(ws, f"A{r}", e.date, border=True, fmt="DD.MM.YYYY")
        put(ws, f"B{r}", e.registrar, border=True, align=LEFT)
        put(ws, f"C{r}", e.purpose, border=True, align=LEFT)
        put(ws, f"D{r}", e.amount if e.kind in (CHARGE, OPENING) else None, border=True, fmt=MONEY)
        put(ws, f"E{r}", e.amount if e.kind == PAYMENT else None, border=True, fmt=MONEY)
        r += 1
    put(ws, f"A{r}", "Обороты за период", bold=True, border=True)
    put(ws, f"D{r}", f"=SUM(D14:D{r - 1})", bold=True, border=True, fmt=MONEY)
    put(ws, f"E{r}", f"=SUM(E14:E{r - 1})", bold=True, border=True, fmt=MONEY)
    put(ws, f"A{r + 1}", f"Задолженность на {case.as_of:%d.%m.%Y}", bold=True)
    put(ws, f"E{r + 1}", f"=D{r}-E{r}", bold=True, fmt=MONEY)
    _sign(ws, r + 3, case)
    widths(ws, {"A": 12, "B": 52, "C": 24, "D": 14, "E": 14})
    for _ws in wb.worksheets:
        printable(_ws, landscape=False)
    wb.save(path)


def write_penalty_card(path, case: Case, res: CategoryResult):
    """Образец: «Расчет пени по задолженности собственника жилья»."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Расчет пени"
    _head(ws, case, f"Расчет пени по задолженности собственника жилья", res.category)
    ws["B2"].value = None
    put(ws, "C1", f"Расчет пени по задолженности собственника жилья за период с {case.period_from:%d.%m.%Y} по "
                  f"{case.period_to:%d.%m.%Y}г.", bold=True)
    ws["A3"].value = None
    heads = ["Месяц", "Начислено за месяц", "Подлежит оплате до", "Дата оплаты", "Сумма оплаты", "Ставка ЦБ, %",
             "Делитель", "Доля ставки", "Расчётная сумма", "Период", "Количество дней", "Сумма пени"]
    for i, t in enumerate(heads):
        c = ws.cell(13, i + 1)
        put(ws, c.coordinate, t, bold=True, border=True, align=CENTER)
    r = 14
    first_row = r
    items_by_line = {}
    for it in res.penalty_items:
        items_by_line.setdefault(id(it.line), []).append(it)
    for ln in res.lines:
        first = True
        for it in items_by_line.get(id(ln), []):
            segs = it.segments or [None]
            for k, s in enumerate(segs):
                put(ws, f"A{r}", ln.label if first else None, border=True)
                put(ws, f"B{r}", ln.amount if first else None, border=True, fmt=MONEY)
                put(ws, f"C{r}", f"{ln.due:%d.%m.%Y}" if first else None, border=True)
                put(ws, f"D{r}", f"{it.pay_date:%d.%m.%Y}" if it.pay_date and k == 0 else None, border=True)
                put(ws, f"E{r}", it.base if it.pay_date and k == 0 else None, border=True, fmt=MONEY)
                put(ws, f"I{r}", it.base, border=True, fmt=MONEY)
                if s:
                    put(ws, f"F{r}", s.rate, border=True)
                    put(ws, f"G{r}", s.divisor, border=True)
                    put(ws, f"H{r}", f"=F{r}/G{r}", border=True, fmt="0.0000")
                    put(ws, f"J{r}", s.period, border=True)
                    put(ws, f"K{r}", s.days, border=True)
                    put(ws, f"L{r}", f"=I{r}*H{r}*K{r}/100", border=True, fmt=MONEY)
                else:
                    for col in "FGHJKL":
                        put(ws, f"{col}{r}", None, border=True)
                first = False
                r += 1
        if first:  # начисление без строк расчёта
            put(ws, f"A{r}", ln.label, border=True)
            put(ws, f"B{r}", ln.amount, border=True, fmt=MONEY)
            put(ws, f"C{r}", f"{ln.due:%d.%m.%Y}", border=True)
            r += 1
    put(ws, f"A{r}", "Итого", bold=True, border=True)
    # начисления учитываем один раз - по строкам, где заполнена колонка C (первая строка месяца)
    put(ws, f"B{r}", f'=SUMIF(C{first_row}:C{r - 1},"<>",B{first_row}:B{r - 1})', bold=True, border=True, fmt=MONEY)
    put(ws, f"E{r}", f"=SUM(E{first_row}:E{r - 1})", bold=True, border=True, fmt=MONEY)
    put(ws, f"L{r}", f"=SUM(L{first_row}:L{r - 1})", bold=True, border=True, fmt=MONEY)
    from ..money import fmt as money_fmt
    put(ws, f"A{r + 2}", f"Всего:    {money_fmt(res.debt)} руб. (задолженность)")
    put(ws, f"A{r + 3}", f"                  {money_fmt(res.penalty)} руб. (пени)")
    _sign(ws, r + 5, case)
    widths(ws, {"A": 15, "B": 13, "C": 13, "D": 12, "E": 12, "F": 9, "G": 9, "H": 10, "I": 13, "J": 20, "K": 11,
                "L": 13})
    for _ws in wb.worksheets:
        printable(_ws, landscape=True)
    wb.save(path)
