"""Разбор выгрузки 1С «Шахматка»: услуги по строкам, месяцы по столбцам (Начислено / Оплата), остатки."""
from __future__ import annotations

import calendar
import re
from datetime import date, timedelta
from pathlib import Path

from ..models import CHARGE, OPENING, PAYMENT, Account, Entry
from ..money import D
from .onec import ParseError, _rows, classify_category

RU_MONTHS = {m: i for i, m in enumerate(["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август",
                                          "сентябрь", "октябрь", "ноябрь", "декабрь"], 1)}
MONTH_RE = re.compile(r"^([А-Яа-я]+)\s+(\d{4})$")


def is_chess(rows: list[list]) -> bool:
    text = " ".join(str(c) for r in rows[:12] for c in r if c)
    return "Шахматка" in text and "Начальный остаток" in text


def _month_end(y: int, m: int) -> date:
    return date(y, m, calendar.monthrange(y, m)[1])


def parse_chess(path: str | Path, rows: list[list] | None = None) -> Account:
    path = Path(path)
    rows = rows if rows is not None else _rows(path)
    acc = Account(source=path.name)
    # шапка
    for r in rows[:6]:
        cells = [str(c).strip() for c in r if c not in (None, "")]
        if len(cells) == 1 and not cells[0].startswith(("Шахматка", "за период", "Услуга")):
            if not acc.number:
                acc.number = cells[0]  # «551010009822, МО г.Перми»
            elif not acc.address:
                acc.address = cells[0]
    num, _, org = acc.number.partition(",")
    acc.number, acc.owner = num.strip(), ""
    hdr = next((i for i, r in enumerate(rows) if r and str(r[0] or "").strip() == "Услуга"), None)
    if hdr is None:
        raise ParseError(f"{path.name}: не найдена строка «Услуга»")
    top, sub = rows[hdr], rows[hdr + 1]
    cols = {}  # индекс столбца -> ("opening"|"closing"|"total"|(год, месяц), "Начислено"|"Оплата")
    cur = None
    for j, h in enumerate(top):
        h = str(h).strip() if h not in (None, "") else ""
        if h:
            m = MONTH_RE.match(h)
            if m and m.group(1).lower() in RU_MONTHS:
                cur = (int(m.group(2)), RU_MONTHS[m.group(1).lower()])
            elif h == "Начальный остаток":
                cur = "opening"
            elif h == "Итог":
                cur = "total"
            elif h == "Конечный остаток":
                cur = "closing"
            elif j > 0:
                cur = None
        kind = str(sub[j]).strip() if j < len(sub) and sub[j] not in (None, "") else ""
        if cur is not None and j >= 2:
            cols[j] = (cur, kind)
    for r in rows[hdr + 2:]:
        name = str(r[0] or "").strip()
        if not name:
            continue
        if name.lower().startswith("итог"):
            break
        cat = classify_category(name)
        opening = D(r[next(j for j, (c, _) in cols.items() if c == "opening")] or 0) if any(
            c == "opening" for c, _ in cols.values()) else D(0)
        charged = paid = D(0)
        for j, (c, kind) in cols.items():
            v = r[j] if j < len(r) else None
            if v in (None, "") or not isinstance(c, tuple):
                continue
            amount = D(v)
            when = _month_end(*c)
            if kind == "Оплата" or amount < 0 and kind != "Начислено":
                acc.entries.append(Entry(when, PAYMENT, cat, name, abs(amount)))
                paid += abs(amount)
            else:
                acc.entries.append(Entry(when, CHARGE, cat, name, amount))
                charged += amount
        if opening:
            first = min((c for c, _ in cols.values() if isinstance(c, tuple)), default=None)
            when = _month_end(*first) - timedelta(days=_month_end(*first).day)  # конец предыдущего месяца
            acc.entries.append(Entry(when, OPENING, cat, name, opening))
        closing_col = next((j for j, (c, _) in cols.items() if c == "closing"), None)
        if closing_col is not None and r[closing_col] not in (None, ""):
            if abs(D(r[closing_col]) - (opening + charged - paid)) > D("0.01"):
                acc.warnings.append(f"{name}: конечный остаток {r[closing_col]} не сходится с расчётным "
                                    f"{opening + charged - paid}")
    return acc
