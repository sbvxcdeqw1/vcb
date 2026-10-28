"""Разбор выгрузки 1С «Карточка расчетов» (xls/xlsx)."""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

from ..models import CAPREPAIR, CHARGE, PAYMENT, PENALTY, SERVICES, Account, Entry
from ..money import D

CAPREPAIR_RE = re.compile(r"капитальн|капремонт|кап\.\s*ремонт", re.I)
LS_RE = re.compile(r"по лицевому счету\s*([\w\-]+)\s*,?\s*(.*)", re.I)
DATE_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")


class ParseError(ValueError):
    pass


def _rows(path: Path) -> list[list]:
    if path.suffix.lower() == ".xls":
        import xlrd

        wb = xlrd.open_workbook(str(path))
        sh = wb.sheet_by_index(0)
        out = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_DATE:
                    row.append(xlrd.xldate_as_datetime(cell.value, wb.datemode))
                elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    row.append(None)
                else:
                    row.append(cell.value)
            out.append(row)
        return out
    import openpyxl

    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str) and DATE_RE.match(v.strip()):
        return datetime.strptime(v.strip(), "%d.%m.%Y").date()
    return None


def _num(v):
    if v in (None, ""):
        return None
    return D(v)


def classify_category(*texts: str) -> str:
    return CAPREPAIR if CAPREPAIR_RE.search(" ".join(texts)) else SERVICES


def parse_card(path: str | Path) -> Account:
    path = Path(path)
    rows = _rows(path)
    acc = Account(source=path.name)
    head = None
    for i, row in enumerate(rows):
        cells = [str(c).strip() for c in row if c not in (None, "")]
        if not cells:
            continue
        if head is None:
            text = " ".join(cells)
            m = LS_RE.search(text)
            if m:
                acc.number, acc.owner = m.group(1), m.group(2).strip()
                # адрес - следующая непустая строка
                for nxt in rows[i + 1:i + 3]:
                    t = " ".join(str(c).strip() for c in nxt if c not in (None, ""))
                    if t and not t.lower().startswith("дата"):
                        acc.address = t
                        break
            if cells[0].lower() == "дата" and any("регистратор" in c.lower() for c in cells):
                head = {c.strip().lower(): j for j, c in enumerate(row) if isinstance(c, str)}
            continue
        d = _date(row[head["дата"]])
        if d is None:
            continue  # «Сальдо», «Обороты» и пр.
        registrar = str(row[head["регистратор"]] or "")
        oper = str(row[head.get("операция", head["регистратор"])] or "")
        charged = _num(row[head["начислено"]]) if "начислено" in head else None
        paid = _num(row[head["оплата"]]) if "оплата" in head else None
        cat = classify_category(oper, registrar)
        if "пени" in oper.lower() or "пени" in registrar.lower():
            kind, amount = PENALTY, charged or paid
        elif paid:
            kind, amount = PAYMENT, paid
        elif charged is not None:
            kind, amount = CHARGE, charged
        else:
            continue
        if amount is None:
            continue
        acc.entries.append(Entry(d, kind, cat, oper.strip(), amount, registrar.strip()))
    if head is None:
        raise ParseError(f"{path.name}: не найдена таблица с колонками Дата/Регистратор/Операция")
    return acc
