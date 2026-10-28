"""Данные должников: реестр (csv/xlsx с набором полей) + запасной разбор адреса из шапки выгрузки."""
from __future__ import annotations

import csv
import re
from pathlib import Path

from .models import Account, Debtor

FIELD_ALIASES = {
    "file": ["file", "файл", "выгрузка"], "fio": ["fio", "фио", "должник"], "fio_gen": ["fio_gen", "фио_род", "фио (р.п.)"],
    "city": ["city", "город", "населенный пункт"], "street": ["street", "улица"], "house": ["house", "дом"],
    "flat": ["flat", "кв", "квартира"], "area": ["area", "площадь"], "account": ["account", "лс", "лицевой счет"],
    "birth_date": ["birth_date", "дата рождения"], "birth_place": ["birth_place", "место рождения"],
    "share": ["share", "доля"],
}
_ALIAS = {a: k for k, al in FIELD_ALIASES.items() for a in al}


def read_registry(path: str | Path) -> list[dict]:
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl

        rows = [[("" if c is None else str(c)).strip() for c in r]
                for r in openpyxl.load_workbook(path, data_only=True).worksheets[0].iter_rows(values_only=True)]
    else:
        text = path.read_text(encoding="utf-8-sig")
        dialect = csv.Sniffer().sniff(text.splitlines()[0], delimiters=";,\t")
        rows = list(csv.reader(text.splitlines(), dialect))
    head = [_ALIAS.get(h.strip().lower(), h.strip().lower()) for h in rows[0]]
    return [dict(zip(head, r)) for r in rows[1:] if any(x for x in r)]


def debtor_from_row(row: dict) -> Debtor:
    known = {k: row.get(k, "") for k in ("fio", "city", "street", "house", "flat", "area", "account", "birth_date",
                                         "birth_place", "share")}
    d = Debtor(**known)
    d.extra = {k: v for k, v in row.items() if k not in known and k != "file" and v}
    return d


def debtor_from_account(acc: Account, default_city: str = "") -> Debtor:
    """«Пермь г, Кировоградская, д.180,  24» -> город, улица, дом, квартира (best effort)."""
    parts = [p.strip() for p in acc.address.split(",") if p.strip()]
    d = Debtor(fio=acc.owner, account=acc.number, city=default_city)
    if parts:
        m = re.match(r"(.+?)\s+г$", parts[0]) or re.match(r"г\.?\s*(.+)$", parts[0])
        if m:
            d.city = m.group(1).strip().capitalize() if m.group(1).isupper() else m.group(1).strip()
            parts = parts[1:]
    if parts:
        d.street = parts[0]
    for p in parts[1:]:
        m = re.match(r"д\.?\s*(\S+)", p)
        if m:
            d.house = m.group(1)
        elif re.fullmatch(r"(кв\.?\s*)?[\dа-яА-Я/\-]+", p):
            d.flat = re.sub(r"^кв\.?\s*", "", p)
    return d
