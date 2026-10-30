"""Предзаполнение данных должника из загруженных файлов (шахматка / карточка и реестр оплат)."""
from __future__ import annotations

import re
from pathlib import Path

from .debtors import debtor_from_account
from .names import genitive_fio

FIO_RE = r"[А-ЯЁ][а-яё\-]+(?: [А-ЯЁ][а-яё\-]+){1,2}"


def _fio_from_register(path: str, number: str) -> str:
    """ФИО плательщика из назначения платежей реестра («<лицевой счёт>, ФИО (л/с) / …»)."""
    from .parsers.payments import filter_by_account, parse_registers

    try:
        pays, _ = filter_by_account(parse_registers(path), number)
    except Exception:  # noqa: BLE001
        return ""
    digits = re.sub(r"\D", "", number)
    for p in pays:
        m = re.match(rf"\s*(?:\d+|[А-ЯA-Z]\d+)[,;\s]+({FIO_RE})\b", p.info)
        if m and (not digits or digits.lstrip("0") in re.sub(r"\D", "", p.info.split(m.group(1))[0]).lstrip("0")):
            return m.group(1)
    return ""


def autofill(chess_path: str = "", register_path: str = "", region: str = "perm_krai") -> tuple[dict, list[str]]:
    """Возвращает (поля должника, пояснения). Что не нашлось - не попадает в словарь."""
    from .parsers.onec import parse_any

    values: dict[str, str] = {}
    notes: list[str] = []
    acc = None
    if chess_path and Path(chess_path).is_file():
        acc = parse_any(chess_path)
        d = debtor_from_account(acc)
        for key in ("fio", "city", "street", "house", "flat"):
            if getattr(d, key):
                values[key] = getattr(d, key)
        if acc.number:
            values["account"] = acc.number
    if not values.get("fio") and register_path and Path(register_path).exists():
        fio = _fio_from_register(register_path, acc.number if acc else "")
        if fio:
            values["fio"] = fio
            notes.append("ФИО взято из назначения платежей реестра")
    if values.get("fio"):
        gen = genitive_fio(values["fio"])
        if gen:
            values["fio_gen"] = gen
            notes.append("ФИО в родительном падеже составлено по правилам — проверьте")
    if values.get("street") or values.get("city"):
        try:
            from .court.registry import CourtRegistry
            from .models import Debtor

            reg = CourtRegistry.load(region)
            court = reg.find(Debtor(**{k: values.get(k, "") for k in ("city", "street", "house", "flat")}),
                             "magistrate")
            if court:
                values["court"] = court.name
                notes += reg.notes
            elif reg.candidates:
                notes.append("участок не определён однозначно; кандидаты: " + "; ".join(c.name for c in reg.candidates[:4]))
        except Exception:  # noqa: BLE001
            pass
    return values, notes
