"""ФИО в родительном падеже («Взыскать с …») - по правилам; результат нужно проверять глазами."""
from __future__ import annotations

import re

_VOWELS = "аеёиоуыэюя"


def _surname(w: str, female: bool) -> str:
    lw = w.lower()
    if female:
        if re.search(r"(ова|ева|ёва|ина|ына)$", lw):
            return w[:-1] + "ой"
        if lw.endswith("ская") or lw.endswith("цкая") or lw.endswith("ная") or lw.endswith("ая"):
            return w[:-2] + "ой"
        return w  # Кох, Ким, Гусь - женские фамилии на согласную не склоняются
    if re.search(r"(ов|ев|ёв|ин|ын)$", lw):
        return w + "а"
    if lw.endswith(("ский", "цкий", "ной", "ский")) or lw.endswith("ий"):
        return w[:-2] + "ого"
    if lw.endswith("ой"):
        return w[:-2] + "ого"
    if lw.endswith("ый"):
        return w[:-2] + "ого"
    if lw.endswith(("ь", "й")):
        return w[:-1] + "я"
    if lw[-1] not in _VOWELS:
        return w + "а"
    if lw.endswith("а") and not lw.endswith(("ия",)):
        return w[:-1] + "ы"
    return w


def _first(w: str, female: bool) -> str:
    lw = w.lower()
    if lw.endswith("ия"):
        return w[:-1] + "и"
    if lw.endswith("я"):
        return w[:-1] + "и"
    if lw.endswith("а"):
        return w[:-1] + ("и" if lw[-2:-1] in "гкхжчшщ" else "ы")
    if female:
        return w  # Любовь, Ольга… (на мягкий знак женские имена: Любови)
    if lw.endswith("ий"):
        return w[:-2] + "ия"
    if lw.endswith(("й", "ь")):
        return w[:-1] + "я"
    return w + "а" if lw[-1] not in _VOWELS else w


def _patronymic(w: str, female: bool) -> str:
    lw = w.lower()
    if lw.endswith("на"):
        return w[:-1] + "ы"
    if lw.endswith("ич"):
        return w + "а"
    return w


def genitive_fio(fio: str) -> str:
    """«Иванов Иван Иванович» -> «Иванова Ивана Ивановича». Пусто, если ФИО не из 2-3 слов."""
    parts = fio.split()
    if len(parts) not in (2, 3) or not all(re.fullmatch(r"[А-ЯЁа-яё\-]+", p) for p in parts):
        return ""
    female = (len(parts) == 3 and parts[2].lower().endswith(("на", "кызы")))
    if len(parts) == 3 and not female and not parts[2].lower().endswith(("ич", "оглы")):
        return ""  # отчество не распознано - не гадаем
    out = [_surname(parts[0], female), _first(parts[1], female)]
    if len(parts) == 3:
        out.append(_patronymic(parts[2], female))
    return " ".join(out)
