"""Реестр банковских документов (поступления на расчётный счёт) и классификация оплат."""
from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from ..money import D
from .onec import ParseError, _rows

RESIDENT, ENFORCEMENT, THIRD_PARTY = "resident", "enforcement", "third_party"
KIND_TITLES = {RESIDENT: "от плательщика", ENFORCEMENT: "исполнительное производство", THIRD_PARTY: "третье лицо / бюджет"}

PERIOD_RE = [re.compile(r"\b(\d{2})\.(\d{4})-\d{2}\.\d{4}\b"),  # 02.2018-02.2018
             re.compile(r"МЕСЯЦ:(\d{2})(\d{4})", re.I),  # МЕСЯЦ:082019
             re.compile(r"\bза\s+(\d{2})\.(\d{2,4})\b")]  # за 01.22, за 02.2025


@dataclass
class Payment:
    ref: str
    date: date
    amount: Decimal
    info: str
    kind: str = RESIDENT
    reason: str = ""
    include: bool = True
    period_hint: tuple[int, int] | None = None  # (год, месяц), указанный плательщиком (справочно)
    flags: list[str] = field(default_factory=list)


def _hint(info: str):
    for rx in PERIOD_RE:
        m = rx.search(info)
        if m:
            mm, yy = int(m.group(1)), int(m.group(2))
            yy = yy + 2000 if yy < 100 else yy
            if 1 <= mm <= 12:
                return yy, mm
    return None


def parse_register(path: str | Path) -> list[Payment]:
    rows = _rows(Path(path))
    hdr = next((i for i, r in enumerate(rows) if r and any(str(c).strip() == "Дата" for c in r)
                and any(str(c).strip() == "Сумма" for c in r)), None)
    if hdr is None:
        raise ParseError(f"{Path(path).name}: не найдена шапка реестра (Дата / Сумма / Информация)")
    head = {str(c).strip(): j for j, c in enumerate(rows[hdr]) if c}
    out = []
    for r in rows[hdr + 1:]:
        d = r[head["Дата"]]
        if isinstance(d, str) and re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", d.strip()):
            d = datetime.strptime(d.strip(), "%d.%m.%Y").date()
        elif isinstance(d, datetime):
            d = d.date()
        else:
            continue
        amount = r[head["Сумма"]]
        if amount in (None, ""):
            continue
        info = str(r[head.get("Информация", 0)] or "")
        n = len(out) + 1
        out.append(Payment(f"П{n:03d}", d, D(amount), info, period_hint=_hint(info)))
    return out


def classify(payments: list[Payment], cfg: dict) -> None:
    """Проставляет kind/include по шаблонам из настроек; ручные решения (overrides) главнее."""
    pc = cfg.get("payments", {})
    enf = [re.compile(p, re.I) for p in pc.get("enforcement_patterns", [])]
    third = [re.compile(p, re.I) for p in pc.get("third_party_patterns", [])]
    policy = pc.get("third_party_policy", "ask")  # include | exclude | ask
    for p in payments:
        hit = next((rx.pattern for rx in enf if rx.search(p.info)), None)
        if hit:
            p.kind, p.reason, p.include = ENFORCEMENT, f"признак ИП: {hit}", False
            continue
        hit = next((rx.pattern for rx in third if rx.search(p.info)), None)
        if hit:
            p.kind, p.reason = THIRD_PARTY, f"реквизиты бюджетного плательщика: {hit}"
            p.include = policy in ("include", "ask")
            if policy == "ask":
                p.flags.append("нужно решение: зачитывать ли")
            continue
        p.reason = "обычная оплата"


def apply_overrides(payments: list[Payment], path: str | Path) -> int:
    """CSV (;): дата;сумма;решение(зачесть|исключить);комментарий;номер - точечные решения по платежам.
    Пятый столбец (необязателен) - порядковый номер среди платежей с той же датой и суммой (когда их несколько).
    Возвращает число применённых решений."""
    applied = 0
    for row in csv.reader(Path(path).read_text(encoding="utf-8-sig").splitlines(), delimiter=";"):
        if not row or row[0].lower().startswith(("дата", "#")) or len(row) < 3:
            continue
        d = datetime.strptime(row[0].strip(), "%d.%m.%Y").date()
        any_amount = row[1].strip() in ("*", "")
        amt = None if any_amount else D(row[1])
        what = row[2].strip().lower()
        nth = int(row[4]) if len(row) > 4 and row[4].strip().isdigit() else None
        same = [p for p in payments if p.date == d and (any_amount or p.amount == amt)]
        for k, p in enumerate(same, 1):
            if nth is None or nth == k:
                p.include = what.startswith(("зачесть", "включ", "include"))
                p.flags = [f for f in p.flags if not f.startswith("нужно")] + [f"решение вручную: {what}"]
                applied += 1
    return applied


def parse_registers(path: str | Path) -> list[Payment]:
    """Реестр оплат: один файл или папка с несколькими файлами (платежи объединяются)."""
    path = Path(path)
    files = sorted(f for f in path.iterdir() if f.suffix.lower() in (".xls", ".xlsx") and not f.name.startswith("~")) \
        if path.is_dir() else [path]
    out: list[Payment] = []
    for f in files:
        out += parse_register(f)
    for n, p in enumerate(out, 1):
        p.ref = f"П{n:03d}"
    return out


LEAD_ACCOUNT = re.compile(r"^\s*(\d{6,})\b")


def filter_by_account(pays: list[Payment], number: str) -> tuple[list[Payment], int]:
    """В общем реестре банка оставляет платежи нужного лицевого счёта («100000000001, ФИО (л/с) / ...»).
    Если в реестре нет ни одного платежа с лицевым счётом в начале назначения - возвращает всё без изменений."""
    if not number:
        return pays, 0
    if not number.isdigit():
        # «К05»: лицевой счёт с буквами. В реестре он может быть записан так же или цифрами с нулями («00000005»)
        digits = re.sub(r"\D", "", number)

        def mine(p: Payment) -> bool:
            if re.match(rf"^\s*{re.escape(number)}\b", p.info, re.I):
                return True
            m = LEAD_ACCOUNT.match(p.info)
            return bool(m and digits and int(m.group(1)) == int(digits))

        own = [p for p in pays if mine(p)]
        return (own, len(pays) - len(own)) if own else (pays, 0)
    attributed = [p for p in pays if LEAD_ACCOUNT.match(p.info)]
    if not attributed:
        return pays, 0
    keep = [p for p in pays if (m := LEAD_ACCOUNT.match(p.info)) and m.group(1) == number]
    return keep, len(pays) - len(keep)
