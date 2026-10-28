"""Сальдо по начислениям и оплатам, распределение оплат (FIFO по сроку платежа)."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from ..models import CHARGE, OPENING, PAYMENT, Entry
from ..money import q

MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь",
          "декабрь"]


@dataclass
class Portion:
    """Часть начисления: погашена оплатой `closed` или остаётся долгом (closed=None)."""

    amount: Decimal
    closed: date | None = None
    ref: str = ""  # какой оплатой погашена


@dataclass
class ChargeLine:
    year: int
    month: int
    accrued: date
    due: date
    amount: Decimal
    portions: list[Portion] = field(default_factory=list)
    opening: bool = False  # начальный остаток (состав по месяцам неизвестен)

    @property
    def label(self) -> str:
        return "Начальный остаток" if self.opening else f"{MONTHS[self.month - 1]} {self.year}"

    @property
    def debt(self) -> Decimal:
        return sum((p.amount for p in self.portions if p.closed is None), Decimal(0))


def due_date(year: int, month: int, due_day: int) -> date:
    y, m = (year + 1, 1) if month == 12 else (year, month + 1)
    return date(y, m, due_day)


def build_charge_lines(entries: list[Entry], *, as_of: date, due_day: int = 10, with_opening: bool = False,
                       by_purpose: bool = True, ignored: list | None = None) -> list[ChargeLine]:
    """Месячные начисления + распределение оплат.

    * Начисления за месяц (в т.ч. по нескольким услугам) складываются в одну строку.
    * Оплата с указанным периодом (`Entry.period`) гасит начисление этого периода (`by_purpose`); остаток и оплаты
      без периода идут на самые ранние долги (ст. 319.1 ГК РФ). Оплата за период до начала расчёта в расчёт
      не входит (попадает в `ignored`), оплата за ещё не начисленный месяц ждёт начисления.
    * Начальный остаток в расчёт не входит (`with_opening=True` - только для сверок)."""
    kinds = (CHARGE, OPENING) if with_opening else (CHARGE,)
    grouped: dict[tuple, Decimal] = {}
    last_date: dict[tuple, date] = {}
    for e in entries:
        if e.date <= as_of and e.kind in kinds and e.amount > 0:
            key = (e.kind == OPENING, e.date.year, e.date.month)
            grouped[key] = grouped.get(key, Decimal(0)) + e.amount
            last_date[key] = max(last_date.get(key, e.date), e.date)
    credits = [e for e in entries if e.date <= as_of and (
        e.kind == PAYMENT or (e.kind in kinds and e.amount < 0))]
    lines = []
    for (is_open, y, m), amount in grouped.items():
        ln = ChargeLine(y, m, last_date[(is_open, y, m)], due_date(y, m, due_day), amount, opening=is_open)
        ln.portions = [Portion(ln.amount)]
        lines.append(ln)
    lines.sort(key=lambda c: (c.due, c.accrued))
    by_month = {(ln.year, ln.month): ln for ln in lines if not ln.opening}
    first_month = min(by_month, default=None)
    events = sorted([("c", ln.accrued, ln) for ln in lines] +
                    [("p", e.date, (abs(e.amount), e.ref, e.period)) for e in credits],
                    key=lambda x: (x[1], 0 if x[0] == "c" else 1))
    open_lines: list[ChargeLine] = []
    advance: list[tuple[Decimal, str]] = []  # оплаты без периода, поступившие раньше начислений
    pending: dict[tuple, list[tuple[Decimal, str]]] = {}  # оплаты за ещё не начисленный месяц

    def take_from(ln: ChargeLine, amount: Decimal, when: date, ref: str) -> Decimal:
        tail = ln.portions[-1]
        if tail.closed is not None or amount <= 0:
            return amount
        take = min(amount, tail.amount)
        if take < tail.amount:
            tail.amount -= take
            ln.portions.insert(len(ln.portions) - 1, Portion(take, when, ref))
        else:
            tail.closed, tail.ref = when, ref
        return amount - take

    def pay_oldest(amount: Decimal, when: date, ref: str) -> Decimal:
        for ln in open_lines:
            if amount <= 0:
                break
            amount = take_from(ln, amount, when, ref)
        return amount

    def pay(amount: Decimal, when: date, ref: str, target) -> Decimal:
        if target and by_purpose:
            ln = by_month.get(target)
            if ln is not None and ln.accrued > when:
                pending.setdefault(target, []).append((amount, ref))
                return Decimal(0)
            if ln is not None:
                amount = take_from(ln, amount, when, ref)
            elif first_month and target < first_month:
                if ignored is not None:
                    ignored.append((ref, amount, "относится к периоду до начала расчёта"))
                return Decimal(0)
        return pay_oldest(amount, when, ref)

    for kind, when, payload in events:
        if kind == "c":
            open_lines.append(payload)
            key = (payload.year, payload.month)
            for amt, ref in pending.pop(key, []):
                rest = take_from(payload, amt, when, ref)
                if rest > 0 and (rest := pay_oldest(rest, when, ref)) > 0:
                    advance.append((rest, ref))
            advance[:] = [(rest, ref) for amt, ref in advance if (rest := pay_oldest(amt, when, ref)) > 0]
        else:
            amt, ref, target = payload
            if (rest := pay(amt, when, ref, target)) > 0:
                advance.append((rest, ref))
        open_lines[:] = [ln for ln in open_lines if ln.debt > 0]
    return lines


def select_lines(lines: list[ChargeLine], *, period_from: date | None = None, period_to: date | None = None,
                 min_due: date | None = None) -> list[ChargeLine]:
    """Строки для заявления: период и срок давности (по сроку платежа). Начальный остаток не входит никогда."""
    out = []
    for ln in lines:
        if ln.opening:
            continue
        if period_from and ln.accrued < period_from:
            continue
        if period_to and ln.accrued > period_to:
            continue
        if min_due and ln.due < min_due:
            continue
        out.append(ln)
    return out


def total_debt(lines: list[ChargeLine]) -> Decimal:
    return q(sum((ln.debt for ln in lines), Decimal(0)))
