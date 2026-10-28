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


def build_charge_lines(entries: list[Entry], *, as_of: date, due_day: int = 10) -> list[ChargeLine]:
    """Все начисления (и начальный остаток) + FIFO-распределение оплат и отрицательных перерасчётов.

    Оплаты гасят самое старое начисление (ст. 319.1 ГК РФ), поэтому FIFO идёт по всей истории; отбор строк
    в заявление - отдельно, `select_lines`."""
    charges = [e for e in entries if e.date <= as_of and (
        (e.kind == CHARGE and e.amount > 0) or (e.kind == OPENING and e.amount > 0))]
    credits = [e for e in entries if e.date <= as_of and (
        e.kind == PAYMENT or (e.kind in (CHARGE, OPENING) and e.amount < 0))]
    lines = []
    for e in charges:
        ln = ChargeLine(e.date.year, e.date.month, e.date, due_date(e.date.year, e.date.month, due_day), e.amount,
                        opening=(e.kind == OPENING))
        ln.portions = [Portion(ln.amount)]
        lines.append(ln)
    lines.sort(key=lambda c: (c.due, c.accrued))
    events = sorted([("c", ln.accrued, ln) for ln in lines] + [("p", e.date, abs(e.amount)) for e in credits],
                    key=lambda x: (x[1], 0 if x[0] == "c" else 1))
    open_lines: list[ChargeLine] = []
    advance = Decimal(0)  # оплата, поступившая раньше начисления

    def apply(amount: Decimal, when: date) -> Decimal:
        for ln in open_lines:
            if amount <= 0:
                break
            tail = ln.portions[-1]
            if tail.closed is not None:
                continue
            take = min(amount, tail.amount)
            amount -= take
            if take < tail.amount:
                tail.amount -= take
                ln.portions.insert(len(ln.portions) - 1, Portion(take, when))
            else:
                tail.closed = when
        return amount

    for kind, when, payload in events:
        if kind == "c":
            open_lines.append(payload)
            if advance > 0:
                advance = apply(advance, when)
        else:
            advance = apply(payload, when)
        open_lines[:] = [ln for ln in open_lines if ln.debt > 0]
    return lines


def select_lines(lines: list[ChargeLine], *, period_from: date | None = None, period_to: date | None = None,
                 min_due: date | None = None, include_opening: bool = False) -> list[ChargeLine]:
    """Строки для заявления: период, срок давности (по сроку платежа), начальный остаток - по решению."""
    out = []
    for ln in lines:
        if ln.opening:
            if not include_opening:
                continue
        else:
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
