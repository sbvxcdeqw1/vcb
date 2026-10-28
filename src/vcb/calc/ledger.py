"""Сальдо по начислениям и оплатам, распределение оплат (FIFO по сроку платежа)."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from ..models import CHARGE, PAYMENT, Account, Entry
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

    @property
    def label(self) -> str:
        return f"{MONTHS[self.month - 1]} {self.year}"

    @property
    def debt(self) -> Decimal:
        return sum((p.amount for p in self.portions if p.closed is None), Decimal(0))


def due_date(year: int, month: int, due_day: int) -> date:
    y, m = (year + 1, 1) if month == 12 else (year, month + 1)
    return date(y, m, due_day)


def build_charge_lines(entries: list[Entry], *, period_from: date, as_of: date, due_day: int = 10,
                       period_to: date | None = None) -> list[ChargeLine]:
    """Начисления периода + FIFO-распределение оплат (и отрицательных перерасчётов)."""
    period_to = period_to or as_of
    charges = [e for e in entries if e.kind == CHARGE and e.amount > 0 and period_from <= e.date <= period_to]
    credits = [e for e in entries if (e.kind == PAYMENT or (e.kind == CHARGE and e.amount < 0))
               and period_from <= e.date <= as_of]
    lines = sorted((ChargeLine(e.date.year, e.date.month, e.date, due_date(e.date.year, e.date.month, due_day),
                               e.amount) for e in charges), key=lambda c: (c.due, c.accrued))
    for ln in lines:
        ln.portions = [Portion(ln.amount)]
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


def total_debt(lines: list[ChargeLine]) -> Decimal:
    return q(sum((ln.debt for ln in lines), Decimal(0)))
