"""Пени по ч. 14 ст. 155 ЖК РФ: 1/300 ключевой ставки с 31-го дня просрочки, 1/130 с 91-го дня.

Режимы счёта дней:
  calendar  - как в образцах заказчика: льготный месяц и два месяца по 1/300 считаются календарными
              (срок 10.07 -> пени с 11.08, 1/300 - 60 дней, 1/130 с 11.10);
  statutory - строго по дням закона (31-й..90-й день - 1/300, с 91-го - 1/130).
"""
from __future__ import annotations

import calendar
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal

from ..money import D, q
from .ledger import ChargeLine


def add_months(d: date, n: int) -> date:
    y, m = divmod(d.year * 12 + d.month - 1 + n, 12)
    return date(y, m + 1, min(d.day, calendar.monthrange(y, m + 1)[1]))


@dataclass
class RateTable:
    """Ключевая ставка, % годовых. Без истории - одна ставка на любую дату (как в образцах)."""

    default: Decimal = Decimal("9.5")
    history: list[tuple[date, Decimal]] = field(default_factory=list)  # (действует с, ставка)
    cap: Decimal | None = None  # потолок ставки для пени за ЖКУ (в образцах заказчика - 9,5%)
    cap_from: date | None = None
    cap_to: date | None = None

    def at(self, d: date) -> Decimal:
        rate = None
        for since, r in sorted(self.history):
            if since <= d:
                rate = r
        rate = self.default if rate is None else rate
        if self.cap is not None and (self.cap_from is None or d >= self.cap_from) and (
                self.cap_to is None or d <= self.cap_to):
            rate = min(rate, self.cap)
        return rate


@dataclass
class PenaltyRules:
    mode: str = "calendar"
    divisor1: int = 300
    divisor2: int = 130
    tier1_days: int = 60
    grace_days: int = 30  # только для statutory
    due_day: int = 10
    include_payment_day: bool = False
    moratoria: list[tuple[date, date]] = field(default_factory=list)  # пени не начисляются


@dataclass
class Segment:
    start: date
    end: date
    days: int
    divisor: int
    rate: Decimal
    base: Decimal
    penalty: Decimal  # без округления

    @property
    def period(self) -> str:
        return f"{self.start:%d.%m.%y}-{self.end:%d.%m.%y}"


@dataclass
class PenaltyItem:
    """Строка расчёта пени: начисление (или его погашенная часть) + сегменты по ставкам."""

    line: ChargeLine
    base: Decimal
    pay_date: date | None
    segments: list[Segment]

    @property
    def penalty(self) -> Decimal:
        return sum((s.penalty for s in self.segments), Decimal(0))


def _overlap(a: date, b: date, rules: PenaltyRules) -> int:
    n = 0
    for m0, m1 in rules.moratoria:
        lo, hi = max(a, m0), min(b, m1)
        if lo <= hi:
            n += (hi - lo).days + 1
    return n


def _segments(due: date, last: date, rules: PenaltyRules) -> list[tuple[date, date, int, int]]:
    """[(с, по, дней, делитель)] для просрочки от срока `due` до последнего дня начисления `last`."""
    if rules.mode == "calendar":
        s1, s2 = add_months(due, 1) + timedelta(1), add_months(due, 3) + timedelta(1)
        out = []
        if last >= s1:
            e1 = min(s2 - timedelta(1), last)
            n1 = rules.tier1_days if last >= s2 - timedelta(1) else (e1 - s1).days + 1
            out.append((s1, e1, n1, rules.divisor1))
        if last >= s2:
            out.append((s2, last, (last - s2).days + 1, rules.divisor2))
        return out
    s1 = due + timedelta(rules.grace_days + 1)
    s2 = due + timedelta(rules.grace_days + rules.tier1_days + 1)
    out = []
    if last >= s1:
        e1 = min(s2 - timedelta(1), last)
        out.append((s1, e1, (e1 - s1).days + 1, rules.divisor1))
    if last >= s2:
        out.append((s2, last, (last - s2).days + 1, rules.divisor2))
    return out


def _join_equal(segs: list, rules: PenaltyRules) -> list:
    """Если обе ставки совпадают (пени по капремонту: только 1/300), периоды объединяются в один."""
    if rules.divisor1 == rules.divisor2 and len(segs) == 2:
        (a, b, n1, d), (_, c, n2, _d) = segs
        return [(a, c, n1 + n2, d)]
    return segs


def calc_penalty(lines: list[ChargeLine], as_of: date, rules: PenaltyRules, rates: RateTable) -> list[PenaltyItem]:
    items = []
    for ln in lines:
        for p in ln.portions:
            if p.closed is None:
                end, rate_day = as_of, as_of
            else:
                end = p.closed if rules.include_payment_day else p.closed - timedelta(1)
                rate_day = p.closed
            rate = rates.at(rate_day)
            segs = []
            for s, e, n, div in _join_equal(_segments(ln.due, end, rules), rules):
                n = max(n - _overlap(s, e, rules), 0)
                if n:
                    segs.append(Segment(s, e, n, div, rate, p.amount, p.amount * rate / div * n / 100))
            items.append(PenaltyItem(ln, p.amount, p.closed, segs))
    return items


def total_penalty(items: list[PenaltyItem]):
    return q(sum((i.penalty for i in items), Decimal(0)))
