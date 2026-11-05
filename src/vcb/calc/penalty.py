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
    rate_on: str = "as_of"  # as_of - одна ставка на дату расчёта для всего расчёта (как в образцах заказчика);
    #                         payment - для погашенной части ставка на день её оплаты
    # (с, по, долг_возник_до): пени за дни окна не начисляются; если задана `долг_возник_до` - только на начисления,
    # возникшие (конец расчётного месяца) раньше этой даты
    moratoria: list[tuple] = field(default_factory=list)


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


def _overlap(a: date, b: date, rules: PenaltyRules, accrued: date | None = None) -> int:
    n = 0
    for m0, m1, *rest in rules.moratoria:
        if rest and rest[0] and accrued is not None and accrued >= rest[0]:
            continue  # долг возник после начала моратория - на него пени идут
        lo, hi = max(a, m0), min(b, m1)
        if lo <= hi:
            n += (hi - lo).days + 1
    return n


def _split_moratoria(s: date, e: date, n: int, rules: PenaltyRules, accrued: date | None):
    """Период просрочки без дней моратория: [(с, по, дней)]. Пересечение с мораторием разрезает период, и в расчёт
    попадают только куски вне моратория. `n` - дней периода по правилам счёта (1/300 в режиме calendar - ровно 60),
    разница с календарной длиной относится к последнему куску."""
    windows = sorted((m0, m1) for m0, m1, *rest in rules.moratoria
                     if not (rest and rest[0] and accrued is not None and accrued >= rest[0]))
    pieces, cur = [], s
    for m0, m1 in windows:
        if m1 < cur or m0 > e:
            continue
        if m0 > cur:
            pieces.append((cur, m0 - timedelta(1)))
        cur = max(cur, m1 + timedelta(1))
    if cur <= e:
        pieces.append((cur, e))
    if not pieces:
        return []
    total = max(n - _overlap(s, e, rules, accrued), 0)
    out = [(a, b, (b - a).days + 1) for a, b in pieces]
    a, b, k = out[-1]
    out[-1] = (a, b, k + total - sum(x[2] for x in out))
    return [x for x in out if x[2] > 0]


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
                rate_day = p.closed if rules.rate_on == "payment" else as_of
            rate = rates.at(rate_day)
            segs = []
            for s, e, n, div in _join_equal(_segments(ln.due, end, rules), rules):
                for a, b, k in _split_moratoria(s, e, n, rules, ln.accrued):
                    segs.append(Segment(a, b, k, div, rate, p.amount, p.amount * rate / div * k / 100))
            items.append(PenaltyItem(ln, p.amount, p.closed, segs))
    return items


def total_penalty(items: list[PenaltyItem]):
    return q(sum((i.penalty for i in items), Decimal(0)))


@dataclass
class ScheduleRow:
    """Строка расчёта пени по начислению: остаток долга `base` за период `segment`; оплата `pay_date`/`pay_amount`
    уменьшает остаток с этой строки (строка без периода - оплата, после которой пени дальше не идут)."""

    base: Decimal | None
    segment: Segment | None = None
    pay_date: date | None = None
    pay_amount: Decimal = Decimal(0)


def line_schedule(ln: ChargeLine, items: list[PenaltyItem], as_of: date, rules: PenaltyRules) -> list[ScheduleRow]:
    """Пени по начислению на остаток долга во времени: сначала вся сумма за свой период, после каждой оплаты -
    уменьшенная сумма за следующий период. Итог равен сумме пени по частям (`calc_penalty`)."""
    mine = [it for it in items if it.line is ln]
    if not mine:
        return []
    shift = timedelta(0 if rules.include_payment_day else 1)
    ends = [(it.pay_date - shift if it.pay_date else as_of, it) for it in mine]
    last = max(e for e, _ in ends)
    master = [(a, b, k, div) for s, e, n, div in _join_equal(_segments(ln.due, last, rules), rules)
              for a, b, k in _split_moratoria(s, e, n, rules, ln.accrued)]
    breaks = sorted({e + timedelta(1) for e, it in ends if it.pay_date})
    rows: list[ScheduleRow] = []
    for a, b, k, div in master:
        cuts = [a] + [d for d in breaks if a < d <= b]
        spans = [(x, (cuts[i + 1] - timedelta(1)) if i + 1 < len(cuts) else b) for i, x in enumerate(cuts)]
        days = [(y - x).days + 1 for x, y in spans]
        days[-1] += k - sum(days)
        for i in range(len(days) - 1, 0, -1):  # 60 дней ступени меньше календарных - недостающее с конца
            if days[i] < 0:
                days[i - 1] += days[i]
                days[i] = 0
        for (x, y), n in zip(spans, days):
            if n <= 0:
                continue
            by_rate: dict[Decimal, Decimal] = {}
            for e, it in ends:
                if e >= y and it.segments:
                    by_rate[it.segments[0].rate] = by_rate.get(it.segments[0].rate, Decimal(0)) + it.base
            for rate, base in by_rate.items():
                rows.append(ScheduleRow(base, Segment(x, y, n, div, rate, base, base * rate / div * n / 100)))
    # оплаты: в строку, с которой остаток уменьшился, или отдельной строкой без периода
    paid: dict[date, Decimal] = {}
    for it in mine:
        if it.pay_date:
            paid[it.pay_date] = paid.get(it.pay_date, Decimal(0)) + it.base
    for d in sorted(paid):
        start = d - shift + timedelta(1)
        row = next((r for r in rows if r.segment and r.segment.start == start and not r.pay_date), None)
        if row:
            row.pay_date, row.pay_amount = d, paid[d]
        else:
            pos = next((i for i, r in enumerate(rows) if r.segment and r.segment.start > start), len(rows))
            rows.insert(pos, ScheduleRow(None, None, d, paid[d]))
    return rows
