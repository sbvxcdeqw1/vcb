"""Сверка «Карточки расчетов» с «Шахматкой» и приведение шахматки к карточке.

Карточка - итог по лицевому счёту (все услуги, точные даты платежей), шахматка - разбивка по услугам.
Если в шахматке нет какой-то услуги, помесячные суммы не сойдутся: разницу показываем отдельной строкой.
"""
from __future__ import annotations

import copy
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from .models import CHARGE, OPENING, PAYMENT, Account, Entry
from .parsers.onec import classify_category

UNALLOCATED = "Не распределено по услугам (разница с карточкой)"
ZERO = Decimal(0)


@dataclass
class Report:
    months: list[tuple[int, int]] = field(default_factory=list)  # период, в котором сравниваем
    charge_diff: dict[tuple[int, int], Decimal] = field(default_factory=dict)  # карточка − шахматка
    payment_diff: dict[tuple[int, int], Decimal] = field(default_factory=dict)
    card_charges: Decimal = ZERO
    chess_charges: Decimal = ZERO
    card_payments: Decimal = ZERO
    chess_payments: Decimal = ZERO
    card_balance_start: Decimal | None = None  # сальдо карточки на начало периода шахматки (услуги + пени)
    chess_opening: Decimal = ZERO
    chess_gap_months: list[tuple[int, int]] = field(default_factory=list)  # есть в карточке, нет в шахматке
    notes: list[str] = field(default_factory=list)

    @property
    def opening_diff(self) -> Decimal | None:
        return None if self.card_balance_start is None else self.card_balance_start - self.chess_opening

    @property
    def principal_gap(self) -> Decimal:
        """Сколько долга по начислениям/оплатам периода есть в карточке и нет в шахматке."""
        return (self.card_charges - self.chess_charges) - (self.card_payments - self.chess_payments)


def _sum(entries, kind, months):
    d = defaultdict(Decimal)
    for e in entries:
        if e.kind == kind and (e.date.year, e.date.month) in months:
            d[(e.date.year, e.date.month)] += e.amount
    return d


def compare(card: Account, chess: Account) -> Report:
    r = Report()
    ch_months = sorted({(e.date.year, e.date.month) for e in chess.entries if e.kind in (CHARGE, PAYMENT)})
    if not ch_months:
        r.notes.append("в шахматке нет проводок")
        return r
    lo, hi = ch_months[0], ch_months[-1]
    card_months = sorted({(e.date.year, e.date.month) for e in card.entries if e.kind in (CHARGE, PAYMENT)})
    r.months = [m for m in card_months if lo <= m <= hi]
    r.chess_gap_months = [m for m in card_months if m > hi]
    mset = set(r.months)
    a, b = _sum(card.entries, CHARGE, mset), _sum(chess.entries, CHARGE, mset)
    pa, pb = _sum(card.entries, PAYMENT, mset), _sum(chess.entries, PAYMENT, mset)
    for m in r.months:
        if abs(a[m] - b[m]) > Decimal("0.005"):
            r.charge_diff[m] = a[m] - b[m]
        if abs(pa[m] - pb[m]) > Decimal("0.005"):
            r.payment_diff[m] = pa[m] - pb[m]
    r.card_charges, r.chess_charges = sum(a.values(), ZERO), sum(b.values(), ZERO)
    r.card_payments, r.chess_payments = sum(pa.values(), ZERO), sum(pb.values(), ZERO)
    r.chess_opening = sum((e.amount for e in chess.entries if e.kind == OPENING), ZERO)
    if card.saldo_start is not None:
        # сальдо карточки на 1-е число первого месяца шахматки (оно включает и пени)
        bal = card.saldo_start
        for e in card.entries:
            if e.date < date(lo[0], lo[1], 1):
                bal += e.amount if e.kind in (CHARGE, "penalty") and not _is_pen_payment(e) else -e.amount
        r.card_balance_start = bal
    if r.chess_gap_months:
        r.notes.append(f"шахматка заканчивается {hi[1]:02d}.{hi[0]}, а карточка идёт до "
                       f"{r.chess_gap_months[-1][1]:02d}.{r.chess_gap_months[-1][0]}")
    return r


def _is_pen_payment(e: Entry) -> bool:
    return e.kind == "penalty" and "оплат" in e.purpose.lower()


def reconciled(card: Account, chess: Account) -> Account:
    """Шахматка + строка «не распределено» так, чтобы помесячные итоги сошлись с карточкой."""
    rep = compare(card, chess)
    out = copy.copy(chess)
    out.entries = list(chess.entries)
    out.warnings = list(chess.warnings)
    last = lambda ym: date(ym[0], ym[1], 28) if ym[1] == 2 else date(ym[0], ym[1], 30 if ym[1] in (4, 6, 9, 11) else 31)
    for m, d in rep.charge_diff.items():
        out.entries.append(Entry(last(m), CHARGE, classify_category(UNALLOCATED), UNALLOCATED, d))
    for m, d in rep.payment_diff.items():
        out.entries.append(Entry(last(m), PAYMENT, classify_category(UNALLOCATED), UNALLOCATED, d))
    for m in rep.chess_gap_months:  # месяцы, которых в шахматке ещё нет: берём из карточки целиком
        for e in card.entries:
            if (e.date.year, e.date.month) == m and e.kind in (CHARGE, PAYMENT):
                out.entries.append(Entry(e.date, e.kind, e.category, UNALLOCATED, e.amount))
    return out


def format_report(rep: Report) -> str:
    from .money import fmt

    L = ["СВЕРКА КАРТОЧКИ И ШАХМАТКИ", ""]
    if rep.months:
        L.append(f"Сравниваемый период: {rep.months[0][1]:02d}.{rep.months[0][0]} – {rep.months[-1][1]:02d}.{rep.months[-1][0]}")
    L.append(f"Начислено:  карточка {fmt(rep.card_charges)}  шахматка {fmt(rep.chess_charges)}  разница {fmt(rep.card_charges - rep.chess_charges)}")
    L.append(f"Оплачено:   карточка {fmt(rep.card_payments)}  шахматка {fmt(rep.chess_payments)}  разница {fmt(rep.card_payments - rep.chess_payments)}")
    L.append(f"Долг по начислениям периода, которого нет в шахматке: {fmt(rep.principal_gap)}")
    if rep.card_balance_start is not None:
        L.append(f"Начальный остаток: шахматка {fmt(rep.chess_opening)}, сальдо карточки на начало периода "
                 f"{fmt(rep.card_balance_start)} (в нём и пени), разница {fmt(rep.opening_diff)}")
    L.append(f"Месяцев с расхождением по начислениям: {len(rep.charge_diff)} из {len(rep.months)}; "
             f"по оплатам: {len(rep.payment_diff)}")
    L += [""] + rep.notes
    return "\n".join(L)
