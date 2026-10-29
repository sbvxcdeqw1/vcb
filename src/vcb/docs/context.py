"""Всё, что нужно генераторам документов по одному должнику."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from ..calc.ledger import ChargeLine
from ..calc.penalty import PenaltyItem
from ..models import Account, Debtor, Entry, Party


@dataclass
class CategoryResult:
    category: str
    entries: list[Entry]  # проводки периода (начисления и оплаты)
    lines: list[ChargeLine]
    penalty_items: list[PenaltyItem]
    debt: Decimal
    penalty: Decimal
    rate: Decimal
    ignored: list = field(default_factory=list)  # оплаты вне периода расчёта: (ref, сумма, причина)
    grid_entries: list = field(default_factory=list)  # начисления и разнесённые оплаты по услугам и месяцам (шахматка)
    card_rows: list = field(default_factory=list)  # строки карточки: (дата, регистратор, операция, начислено, оплата)
    skipped: list = field(default_factory=list)  # платежи, не вошедшие в расчёт: (дата, сумма, причина)
    all_lines: list = field(default_factory=list)  # включая начальный остаток и периоды вне заявления

    @property
    def total(self) -> Decimal:
        return self.debt + self.penalty


@dataclass
class Case:
    account: Account
    debtor: Debtor
    claimant: Party
    representative: Party
    period_from: date
    period_to: date
    as_of: date
    results: dict[str, CategoryResult] = field(default_factory=dict)
    court_name: str = ""
    court_address: str = ""
    court_level: str = ""  # district | magistrate
    doc_kind: str = "claim"  # claim | order
    fee: Decimal = Decimal(0)
    legal_costs: Decimal = Decimal(0)
    legal_contract: str = ""  # «№ 25-25/юр от 27.06.2025г.»
    legal_executor: str = ""  # «ООО «СВТ»»
    notes: list[str] = field(default_factory=list)  # предупреждения для сводки
    payments: list = field(default_factory=list)  # оплаты из реестра (parsers.payments.Payment)
    excluded_services: list = field(default_factory=list)  # услуги, снятые оператором с расчёта
    scenarios: list = field(default_factory=list)  # (название, долг, пени) при разных решениях по оплатам

    @property
    def debt(self) -> Decimal:
        return sum((r.debt for r in self.results.values()), Decimal(0))

    @property
    def penalty(self) -> Decimal:
        return sum((r.penalty for r in self.results.values()), Decimal(0))

    @property
    def price(self) -> Decimal:
        return self.debt + self.penalty
