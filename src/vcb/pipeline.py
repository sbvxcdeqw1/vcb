from __future__ import annotations

import re
from dataclasses import replace
from datetime import date
from decimal import Decimal
from pathlib import Path

from .calc.ledger import build_charge_lines, total_debt
from .calc.penalty import calc_penalty, total_penalty
from .config import rules_from
from .court.fee import state_fee
from .court.registry import CourtRegistry
from .debtors import debtor_from_account, debtor_from_row
from .docs.cards import write_debt_card, write_penalty_card
from .docs.checkerboard import write_checkerboard
from .docs.claim import write_claim
from .docs.context import Case, CategoryResult
from .models import CAPREPAIR, CHARGE, SERVICES, Debtor, Party
from .money import D, q
from .parsers.onec import parse_card

NAMES = {SERVICES: "ЖКУ", CAPREPAIR: "капремонт"}


def safe(s: str) -> str:
    return re.sub(r"[^\w\-. ]+", "_", s).strip() or "debtor"


def build_case(card_path, row: dict | None, profile: dict, cfg: dict, as_of: date,
               registry: CourtRegistry | None, period: tuple[date | None, date | None] = (None, None)) -> Case:
    acc = parse_card(card_path)
    debtor = debtor_from_row(row) if row else Debtor()
    auto = debtor_from_account(acc, default_city=(registry.default_city if registry else ""))
    for f in ("fio", "city", "street", "house", "flat", "account"):  # пустое в реестре - берём из шапки выгрузки
        if not getattr(debtor, f):
            setattr(debtor, f, getattr(auto, f))
    charges = [e for e in acc.entries if e.kind == CHARGE and e.amount > 0]
    if not charges:
        raise ValueError("в выгрузке нет начислений")
    p_from = period[0] or date(min(e.date for e in charges).year, min(e.date for e in charges).month, 1)
    p_to = period[1] or max(e.date for e in charges)
    rules, rates = rules_from(cfg)
    rate = rates.at(as_of)
    case = Case(acc, debtor, profile["claimant"], profile["representative"], p_from, p_to, as_of)
    for cat in (SERVICES, CAPREPAIR):
        entries = [e for e in acc.entries if e.category == cat]
        if not any(e.kind == CHARGE for e in entries):
            continue
        lines = build_charge_lines(entries, period_from=p_from, period_to=p_to, as_of=as_of, due_day=rules.due_day)
        items = calc_penalty(lines, as_of, rules, rates)
        case.results[cat] = CategoryResult(cat, [e for e in entries if e.kind in (CHARGE, "payment")], lines, items,
                                           total_debt(lines), total_penalty(items), rate)
    legal = profile.get("legal") or {}
    case.legal_costs = D(legal.get("costs", 0))
    case.legal_executor, case.legal_contract = legal.get("executor", ""), legal.get("contract", "")
    _choose_court(case, cfg, registry)
    return case


def _choose_court(case: Case, cfg: dict, registry: CourtRegistry | None):
    c = cfg["court"]
    price = case.price
    kind = c["doc"]
    if kind == "auto":
        kind = "order" if price <= D(c["order_limit"]) else "claim"
    case.doc_kind = kind
    if kind == "order":
        level = "magistrate"
    else:
        level = c["claim_court"]
        if level == "auto":
            level = "magistrate" if price <= D(c["magistrate_claim_limit"]) else "district"
    case.court_level = level
    case.fee = state_fee(price, cfg, level=level, order=(kind == "order"))
    court = registry.find(case.debtor, level) if registry else None
    if court:
        case.court_name, case.court_address = court.name, court.address
    else:
        what = "мирового судьи" if level == "magistrate" else "районного суда"
        case.notes.append(f"не определён суд ({what}) по адресу {case.debtor.address}: нет в справочнике судов")


def write_case(case: Case, out_dir: Path) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    files = []
    stem = "Заявление_на_судебный_приказ" if case.doc_kind == "order" else "Исковое_заявление"
    for name, fn in [(f"{stem}.docx", lambda p: write_claim(p, case)),
                     ("Шахматка.xlsx", lambda p: write_checkerboard(p, case))]:
        fn(out_dir / name)
        files.append(out_dir / name)
    for cat, res in case.results.items():
        for name, fn in [(f"Карточка_расчета_{NAMES[cat]}.xlsx", write_debt_card),
                         (f"Расчет_пени_{NAMES[cat]}.xlsx", write_penalty_card)]:
            fn(out_dir / name, case, res)
            files.append(out_dir / name)
    return files
