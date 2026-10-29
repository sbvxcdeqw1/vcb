from __future__ import annotations

import calendar
import re
from datetime import date, timedelta
from decimal import ROUND_DOWN, Decimal
from pathlib import Path

from .calc.ledger import build_charge_lines, select_lines, total_debt
from .calc.penalty import calc_penalty, total_penalty
from .config import rules_from
from .court.fee import state_fee
from .court.registry import CourtRegistry
from .debtors import debtor_from_account, debtor_from_row
from .docs.allocation import write_allocation
from .docs.cards import write_calc_card, write_penalty_card
from .docs.checkerboard import write_checkerboard
from .docs.claim import write_claim
from .docs.context import Case, CategoryResult
from .models import CAPREPAIR, CHARGE, OPENING, PAYMENT, PENALTY, SERVICES, Debtor, Entry, Party
from .money import D, fmt as fmt_money, q
from .parsers.onec import parse_any

NAMES = {SERVICES: "ЖКУ", CAPREPAIR: "капремонт"}


def safe(s: str) -> str:
    return re.sub(r"[^\w\-. ]+", "_", s).strip() or "debtor"


def _to_date(v) -> date | None:
    v = str(v or "").strip()
    if not v:
        return None
    return date.fromisoformat(v) if "-" in v else date(int(v[6:10]), int(v[3:5]), int(v[0:2]))


def norm_service(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip().lower())


def parse_service_list(text) -> set[str]:
    """«Домофон; видеонаблюдение» (или список) -> нормализованные названия услуг."""
    if not text:
        return set()
    parts = text if isinstance(text, (list, tuple, set)) else re.split(r"[;\n]", str(text))
    return {norm_service(x) for x in parts if norm_service(x)}


def month_end(y: int, m: int) -> date:
    return date(y, m, calendar.monthrange(y, m)[1])


def split_proportional(total: Decimal, weights: dict) -> dict:
    """Делит сумму между услугами пропорционально начислениям месяца с точностью до копейки."""
    ws = {k: max(w, Decimal(0)) for k, w in weights.items()}
    if not ws:
        return {}
    s = sum(ws.values())
    if s == 0:
        first = next(iter(ws))
        return {k: (total if k == first else Decimal(0)) for k in ws}
    raw = {k: total * w / s for k, w in ws.items()}
    base = {k: (v * 100).to_integral_value(rounding=ROUND_DOWN) / 100 for k, v in raw.items()}
    cents = int((total - sum(base.values())) * 100)
    for k in sorted(raw, key=lambda k: raw[k] - base[k], reverse=True)[:max(cents, 0)]:
        base[k] += Decimal("0.01")
    return base


def _grid_entries(entries, lines, cat) -> list[Entry]:
    """Шахматка с учётом реестра: начисления по услугам и оплаты, разнесённые на периоды (по услугам - пропорционально
    начислениям месяца)."""
    by_month: dict[tuple, dict[str, Decimal]] = {}
    for e in entries:
        if e.kind == CHARGE:
            by_month.setdefault((e.date.year, e.date.month), {})
            by_month[(e.date.year, e.date.month)][e.purpose] = by_month[(e.date.year, e.date.month)].get(
                e.purpose, Decimal(0)) + e.amount
    out: list[Entry] = []
    for ym, svc in by_month.items():
        for name, amount in svc.items():
            out.append(Entry(month_end(*ym), CHARGE, cat, name, amount))
    for ln in lines:
        paid = sum((p.amount for p in ln.portions if p.closed is not None), Decimal(0))
        if paid > 0:
            for name, part in split_proportional(paid, by_month.get((ln.year, ln.month), {"Оплата": Decimal(1)})).items():
                if part:
                    out.append(Entry(month_end(ln.year, ln.month), PAYMENT, cat, name, part))
    return out


def _card_rows(entries, ignored) -> list[tuple]:
    """Строки карточки расчёта: начисления по месяцам (все услуги вместе) и оплаты, вошедшие в расчёт."""
    skip = {ref for ref, _, _ in ignored}
    charges: dict[tuple, Decimal] = {}
    rows: list[tuple] = []
    folded: dict[date, Decimal] = {}
    for e in entries:
        if e.kind == CHARGE:
            charges[(e.date.year, e.date.month)] = charges.get((e.date.year, e.date.month), Decimal(0)) + e.amount
        elif e.kind == PAYMENT and e.ref not in skip:
            if e.ref:
                hint = f" за {e.period[1]:02d}.{e.period[0]}" if e.period else ""
                rows.append((e.date, f"Поступление на расчётный счёт от {e.date:%d.%m.%Y}", f"Оплата{hint}", Decimal(0),
                             e.amount))
            else:
                folded[e.date] = folded.get(e.date, Decimal(0)) + e.amount
    for (y, m), amount in charges.items():
        rows.append((month_end(y, m), f"Начисление за {MONTH_NAMES[m - 1]} {y}", "Начисление за услуги", amount,
                     Decimal(0)))
    for d, amount in folded.items():
        rows.append((d, "Оплата по данным шахматки", "Оплата", Decimal(0), amount))
    return sorted(rows, key=lambda r: (r[0], 0 if r[3] else 1))


MONTH_NAMES = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь",
               "декабрь"]


def _in_period(e: Entry, p_from: date, p_to: date) -> bool:
    """Расчёт строго по периоду оператора: начисления - только периода; оплата - если её месяц (из назначения)
    в периоде, а без назначения - если поступила не раньше начала периода."""
    if e.kind == PAYMENT:
        if e.period:
            return (p_from.year, p_from.month) <= e.period <= (p_to.year, p_to.month)
        return e.date >= p_from
    if e.kind == CHARGE:
        return p_from <= e.date <= p_to
    return False


def _min_due(as_of: date, years) -> date | None:
    if not years:
        return None
    try:
        return as_of.replace(year=as_of.year - int(years))
    except ValueError:  # 29 февраля
        return as_of.replace(year=as_of.year - int(years), day=28)


def _calc(entries, *, as_of, p_from, p_to, min_due, rules, rates, by_purpose=True, ignored=None):
    """Долг и пени по одной категории при данном наборе проводок. Возвращает (все строки, строки заявления, пени)."""
    all_lines = build_charge_lines(entries, as_of=as_of, due_day=rules.due_day, by_purpose=by_purpose,
                                   ignored=ignored)
    lines = select_lines(all_lines, period_from=p_from, period_to=p_to, min_due=min_due)
    items = calc_penalty(lines, as_of, rules, rates)
    return all_lines, lines, items


def build_case(card_path, row: dict | None, profile: dict, cfg: dict, as_of: date,
               registry: CourtRegistry | None, period: tuple[date | None, date | None] = (None, None),
               payments_path=None, overrides_path=None, exclude_services=None) -> Case:
    acc = parse_any(card_path)
    row = row or {}
    excluded = parse_service_list(exclude_services) | parse_service_list(row.get("exclude_services"))
    removed = sorted({e.purpose for e in acc.entries if norm_service(e.purpose) in excluded})
    if removed:
        acc.entries = [e for e in acc.entries if norm_service(e.purpose) not in excluded]
    debtor = debtor_from_row(row) if row else Debtor()
    auto = debtor_from_account(acc, default_city=(registry.default_city if registry else ""))
    for f in ("fio", "city", "street", "house", "flat", "account"):  # пустое в реестре - берём из шапки выгрузки
        if not getattr(debtor, f):
            setattr(debtor, f, getattr(auto, f))
    charges = [e for e in acc.entries if e.kind == CHARGE and e.amount > 0]
    if not charges:
        raise ValueError("в выгрузке нет начислений")
    first = min(e.date for e in charges)
    row = row or {}
    p_from = period[0] or _to_date(row.get("period_from")) or date(first.year, first.month, 1)
    p_to = period[1] or _to_date(row.get("period_to")) or max(e.date for e in charges)
    if p_from > p_to:
        raise ValueError(f"расчётный период задан неверно: {p_from:%d.%m.%Y} – {p_to:%d.%m.%Y}")
    case_period = (p_from, p_to)
    rules, rates = rules_from(cfg)
    rate = rates.at(as_of)
    claim = cfg.get("claim", {})
    min_due = _min_due(as_of, claim.get("limitation_years"))
    case = Case(acc, debtor, profile["claimant"], profile["representative"], p_from, p_to, as_of)
    case.excluded_services = removed
    if removed:
        case.notes.append("из расчёта исключены услуги: " + "; ".join(removed))
    case.notes += acc.warnings
    opening_total = sum((e.amount for e in acc.entries if e.kind == OPENING), Decimal(0))
    if opening_total:
        case.notes.append(f"сальдо на начало ({fmt_money(opening_total)} руб.) в расчёт не входит")
    pay_variants = None
    if payments_path:
        pay_variants = _apply_register(case, acc, cfg, payments_path, overrides_path, as_of)
    elif any(e.kind == PAYMENT for e in acc.entries):
        case.notes.append("реестр оплат не задан: оплаты взяты из выгрузки (без назначения платежа, по месяцам)")
    by_purpose = cfg.get("payments", {}).get("allocation", "purpose") == "purpose"
    kw = dict(as_of=as_of, p_from=p_from, p_to=p_to, min_due=min_due, rules=rules, rates=rates)
    for cat in (SERVICES, CAPREPAIR):
        entries = [e for e in acc.entries if e.category == cat and _in_period(e, p_from, p_to)]
        if not any(e.kind == CHARGE for e in entries):
            continue
        ignored: list = []
        all_lines, lines, items = _calc(entries, by_purpose=by_purpose, ignored=ignored, **kw)
        dropped = [ln for ln in all_lines if not ln.opening and p_from <= ln.accrued <= p_to and ln not in lines
                   and ln.debt > 0]
        if dropped:
            case.notes.append(f"{NAMES[cat]}: за сроком давности исключено {len(dropped)} начислений, "
                              f"{fmt_money(sum((ln.debt for ln in dropped), Decimal(0)))} руб. долга (срок платежа до "
                              f"{min_due:%d.%m.%Y})")
        shown = [e for e in entries if e.kind in (CHARGE, PAYMENT)]
        res = CategoryResult(cat, shown, lines, items, total_debt(lines), total_penalty(items), rate)
        res.all_lines, res.ignored = all_lines, ignored
        res.grid_entries = _grid_entries(entries, lines, cat)
        res.card_rows = _card_rows(entries, ignored)
        if cat == SERVICES:
            why = {ref: w for ref, _, w in ignored}
            for pay in case.payments:
                if not pay.include:
                    res.skipped.append((pay.date, pay.amount, pay.reason))
                elif pay.ref in why:
                    res.skipped.append((pay.date, pay.amount, why[pay.ref]))
                elif not _in_period(Entry(pay.date, PAYMENT, SERVICES, "", pay.amount, period=pay.period_hint),
                                    p_from, p_to):
                    res.skipped.append((pay.date, pay.amount, "вне расчётного периода"))
        if ignored:
            case.notes.append(f"{len(ignored)} оплат(ы) на {fmt_money(sum((a for _, a, _ in ignored), Decimal(0)))} руб. "
                              f"отнесены к периодам до начала расчёта и не учтены")
        case.results[cat] = res
        if pay_variants and cat == SERVICES:
            for name, pay_entries, purpose in pay_variants:
                base = [e for e in entries if e.kind != PAYMENT]
                _, ln2, it2 = _calc(base + [e for e in pay_entries if _in_period(e, p_from, p_to)],
                                    by_purpose=purpose, **kw)
                case.scenarios.append((name, total_debt(ln2), total_penalty(it2)))
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
    if registry:
        case.notes += registry.notes
    if court:
        case.court_name, case.court_address = court.name, court.address
    else:
        what = "мирового судьи" if level == "magistrate" else "районного суда"
        case.notes.append(f"не определён суд ({what}) по адресу {case.debtor.address}: нет в справочнике судов")


def _apply_register(case: Case, acc, cfg: dict, payments_path, overrides_path, as_of: date):
    """Оплаты из реестра банковских документов вместо оплат выгрузки: ИП исключаем, остальное - на старые долги."""
    from .parsers.payments import (ENFORCEMENT, KIND_TITLES, THIRD_PARTY, apply_overrides, classify, filter_by_account,
                                   parse_registers)

    pays = parse_registers(payments_path)
    pays, other = filter_by_account(pays, acc.number)
    if other:
        case.notes.append(f"в реестре оплат отобраны платежи лицевого счёта {acc.number}: {len(pays)} из "
                          f"{len(pays) + other}")
    if not pays:
        case.notes.append(f"в реестре оплат нет платежей лицевого счёта {acc.number}")
    classify(pays, cfg)
    if overrides_path:
        apply_overrides(pays, overrides_path)
    starts = [e.date for e in acc.entries if e.kind == OPENING]
    start = (max(starts) + timedelta(days=1)) if starts else min(e.date for e in acc.entries if e.kind == CHARGE).replace(day=1)
    in_range = [p for p in pays if start <= p.date <= as_of]
    case.payments = in_range
    card_paid = sum((e.amount for e in acc.entries if e.kind == PAYMENT and e.date >= start), Decimal(0))
    card_paid += sum((e.amount for e in acc.entries if e.kind == PENALTY and "оплат" in e.purpose.lower()
                      and e.date >= start), Decimal(0))
    reg_total = sum((p.amount for p in in_range), Decimal(0))
    if card_paid and abs(card_paid - reg_total) > Decimal("0.005"):
        case.notes.append(f"оплаты реестра ({fmt_money(reg_total)}) отличаются от оплат в выгрузке 1С "
                          f"({fmt_money(card_paid)}) за тот же период")
    mk = lambda ps: [Entry(p.date, PAYMENT, SERVICES, "Оплата (реестр)", p.amount, p.reason, ref=p.ref,
                           period=p.period_hint) for p in ps]
    acc.entries = [e for e in acc.entries if e.kind != PAYMENT] + mk([p for p in in_range if p.include])
    for kind in (ENFORCEMENT, THIRD_PARTY):
        sub = [p for p in in_range if p.kind == kind]
        if sub:
            tot = sum((p.amount for p in sub), Decimal(0))
            decided = "не зачтено" if not sub[0].include else "зачтено"
            case.notes.append(f"оплаты «{KIND_TITLES[kind]}»: {len(sub)} шт. на {fmt_money(tot)} руб. — {decided}"
                              + ("; нужно ваше решение (payments.third_party_policy / --payment-overrides)"
                                 if any(p.flags for p in sub) else ""))
    # сценарии для сравнения
    resident = [p for p in in_range if p.kind not in (ENFORCEMENT,)]
    only_res = [p for p in in_range if p.kind == "resident"]
    purpose = cfg.get("payments", {}).get("allocation", "purpose") == "purpose"
    return [("Все платежи реестра", mk(in_range), purpose),
            ("Без исполнительного производства", mk(resident), purpose),
            ("Только оплаты от плательщика", mk(only_res), purpose),
            ("Все платежи, без учёта периода в назначении (на самые ранние)", mk(in_range), False),
            ("Только оплаты от плательщика, на самые ранние", mk(only_res), False)]


def write_case(case: Case, out_dir: Path, claim: bool = True, allocation_report: bool = False) -> list[Path]:
    """Шахматка с учётом реестра оплат, карточка расчёта и расчёт пени; по желанию - заявление и отчёт о разноске."""
    out_dir.mkdir(parents=True, exist_ok=True)
    files = []

    def emit(name, fn):
        fn(out_dir / name)
        files.append(out_dir / name)

    emit("Шахматка_с_учетом_оплат.xlsx", lambda p: write_checkerboard(p, case))
    for cat, res in case.results.items():
        emit(f"Карточка_расчета_{NAMES[cat]}.xlsx", lambda p, res=res: write_calc_card(p, case, res))
        emit(f"Расчет_пени_{NAMES[cat]}.xlsx", lambda p, res=res: write_penalty_card(p, case, res))
    if claim:
        stem = "Заявление_на_судебный_приказ" if case.doc_kind == "order" else "Исковое_заявление"
        emit(f"{stem}.docx", lambda p: write_claim(p, case))
    if allocation_report and case.payments and SERVICES in case.results:
        emit("Разноска_оплат.xlsx", lambda p: write_allocation(p, case, case.results[SERVICES]))
    return files
