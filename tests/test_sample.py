"""Сверка с образцами заказчика: Кировоградская 180, кв. 24, расчёт на 13.04.2026."""
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from vcb.calc.ledger import build_charge_lines, total_debt
from vcb.calc.penalty import PenaltyRules, RateTable, calc_penalty, total_penalty
from vcb.models import SERVICES
from vcb.parsers.onec import parse_card

SAMPLE = Path(__file__).parent.parent / "examples" / "karta_kv24.xls"


@pytest.mark.skipif(not SAMPLE.exists(), reason="нет примера выгрузки")
def test_kv24_matches_sample():
    acc = parse_card(SAMPLE)
    assert acc.number == "00-0000000024"
    entries = [e for e in acc.entries if e.category == SERVICES]
    as_of = date(2026, 4, 13)
    lines = build_charge_lines(entries, period_from=date(2024, 6, 1), as_of=as_of)
    assert total_debt(lines) == Decimal("24805.47")
    items = calc_penalty(lines, as_of, PenaltyRules(), RateTable(Decimal("9.5")))
    assert total_penalty(items) == Decimal("4351.96")


def test_statutory_mode_and_payments():
    """Оплата гасит самое старое начисление; пени по погашенной части идут до дня оплаты."""
    from vcb.models import CHARGE, PAYMENT, Entry

    E = lambda d, k, a: Entry(d, k, SERVICES, "x", Decimal(a))
    entries = [E(date(2025, 1, 31), CHARGE, "1000"), E(date(2025, 2, 28), CHARGE, "1000"),
               E(date(2025, 4, 1), PAYMENT, "1500")]
    as_of = date(2025, 6, 30)
    lines = build_charge_lines(entries, period_from=date(2025, 1, 1), as_of=as_of)
    assert lines[0].debt == 0 and lines[1].debt == Decimal("500")
    assert total_debt(lines) == Decimal("500.00")
    items = calc_penalty(lines, as_of, PenaltyRules(mode="statutory"), RateTable(Decimal("10")))
    # январь: срок 10.02, пени с 13.03 (31-й день) по 31.03 включительно = 19 дней по 1/300
    jan = [s for i in items if i.line is lines[0] for s in i.segments]
    assert len(jan) == 1 and jan[0].days == 19 and jan[0].divisor == 300


def test_court_lookup():
    from vcb.court.registry import Court, CourtRegistry, house_matches
    from vcb.models import Debtor

    assert house_matches("2-40 чёт", "12") and not house_matches("2-40 чёт", "13")
    reg = CourtRegistry([Court("a", "magistrate", "Участок 1", territory=[
        {"city": "Пермь", "streets": [{"street": "ул. Кировоградская", "houses": "100-200"}]}])])
    assert reg.find(Debtor(city="Пермь", street="Кировоградская", house="180"), "magistrate").name == "Участок 1"
    assert reg.find(Debtor(city="Пермь", street="Кировоградская", house="20"), "magistrate") is None
