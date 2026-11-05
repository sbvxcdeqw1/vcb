"""Сверка с образцами заказчика (обезличенные файлы кладутся в examples/, в репозиторий не входят)."""
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from vcb.calc.ledger import build_charge_lines, select_lines, total_debt
from vcb.calc.penalty import PenaltyRules, RateTable, calc_penalty, total_penalty
from vcb.models import SERVICES
from vcb.parsers.onec import parse_card

SAMPLE = Path(__file__).parent.parent / "examples" / "karta_kv24.xls"


@pytest.mark.skipif(not SAMPLE.exists(), reason="нет примера выгрузки")
def test_kv24_matches_sample():
    acc = parse_card(SAMPLE)
    assert acc.number.startswith("00-")  # лицевой счёт карточки 1С
    entries = [e for e in acc.entries if e.category == SERVICES]
    as_of = date(2026, 4, 13)
    lines = select_lines(build_charge_lines(entries, as_of=as_of), period_from=date(2024, 6, 1))
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
    lines = select_lines(build_charge_lines(entries, as_of=as_of), period_from=date(2025, 1, 1))
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
        {"city": "Пермь", "streets": [{"street": "ул. Ленина", "houses": "100-200"}]}])])
    assert reg.find(Debtor(city="Пермь", street="Ленина", house="180"), "magistrate").name == "Участок 1"
    assert reg.find(Debtor(city="Пермь", street="Ленина", house="20"), "magistrate") is None


CHESS = Path(__file__).parent.parent / "examples" / "shakhmatka_kv98.xlsx"


@pytest.mark.skipif(not CHESS.exists(), reason="нет примера шахматки")
def test_chess_parse_and_full_debt_matches_closing_balance():
    from vcb.models import OPENING
    from vcb.parsers.onec import parse_any

    acc = parse_any(CHESS)
    assert acc.number.isdigit() and len(acc.number) == 12 and not acc.warnings
    lines = build_charge_lines(acc.entries, as_of=date(2026, 7, 31), with_opening=True)  # с начальным остатком
    assert total_debt(lines) == Decimal("79793.41")  # «Конечный остаток» образца
    assert sum(e.amount for e in acc.entries if e.kind == OPENING) == Decimal("5859.31")


CARD98 = Path(__file__).parent.parent / "examples" / "karta_kv98.xls"


@pytest.mark.skipif(not (CHESS.exists() and CARD98.exists()), reason="нет примеров кв. 98")
def test_card_vs_chess_reconciliation():
    from vcb.parsers.onec import parse_any
    from vcb.reconcile import compare, reconciled

    card, chess = parse_any(CARD98), parse_any(CHESS)
    assert card.saldo_start == Decimal("6060.19") and card.saldo_end == Decimal("162748.26")
    rep = compare(card, chess)
    # «Итог:» начального остатка в образце шахматки (6 935,03) - сальдо карточки с пенями на 31.05.2020
    assert rep.card_balance_start == Decimal("6935.03")
    assert rep.opening_diff == Decimal("1075.72")
    assert rep.principal_gap == Decimal("11269.54")  # в шахматке нет части услуг
    merged = reconciled(card, chess)
    by = lambda acc, k: sum(e.amount for e in acc.entries if e.kind == k and e.date >= date(2020, 6, 1))
    assert by(merged, "charge") == by(card, "charge") and by(merged, "payment") == by(card, "payment")


def test_payment_classification_and_overrides(tmp_path):
    from vcb.config import load_config
    from vcb.parsers.payments import ENFORCEMENT, RESIDENT, THIRD_PARTY, Payment, apply_overrides, classify

    mk = lambda i, d, a, t: Payment(f"П{i}", d, Decimal(a), t)
    ps = [mk(1, date(2019, 6, 24), "0.62", "Погашение долга Судебный приказ N2-1968/19 от 16.04.2019, 37522/19/59003-ИП"),
          mk(2, date(2020, 1, 20), "3000", "100000000001 12.2019-12.2019 к/плата=3000.00"),
          mk(3, date(2025, 6, 5), "978.63", "(991-0501-1540121500-244,л/с02991018855,225) ПД за 01.22")]
    classify(ps, load_config())
    assert [p.kind for p in ps] == [ENFORCEMENT, RESIDENT, THIRD_PARTY]
    assert [p.include for p in ps] == [False, True, True] and not ps[2].flags  # третье лицо зачитывается
    ov = tmp_path / "ov.csv"
    ov.write_text("дата;сумма;решение;комментарий\n05.06.2025;978.63;исключить;ИП\n", encoding="utf-8")
    apply_overrides(ps, ov)
    assert ps[2].include is False


def test_payments_go_to_oldest_charges_regardless_of_period_in_purpose():
    from vcb.models import CHARGE, PAYMENT, Entry

    E = lambda d, k, a, ref="": Entry(d, k, SERVICES, "x", Decimal(a), ref=ref)
    entries = [E(date(2025, 1, 31), CHARGE, "100"), E(date(2025, 2, 28), CHARGE, "100"),
               E(date(2025, 3, 31), CHARGE, "100"), E(date(2025, 4, 15), PAYMENT, "150", "П1")]
    lines = build_charge_lines(entries, as_of=date(2025, 5, 1))
    assert [ln.debt for ln in lines] == [Decimal(0), Decimal(50), Decimal(100)]
    assert {p.ref for ln in lines for p in ln.portions if p.closed} == {"П1"}


def test_payment_goes_to_period_from_purpose_and_unmatched_to_oldest():
    from vcb.models import CHARGE, PAYMENT, Entry

    E = lambda d, k, a, ref="", per=None: Entry(d, k, SERVICES, "x", Decimal(a), ref=ref, period=per)
    base = [E(date(2025, 1, 31), CHARGE, "100"), E(date(2025, 2, 28), CHARGE, "100"), E(date(2025, 3, 31), CHARGE, "100")]
    ignored = []
    pays = [E(date(2025, 4, 15), PAYMENT, "130", "П1", (2025, 3)),   # март 100 + остаток 30 на январь
            E(date(2025, 4, 16), PAYMENT, "40", "П2", (2024, 12))]   # период до начала расчёта - не учитывается
    lines = build_charge_lines(base + pays, as_of=date(2025, 5, 1), ignored=ignored)
    assert [ln.debt for ln in lines] == [Decimal(70), Decimal(100), Decimal(0)]
    assert [i[0] for i in ignored] == ["П2"]
    # режим «на самые ранние»: период в назначении игнорируется
    lines = build_charge_lines(base + pays[:1], as_of=date(2025, 5, 1), by_purpose=False)
    assert [ln.debt for ln in lines] == [Decimal(0), Decimal(70), Decimal(100)]


def test_payment_for_not_yet_accrued_month_waits_for_accrual():
    from vcb.models import CHARGE, PAYMENT, Entry

    E = lambda d, k, a, ref="", per=None: Entry(d, k, SERVICES, "x", Decimal(a), ref=ref, period=per)
    entries = [E(date(2025, 1, 31), CHARGE, "100"), E(date(2025, 2, 28), CHARGE, "100"),
               E(date(2025, 2, 10), PAYMENT, "100", "П1", (2025, 2))]  # оплата за февраль до его начисления
    lines = build_charge_lines(entries, as_of=date(2025, 3, 1))
    assert [ln.debt for ln in lines] == [Decimal(100), Decimal(0)]


def test_period_is_strictly_operator_defined():
    from vcb.models import CHARGE, PAYMENT, Entry
    from vcb.pipeline import _in_period

    E = lambda d, k, per=None: Entry(d, k, SERVICES, "x", Decimal(1), period=per)
    a, b = date(2024, 6, 1), date(2025, 5, 31)
    assert _in_period(E(date(2024, 6, 30), CHARGE), a, b) and not _in_period(E(date(2024, 5, 31), CHARGE), a, b)
    assert not _in_period(E(date(2025, 6, 30), CHARGE), a, b)
    assert _in_period(E(date(2025, 8, 1), PAYMENT, (2025, 5)), a, b)  # за май 2025, оплачено позже
    assert not _in_period(E(date(2025, 1, 1), PAYMENT, (2024, 5)), a, b)  # за период до начала
    assert not _in_period(E(date(2024, 5, 1), PAYMENT), a, b) and _in_period(E(date(2024, 7, 1), PAYMENT), a, b)


def test_real_perm_register_resolves_addresses_from_its_own_data():
    from vcb.court.registry import CourtRegistry
    from vcb.models import Debtor

    reg = CourtRegistry.load("perm_krai")
    ms = next(c for c in reg.courts if c.level == "magistrate" and "№ 1 Кировского района" in c.name and c.territory)
    street = ms.territory[0]["streets"][0]
    house = street["houses"].split(",")[0].split("-")[0].strip() if street["houses"] != "all" else "1"
    d = Debtor(city="Пермь", street=street["street"], house=house)
    assert reg.find(d, "magistrate").name == ms.name  # участок определяется по списку улиц и домов
    dc = reg.find(d, "district")
    assert dc.name == "Кировский районный суд г. Перми" and dc.address.startswith("614113")
    # населённый пункт без разобранных улиц - по тексту описания границ
    assert reg.find(Debtor(city="Александровск", street="Ленина", house="5"), "district").name.startswith("Александровский")


def test_rate_cap_and_history():
    from vcb.config import load_config, rules_from

    rules, rates = rules_from(load_config())
    assert rates.at(date(2026, 4, 13)) == Decimal("9.5")  # ключевая выше, но действует потолок 9,5%
    assert rates.at(date(2021, 6, 1)) == Decimal("5.00")  # до потолка - фактическая ставка ЦБ
    assert rules.moratoria == [(date(2020, 4, 1), date(2020, 12, 31), None),
                               (date(2022, 4, 1), date(2022, 10, 1), None)]  # по умолчанию применяются оба
    cfg = load_config()
    cfg["penalty"]["moratoria_enabled"] = []
    assert rules_from(cfg)[0].moratoria == []


@pytest.mark.skipif(not SAMPLE.exists(), reason="нет примера выгрузки")
def test_full_pipeline_reproduces_sample_with_default_config():
    from vcb.config import load_config
    from vcb.models import Party
    from vcb.pipeline import build_case

    prof = {"claimant": Party(), "representative": Party(), "legal": {}}
    c = build_case(SAMPLE, None, prof, load_config(), date(2026, 4, 13), None)
    res = c.results[SERVICES]
    assert (res.debt, res.penalty) == (Decimal("24805.47"), Decimal("4351.96"))


def test_gui_builds_cli_arguments():
    pytest.importorskip("tkinter")
    import vcb_gui

    argv = vcb_gui.build_argv({"source": "a.xls", "out": "o", "region": "perm_krai", "payments": "p.xlsx",
                               "as_of": "13.04.2026", "period_from": "01.10.2023", "doc": "claim"})
    assert argv[:2] == ["a.xls", "-o"] and "--no-interactive" in argv
    assert argv[argv.index("--as-of") + 1] == "2026-04-13"
    assert argv[argv.index("--period-from") + 1] == "2023-10-01" and "--period-to" not in argv
    assert argv[argv.index("--payments") + 1] == "p.xlsx" and argv[argv.index("--doc") + 1] == "claim"
    with pytest.raises(ValueError):
        vcb_gui.build_argv({"source": ""})
    with pytest.raises(ValueError):
        vcb_gui.build_argv({"source": "a", "as_of": "31/12"})


def test_version_is_consistent():
    import re

    import vcb

    pyproject = (Path(__file__).parent.parent / "pyproject.toml").read_text(encoding="utf-8")
    assert re.search(rf'^version = "{re.escape(vcb.__version__)}"', pyproject, flags=re.M)
    changelog = (Path(__file__).parent.parent / "CHANGELOG.md").read_text(encoding="utf-8")
    assert f"## {vcb.__version__} " in changelog


def test_typography_and_print_setup(tmp_path):
    from openpyxl import load_workbook

    from vcb.config import load_config
    from vcb.docs.claim import NBSP, typo
    from vcb.models import Party
    from vcb.pipeline import build_case, write_case

    assert typo("24 805,47 руб. по ст. 155 № 5, д. 180") == f"24{NBSP}805,47{NBSP}руб. по ст.{NBSP}155 №{NBSP}5, д.{NBSP}180"
    if not SAMPLE.exists():
        return
    prof = {"claimant": Party(), "representative": Party(), "legal": {}}
    case = build_case(SAMPLE, None, prof, load_config(), date(2026, 4, 13), None)
    write_case(case, tmp_path)
    for name, landscape in (("Карточка_расчета_ЖКУ.xlsx", False), ("Расчет_пени_ЖКУ.xlsx", True), ("Расчет_начислений_и_оплат_ЖКУ.xlsx", False)):
        ws = load_workbook(next(tmp_path.glob(name.replace(".xlsx", "_*.xlsx")))).worksheets[0]
        assert ws.page_setup.paperSize == 9 and (ws.page_setup.orientation == "landscape") == landscape
        assert ws.print_title_rows or name.startswith("Расчет_начислений")  # шахматка: блоки со своими шапками
    import docx

    doc = docx.Document(next(tmp_path.glob("*.docx")))
    sec = doc.sections[0]
    assert round(sec.left_margin.cm, 1) == 3.0 and round(sec.right_margin.cm, 1) == 1.0
    assert doc.styles["Normal"].font.size.pt == 14 and doc.styles["Normal"].font.name == "Times New Roman"


def test_every_xlsx_has_signature_block_with_representative(tmp_path):
    from openpyxl import load_workbook

    from vcb.config import load_config
    from vcb.models import Party
    from vcb.pipeline import build_case, write_case

    if not SAMPLE.exists():
        pytest.skip("нет примера выгрузки")
    prof = {"claimant": Party(short_name="ООО «Тест»"), "representative": Party(name="Петров Пётр Петрович",
                                                                            sign_name="П.П.Петров"), "legal": {}}
    case = build_case(SAMPLE, None, prof, load_config(), date(2026, 4, 13), None)
    write_case(case, tmp_path)
    files = sorted(tmp_path.glob("*.xlsx"))
    assert len(files) >= 3
    for f in files:
        for ws in load_workbook(f).worksheets:
            text = "\n".join(str(c.value) for row in ws.iter_rows() for c in row if c.value)
            assert "Представитель ООО «Тест»" in text and "/ П.П.Петров /" in text, f"{f.name}/{ws.title}"
            assert "(подпись)" not in text and "Тело долга и оплаты" not in text, f"{f.name}/{ws.title}"


def test_split_proportional_keeps_cents():
    from vcb.pipeline import split_proportional

    parts = split_proportional(Decimal("100.00"), {"a": Decimal("1"), "b": Decimal("1"), "c": Decimal("1")})
    assert sum(parts.values()) == Decimal("100.00") and sorted(parts.values()) == [Decimal("33.33")] * 2 + [Decimal("33.34")]
    assert split_proportional(Decimal("5"), {"a": Decimal(0), "b": Decimal(0)}) == {"a": Decimal("5"), "b": Decimal(0)}


def test_gui_flags_for_services_and_outputs():
    pytest.importorskip("tkinter")
    import vcb_gui

    argv = vcb_gui.build_argv({"source": "a.xlsx", "exclude_services": "Домофон; видеонаблюдение", "no_claim": True,
                               "allocation_report": True})
    assert argv[argv.index("--exclude-services") + 1] == "Домофон; видеонаблюдение"
    assert "--no-claim" in argv and "--allocation-report" in argv


REG98 = Path(__file__).parent.parent / "examples" / "oplaty_kv98.xlsx"


@pytest.mark.skipif(not (CHESS.exists() and REG98.exists()), reason="нет примеров кв. 98")
def test_chess_plus_register_pipeline(tmp_path):
    from openpyxl import load_workbook

    from vcb.config import load_config
    from vcb.models import Party
    from vcb.parsers.chess import list_services
    from vcb.pipeline import build_case, write_case

    services = dict(list_services(CHESS))
    assert "Домофон" in services and services["Домофон"] == Decimal("2590")
    prof = {"claimant": Party(), "representative": Party(), "legal": {}}
    as_of = date(2026, 7, 31)
    full = build_case(CHESS, None, prof, load_config(), as_of, None, (None, None), REG98)
    less = build_case(CHESS, None, prof, load_config(), as_of, None, (None, None), REG98,
                      exclude_services="домофон;  Видеонаблюдение")
    assert less.results[SERVICES].excluded_services == ["Домофон", "видеонаблюдение"]
    assert less.results[SERVICES].debt < full.results[SERVICES].debt
    res = full.results[SERVICES]
    # карточка (начисления - оплаты) и шахматка с разнесёнными оплатами сходятся с долгом расчёта
    assert sum(r[3] for r in res.card_rows) - sum(r[4] for r in res.card_rows) == res.debt
    grid = res.grid_entries
    assert sum(e.amount for e in grid if e.kind == "charge") - sum(e.amount for e in grid if e.kind == "payment") == res.debt
    files = write_case(full, tmp_path, claim=False)
    # в имени каждого файла: документ, месяц и год расчёта, улица, дом, квартира, фамилия должника
    import re

    from vcb.pipeline import file_suffix

    tail = f"_{file_suffix(full)}.xlsx"
    assert re.fullmatch(r"_[а-я]+_20\d\d_[^_]+(_[^_]+)*_д[^_]+_кв[^_]+_[А-ЯЁ][а-яё\-]+\.xlsx", tail), tail
    assert sorted(f.name for f in files) == sorted([f"Карточка_расчета_ЖКУ{tail}", f"Расчет_пени_ЖКУ{tail}",
                                                     f"Расчет_начислений_и_оплат_ЖКУ{tail}"])
    chess = load_workbook(tmp_path / f"Расчет_начислений_и_оплат_ЖКУ{tail}")
    assert chess.sheetnames == ["Шахматка"]  # всё на одном листе; месяцы переносятся блоками вниз
    ws = load_workbook(tmp_path / f"Карточка_расчета_ЖКУ{tail}").worksheets[0]
    heads = [c.value for row in ws.iter_rows() for c in row if c.value in ("Начислено, руб.", "Задолженность, руб.")]
    assert heads == ["Начислено, руб."]  # столбца нарастающей «Задолженность» в карточке нет
    text = "\n".join(str(c.value) for row in ws.iter_rows() for c in row if c.value)
    assert "Сальдо на" not in text and "не вошедшие" not in text and "Исключены из расчёта" not in text
    assert f"Задолженность на {as_of:%d.%m.%Y}" in text


# ---------------------------------------------------------------- два пакета: ЖКУ + капитальный ремонт
def make_chess(path, number, address, services: dict, months):
    """Минимальная «Шахматка» 1С: services = {название: сумма в месяц}, months = [(год, месяц), ...]."""
    from openpyxl import Workbook

    names = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь",
             "декабрь"]
    ws = Workbook().active
    ws["A2"], ws["A3"], ws["A4"], ws["A5"] = "Шахматка", "за период", f"{number}, МО г.Перми", address
    ws["A6"], ws["C6"] = "Услуга", "Начальный остаток"
    for j, (y, m) in enumerate(months):
        ws.cell(6, 4 + j).value = f"{names[m - 1].capitalize()} {y}"
        ws.cell(7, 4 + j).value = "Начислено"
    for i, (name, amount) in enumerate(services.items()):
        ws.cell(8 + i, 1).value = name
        for j in range(len(months)):
            ws.cell(8 + i, 4 + j).value = amount
    ws.cell(8 + len(services), 1).value = "Итог:"
    ws.parent.save(path)


MONTHS_2024 = [(2024, m) for m in range(6, 13)] + [(2025, 1), (2025, 2)]


def two_packages(tmp_path):
    g, k = tmp_path / "gku_98.xlsx", tmp_path / "kr_98_other_name.xlsx"
    make_chess(g, "100000000001", "г.ПЕРМЬ, Тестовая, д.1, кв. 2", {"Содержание": 1000, "Домофон": 50},
               MONTHS_2024)
    make_chess(k, "К100", "Пермь г, Тестовая, д.1, кв.2", {"Взнос на капремонт": 600}, MONTHS_2024)
    return g, k


def test_two_packages_combined_claim(tmp_path):
    import docx

    from vcb.config import load_config
    from vcb.models import CAPREPAIR, Party
    from vcb.pipeline import Package, build_case, write_case

    g, k = two_packages(tmp_path)
    prof = {"claimant": Party(short_name="ООО «Тест»", name="Общество «Тест»"), "representative": Party(), "legal": {}}
    cfg = load_config()
    as_of = date(2025, 4, 15)
    both = build_case(g, None, prof, cfg, as_of, None, extra_packages=[Package(CAPREPAIR, k)])
    only_g = build_case(g, None, prof, cfg, as_of, None)
    only_k = build_case(k, None, prof, cfg, as_of, None, category=CAPREPAIR)
    assert set(both.results) == {SERVICES, CAPREPAIR} and list(only_k.results) == [CAPREPAIR]
    # логика расчёта та же: суммы объединяются, от общей суммы - цена иска и госпошлина
    assert both.debt == only_g.debt + only_k.debt == Decimal("9450") * 1 + Decimal("5400")
    assert both.penalty == only_g.penalty + only_k.penalty and both.price == both.debt + both.penalty
    # пени по капремонту - только 1/300 (ч. 14.1 ст. 155 ЖК РФ), по ЖКУ - 1/300 и затем 1/130
    assert {s.divisor for it in both.results[CAPREPAIR].penalty_items for s in it.segments} == {300}
    assert {s.divisor for it in both.results[SERVICES].penalty_items for s in it.segments} == {300, 130}
    files = write_case(both, tmp_path / "out")
    names = sorted(f.name.split("_апрель")[0] + ".xlsx" if f.suffix == ".xlsx" else f.name for f in files)
    for cat in ("ЖКУ", "капремонт"):
        assert f"Расчет_начислений_и_оплат_{cat}.xlsx" in names and f"Карточка_расчета_{cat}.xlsx" in names
        assert f"Расчет_пени_{cat}.xlsx" in names
    assert sum(n.endswith(".docx") for n in names) == 1  # один общий иск/приказ
    text = "\n".join(p.text for p in docx.Document(next((tmp_path / "out").glob("*.docx"))).paragraphs)
    text = text.replace("\xa0", " ")
    assert "и по оплате взносов на капитальный ремонт" in text and "статьей 169" in text and "частью 14.1 статьи 155" in text
    assert "частью 14 статьи 155" in text and "протокола общего собрания" in text
    text_k = "\n".join(p.text for p in docx.Document(
        next(f for f in write_case(only_k, tmp_path / "out_k") if f.suffix == ".docx")).paragraphs).replace("\xa0", " ")
    assert "частью 14.1 статьи 155" in text_k and "частью 14 статьи 155" not in text_k
    assert "по оплате взносов на капитальный ремонт общего имущества" in text_k


def test_pair_units_by_address(tmp_path):
    from vcb.pipeline import pair_units

    g, k = two_packages(tmp_path)
    gd, kd = tmp_path / "g", tmp_path / "k"
    gd.mkdir(), kd.mkdir()
    g.rename(gd / "a.xlsx"), k.rename(kd / "completely_different.xlsx")
    make_chess(kd / "alone.xlsx", "1", "г.ПЕРМЬ, Другая, д.1, кв. 5", {"Взнос": 1}, MONTHS_2024[:2])
    units = pair_units(gd, kd)
    assert [(a.name if a else None, b.name if b else None) for a, b in units] == [
        ("a.xlsx", "completely_different.xlsx"), (None, "alone.xlsx")]


KR_REAL = Path(__file__).parent.parent / "examples" / "shakhmatka_kr.xls"


@pytest.mark.skipif(not KR_REAL.exists(), reason="нет примера шахматки капремонта")
def test_real_caprepair_chess():
    from vcb.config import load_config
    from vcb.models import CAPREPAIR, PENALTY, Party
    from vcb.parsers.chess import list_services
    from vcb.parsers.onec import parse_any
    from vcb.pipeline import build_case

    acc = parse_any(KR_REAL)
    assert acc.number and not acc.number.isdigit() and len(acc.owner.split()) == 3 and not acc.warnings  # ЛС «К..»
    assert any(e.kind == PENALTY for e in acc.entries)  # строка «Пени» - не долг
    assert [n for n, _ in list_services(KR_REAL)] == ["Взнос на капремонт"]
    prof = {"claimant": Party(), "representative": Party(), "legal": {}}
    case = build_case(KR_REAL, None, prof, load_config(), date(2026, 10, 30), None, category=CAPREPAIR)
    res = case.results[CAPREPAIR]
    assert res.debt == Decimal("47959.11") and res.penalty > 0  # начальный остаток 50 750,77 в расчёт не входит
    assert {s.divisor for it in res.penalty_items for s in it.segments} == {300}
    assert case.debtor.fio == acc.owner and case.debtor.city == "Пермь"  # город из «ПЕРМЬ, ул.…» без «г.»
    assert case.debtor.street.startswith("ул. ") and case.debtor.house and case.debtor.flat


def test_court_override_and_common_district():
    from vcb.court.registry import Court, CourtRegistry
    from vcb.models import Debtor

    text = "г. Пермь: в границах от ул. Тестовой до ул. Примерной"
    reg = CourtRegistry([
        Court("d1", "district", "Тестовый районный суд г. Перми"),
        Court("m1", "magistrate", "Судебный участок № 1 Тестового района", parent="d1", territory_text=text),
        Court("m2", "magistrate", "Судебный участок № 2 Тестового района", parent="d1", territory_text=text)])
    d = Debtor(city="Пермь", street="Тестовая", house="48")  # дом внутри описанной границы: участок неоднозначен
    assert reg.find(d, "magistrate") is None and len(reg.candidates) == 2
    assert reg.find(d, "district").name == "Тестовый районный суд г. Перми"  # все кандидаты - одного района
    assert "№ 2 Тестового района" in reg.find(d, "magistrate", "№ 2 Тестового района").name


def test_register_filter_by_lettered_account():
    from vcb.parsers.payments import Payment, filter_by_account

    mk = lambda i, info: Payment(f"П{i}", date(2025, 1, i), Decimal("1"), info)
    pays = [mk(1, "00000005, Иванов (л/с) / оплата"), mk(2, "00000009, Петров (л/с) / оплата"), mk(3, "К05, Иванов")]
    own, other = filter_by_account(pays, "К05")
    assert [p.ref for p in own] == ["П1", "П3"] and other == 1
    assert filter_by_account(pays, "Х77")[0] == pays  # счёт не найден - реестр не фильтруется


@pytest.mark.skipif(not (KR_REAL.exists() and (Path(__file__).parent.parent / "examples" / "oplaty_kr.xls").exists()),
                    reason="нет примеров капремонта")
def test_real_caprepair_register_is_enforcement_only():
    from vcb.config import load_config
    from vcb.models import CAPREPAIR, Party
    from vcb.parsers.payments import ENFORCEMENT, classify, parse_registers
    from vcb.pipeline import build_case

    reg = Path(__file__).parent.parent / "examples" / "oplaty_kr.xls"
    pays = parse_registers(reg)
    classify(pays, load_config())
    assert pays and all(p.kind == ENFORCEMENT and not p.include for p in pays)  # «Погашение долга… Исполнительный лист»
    prof = {"claimant": Party(), "representative": Party(), "legal": {}}
    plain = build_case(KR_REAL, None, prof, load_config(), date(2026, 10, 30), None, category=CAPREPAIR)
    paid = build_case(KR_REAL, None, prof, load_config(), date(2026, 10, 30), None, (None, None), reg,
                      category=CAPREPAIR)
    assert paid.results[CAPREPAIR].debt == plain.results[CAPREPAIR].debt  # оплаты по ИП долг не уменьшают
    assert len(paid.results[CAPREPAIR].skipped) == len(pays)


# ---------------------------------------------------------------- ручной зачёт платежей
def make_register(path, rows):
    """Минимальный реестр банковских документов: rows = [(дата 'дд.мм.гггг', сумма, информация)]."""
    from openpyxl import Workbook

    ws = Workbook().active
    ws.append(["Дата", "Сумма", "Информация"])
    for r in rows:
        ws.append(list(r))
    ws.parent.save(path)


def test_manual_payment_overrides(tmp_path):
    from vcb.config import load_config
    from vcb.models import Party
    from vcb.parsers.payments import apply_overrides, classify, parse_register

    g, _ = two_packages(tmp_path)
    reg = tmp_path / "reg.xlsx"
    make_register(reg, [("10.07.2024", 1000, "оплата ЖКУ"), ("10.07.2024", 1000, "оплата ЖКУ"),
                        ("20.08.2024", 500, "по исполнительному производству 123-ИП")])
    pays = parse_register(reg)
    classify(pays, load_config())
    assert [p.include for p in pays] == [True, True, False]

    ov = tmp_path / "ov.csv"
    ov.write_text("дата;сумма;решение;комментарий;номер\n10.07.2024;1000;исключить;вручную;2\n"
                  "20.08.2024;*;зачесть;вручную\n", encoding="utf-8-sig")
    assert apply_overrides(pays, ov) == 2
    assert [p.include for p in pays] == [True, False, True]  # номер выбирает вторую из одинаковых

    from vcb.pipeline import build_case

    prof = {"claimant": Party(short_name="ООО «Тест»"), "representative": Party(), "legal": {}}
    as_of = date(2025, 4, 15)
    base = build_case(g, None, prof, load_config(), as_of, None, payments_path=reg)
    manual = build_case(g, None, prof, load_config(), as_of, None, payments_path=reg, overrides_path=ov)
    assert manual.debt == base.debt - Decimal("500") + Decimal("1000")
    assert any("ручные решения по платежам: 2" in n for n in manual.notes)


# ---------------------------------------------------------------- предзаполнение должника из файлов
def test_genitive_fio():
    from vcb.names import genitive_fio

    assert genitive_fio("Иванов Иван Иванович") == "Иванова Ивана Ивановича"
    assert genitive_fio("Петрова Мария Сергеевна") == "Петровой Марии Сергеевны"
    assert genitive_fio("Достоевский Фёдор Михайлович") == "Достоевского Фёдора Михайловича"
    assert genitive_fio("Организация") == ""


def test_autofill_from_files(tmp_path):
    from vcb.autofill import autofill

    g, _ = two_packages(tmp_path)
    reg = tmp_path / "reg.xlsx"
    make_register(reg, [("10.07.2024", 1000, "100000000001, Иванов Иван Иванович (л/с) / оплата")])
    values, notes = autofill(str(g), str(reg))
    assert (values["city"], values["street"], values["house"], values["flat"]) == ("Пермь", "Тестовая", "1", "2")
    assert values["fio"] == "Иванов Иван Иванович" and values["fio_gen"] == "Иванова Ивана Ивановича"
    assert any("реестра" in n for n in notes)


def _line(year, month, amount, portions):
    from vcb.calc.ledger import ChargeLine, Portion

    nxt = (year + (month == 12), month % 12 + 1)
    ln = ChargeLine(year, month, date(year, month, 28), date(nxt[0], nxt[1], 10), Decimal(amount))
    ln.portions = [Portion(Decimal(a), c) for a, c in portions]
    return ln


def test_moratorium_removes_only_overlapping_days():
    from vcb.calc.penalty import PenaltyRules, RateTable, calc_penalty

    mor = PenaltyRules(moratoria=[(date(2022, 4, 1), date(2022, 10, 1), None)])
    rates, as_of = RateTable(default=Decimal("9.5")), date(2023, 3, 1)
    ln = _line(2022, 1, 1000, [(1000, None)])  # 1/300 с 11.03.2022, 1/130 с 11.05.2022
    with_m, without = calc_penalty([ln], as_of, mor, rates)[0], calc_penalty([ln], as_of, PenaltyRules(), rates)[0]
    removed = sum(s.days for s in without.segments) - sum(s.days for s in with_m.segments)
    assert removed == (date(2022, 10, 1) - date(2022, 4, 1)).days + 1  # 184 дня окна
    assert with_m.segments[0].start == without.segments[0].start and with_m.penalty > 0


def test_payment_stops_penalty_only_for_paid_sum():
    from vcb.calc.penalty import PenaltyRules, RateTable, calc_penalty

    rates, as_of = RateTable(default=Decimal("9.5")), date(2025, 3, 1)
    paid_on = date(2024, 9, 20)
    ln = _line(2024, 6, 1000, [(400, paid_on), (600, None)])
    paid, rest = calc_penalty([ln], as_of, PenaltyRules(), rates)
    assert paid.base == Decimal(400) and paid.segments[-1].end == date(2024, 9, 19)
    assert rest.base == Decimal(600) and rest.segments[-1].end == as_of  # на остаток пени продолжаются


def test_moratorium_splits_period_dates():
    """Остаток за июнь 2020 (срок 10.07.2020): период 1/300 целиком в моратории 2020, период 1/130 разрезан."""
    from vcb.calc.penalty import PenaltyRules, RateTable, calc_penalty

    rules = PenaltyRules(moratoria=[(date(2020, 4, 1), date(2020, 12, 31), None),
                                    (date(2022, 4, 1), date(2022, 10, 1), None)])
    ln = _line(2020, 6, "215.53", [("215.53", None)])
    segs = calc_penalty([ln], date(2026, 11, 5), rules, RateTable(default=Decimal("9.5")))[0].segments
    assert [(s.period, s.days, s.divisor) for s in segs] == [
        ("01.01.21-31.03.22", 455, 130), ("02.10.22-05.11.26", 1496, 130)]


def test_penalty_schedule_full_sum_then_reduced():
    """Пени по остатку: вся сумма до дня оплаты, со дня оплаты - уменьшенная сумма; итог как по частям."""
    from vcb.calc.penalty import PenaltyRules, RateTable, calc_penalty, line_schedule

    rules, as_of = PenaltyRules(), date(2025, 3, 1)
    ln = _line(2024, 6, 1000, [(400, date(2024, 9, 20)), (600, None)])
    items = calc_penalty([ln], as_of, rules, RateTable(default=Decimal("9.5")))
    rows = line_schedule(ln, items, as_of, rules)
    shown = [(r.base, r.segment.period if r.segment else None, r.pay_date) for r in rows]
    assert shown == [(Decimal(1000), "11.08.24-19.09.24", None),
                     (Decimal(600), "20.09.24-10.10.24", date(2024, 9, 20)),
                     (Decimal(600), "11.10.24-01.03.25", None)]
    assert sum(r.segment.penalty for r in rows) == sum(i.penalty for i in items)


def test_claim_attachments_all_have_copies(tmp_path):
    import docx

    from vcb.config import load_config
    from vcb.models import CAPREPAIR, Party
    from vcb.pipeline import Package, build_case, write_case

    g, k = two_packages(tmp_path)
    prof = {"claimant": Party(short_name="ООО «Тест»"), "representative": Party(), "legal": {}}
    case = build_case(g, None, prof, load_config(), date(2025, 4, 15), None, extra_packages=[Package(CAPREPAIR, k)])
    doc = next(f for f in write_case(case, tmp_path / "out") if f.suffix == ".docx")
    paras = [p.text.replace("\xa0", " ") for p in docx.Document(doc).paragraphs]
    att = paras[paras.index("Приложение:") + 1:]
    att = [t for t in att if t[:1].isdigit()]
    assert att and all(t.endswith(" - 1 экз.") for t in att)
    assert any("Расчёт начислений и оплат по услугам" in t for t in att)
    assert not any("Шахматка" in t for t in att)


def test_file_names_have_month_address_and_surname(tmp_path):
    from vcb.config import load_config
    from vcb.models import Party
    from vcb.pipeline import build_case, file_suffix, write_case

    g, _ = two_packages(tmp_path)
    prof = {"claimant": Party(short_name="ООО «Тест»"), "representative": Party(), "legal": {}}
    # адрес должника другой - в имени файла всё равно адрес квартиры с долгом (из шапки шахматки)
    row = {"fio": "Иванов Иван Иванович", "street": "Мира", "house": "7", "flat": "11"}
    case = build_case(g, row, prof, load_config(), date(2026, 11, 5), None)
    assert file_suffix(case) == "ноябрь_2026_Тестовая_д1_кв2_Иванов"
    names = sorted(f.name for f in write_case(case, tmp_path / "out"))
    assert all(n.split(".")[0].endswith("_ноябрь_2026_Тестовая_д1_кв2_Иванов") for n in names), names
    flat = tmp_path / "flat.xlsx"
    make_chess(flat, "100000000002", "г.ПЕРМЬ, Ленина, д.55/1, кв. 98", {"Содержание": 100}, MONTHS_2024)
    case = build_case(flat, {"fio": "Петров Пётр Петрович"}, prof, load_config(), date(2026, 11, 5), None)
    assert file_suffix(case) == "ноябрь_2026_Ленина_д55-1_кв98_Петров"
    assert any(n.startswith(("Заявление_на_судебный_приказ_", "Исковое_заявление_")) for n in names)


def test_one_rate_for_whole_calculation():
    """Часть, погашенная, когда ключевая ставка была ниже потолка, считается по той же ставке, что и остаток."""
    from vcb.calc.penalty import PenaltyRules, calc_penalty, line_schedule
    from vcb.config import load_config, rules_from

    _, rates = rules_from(load_config())
    as_of = date(2026, 11, 5)
    ln = _line(2022, 11, "3977.43", [("2256.38", date(2023, 7, 20)), ("1721.05", None)])
    items = calc_penalty([ln], as_of, PenaltyRules(), rates)
    assert {s.rate for i in items for s in i.segments} == {Decimal("9.5")}
    rows = line_schedule(ln, items, as_of, PenaltyRules())
    assert [r.base for r in rows][:2] == [Decimal("3977.43"), Decimal("3977.43")]  # без раздвоения по ставкам
    items = calc_penalty([ln], as_of, PenaltyRules(rate_on="payment"), rates)
    assert rates.at(date(2023, 7, 20)) < Decimal("9.5") and items[0].segments[0].rate == rates.at(date(2023, 7, 20))


def test_flat_address_used_for_court_claim_and_files(tmp_path):
    """Адрес должника, введённый вручную, не подменяет адрес квартиры с долгом в заявлении и в выборе суда."""
    import docx

    from vcb.config import load_config
    from vcb.models import Party
    from vcb.pipeline import build_case, write_case

    g, _ = two_packages(tmp_path)
    prof = {"claimant": Party(short_name="ООО «Тест»"), "representative": Party(), "legal": {}}
    row = {"fio": "Иванов Иван Иванович", "street": "Мира", "house": "7", "flat": "11"}
    case = build_case(g, row, prof, load_config(), date(2026, 11, 5), None)
    assert (case.premises.street, case.premises.house, case.premises.flat) == ("Тестовая", "1", "2")
    doc = next(f for f in write_case(case, tmp_path / "o") if f.suffix == ".docx")
    text = " ".join(p.text for p in docx.Document(doc).paragraphs).replace("\xa0", " ")
    assert "жилого помещения, расположенного по адресу: г. Пермь, ул. Тестовая, д. 1, кв. 2" in text
    assert "Иванов Иван Иванович." in text and ".." not in text


def test_fixed_rate_from_config():
    from vcb.config import load_config, rules_from

    cfg = load_config()
    cfg["penalty"]["fixed_rate"] = "8"
    _, rates = rules_from(cfg)
    assert rates.at(date(2021, 1, 1)) == rates.at(date(2026, 11, 5)) == Decimal("8")
