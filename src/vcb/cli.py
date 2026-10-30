from __future__ import annotations

import argparse
import csv
import sys
from datetime import date
from pathlib import Path

from . import __version__
from .config import DATA, load_config
from .court.registry import CourtRegistry
from .debtors import read_registry
from .money import fmt
from .models import CAPREPAIR, SERVICES
from .pipeline import NAMES, Package, build_case, pair_units, safe, write_case
from .profiles import choose_profile, load_profile


def check(argv):
    """vcb check КАРТОЧКА ШАХМАТКА [--merge ИТОГ.xlsx]: сверка выгрузок и шахматка, приведённая к карточке."""
    from .docs.checkerboard import write_checkerboard
    from .docs.context import Case
    from .models import Debtor, Party
    from .parsers.onec import parse_any
    from .reconcile import compare, format_report, reconciled

    ap = argparse.ArgumentParser(prog="vcb check")
    ap.add_argument("card"), ap.add_argument("chess")
    ap.add_argument("--merge", help="сохранить шахматку, дополненную строкой «не распределено по услугам»")
    a = ap.parse_args(argv)
    card, chess = parse_any(a.card), parse_any(a.chess)
    for w in card.warnings + chess.warnings:
        print("! ", w)
    print(format_report(compare(card, chess)))
    if a.merge:
        acc = reconciled(card, chess)
        case = Case(acc, Debtor(), Party(), Party(), date.today(), date.today(), max(e.date for e in acc.entries))
        write_checkerboard(a.merge, case)
        print(f"\nШахматка со строкой «не распределено» сохранена: {a.merge}")
    return 0


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] == "check":
        return check(argv[1:])
    return run(argv)


def run(argv):
    ap = argparse.ArgumentParser(prog="vcb", description="Заявления в суд и карточки расчёта по ЖКУ из выгрузок 1С")
    ap.add_argument("--version", action="version", version=f"vcb {__version__}")
    ap.add_argument("input", nargs="?", help="шахматка ЖКУ (xls/xlsx) или папка с шахматками (пакетный режим)")
    ap.add_argument("--cap-input", help="шахматка капитального ремонта (файл или папка); можно вместе с ЖКУ или одна")
    ap.add_argument("--cap-payments", help="реестр оплат по капитальному ремонту (файл или папка)")
    ap.add_argument("--cap-payment-overrides", help="csv с ручными решениями по платежам реестра капремонта")
    ap.add_argument("--cap-exclude-services", help="услуги шахматки капремонта, снятые с расчёта, через «;»")
    ap.add_argument("-o", "--out", default="out", help="папка результата")
    ap.add_argument("-r", "--registry", help="реестр должников (csv/xlsx): набор полей по каждому должнику")
    ap.add_argument("--profile", help="профиль взыскателя и представителя (yaml); без него - спросим")
    ap.add_argument("--region", default="perm_krai", help="регион для определения подсудности (файл data/courts/<регион>.yaml)")
    ap.add_argument("--courts", help="свой справочник судов (yaml)")
    ap.add_argument("--config", help="свои настройки расчёта (yaml)")
    ap.add_argument("--as-of", help="дата расчёта ГГГГ-ММ-ДД (по умолчанию сегодня)")
    ap.add_argument("--period-from"), ap.add_argument("--period-to")
    ap.add_argument("--payments", help="реестр оплат (файл или папка): оплаты разносятся по назначению платежа; общий реестр фильтруется по лицевому счёту")
    ap.add_argument("--payment-overrides", help="csv (;): дата;сумма;зачесть|исключить;комментарий")
    ap.add_argument("--exclude-services", help="услуги шахматки, снятые с расчёта, через «;» (например: Домофон; видеонаблюдение)")
    ap.add_argument("--no-claim", action="store_true", help="не готовить заявление/иск")
    ap.add_argument("--allocation-report", action="store_true", help="дополнительно сохранить отчёт о разноске оплат")
    ap.add_argument("--doc", choices=["auto", "order", "claim"], help="судебный приказ / иск")
    ap.add_argument("--no-interactive", action="store_true")
    a = ap.parse_args(argv)

    cfg = load_config(a.config)
    if a.doc:
        cfg["court"]["doc"] = a.doc
    as_of = date.fromisoformat(a.as_of) if a.as_of else date.today()
    period = (date.fromisoformat(a.period_from) if a.period_from else None,
              date.fromisoformat(a.period_to) if a.period_to else None)
    profile = load_profile(a.profile) if a.profile else choose_profile(interactive=not a.no_interactive)
    try:
        registry = CourtRegistry.load(a.region, Path(a.courts) if a.courts else None)
    except FileNotFoundError:
        registry = None

    if not a.input and not a.cap_input:
        ap.error("укажите шахматку ЖКУ и/или --cap-input (шахматку капитального ремонта)")
    cap_ov = Path(a.cap_payment_overrides) if a.cap_payment_overrides else None
    units = pair_units(Path(a.input) if a.input else None, Path(a.cap_input) if a.cap_input else None)
    rows = {}
    if a.registry:
        for r in read_registry(a.registry):
            if r.get("file"):
                rows[r["file"].strip().lower()] = r
    out = Path(a.out)
    summary, failed = [], 0
    for gku, kr in units:
        title = " + ".join(x.name for x in (gku, kr) if x)
        try:
            row = (rows.get(gku.name.lower()) if gku else None) or (rows.get(kr.name.lower()) if kr else None)
            cap_ov = Path(a.cap_payment_overrides) if a.cap_payment_overrides else None
            extra = [Package(CAPREPAIR, kr, Path(a.cap_payments) if a.cap_payments else None,
                             a.cap_exclude_services, cap_ov)] if gku and kr else None
            if gku:
                case = build_case(gku, row, profile, cfg, as_of, registry, period, a.payments, a.payment_overrides,
                                  a.exclude_services, SERVICES, extra)
            else:
                case = build_case(kr, row, profile, cfg, as_of, registry, period, a.cap_payments,
                                  cap_ov or a.payment_overrides, a.cap_exclude_services, CAPREPAIR)
            folder = out / safe(f"{case.account.number}_{case.debtor.fio}")
            write_case(case, folder, claim=not a.no_claim, allocation_report=a.allocation_report)
            kinds = " + ".join(NAMES[c] for c in case.results)
            summary.append([title, case.debtor.fio, case.account.number, kinds, fmt(case.debt), fmt(case.penalty),
                            fmt(case.price), fmt(case.fee), case.doc_kind, case.court_name, "; ".join(case.notes),
                            str(folder)])
            print(f"OK   {title} [{kinds}]: долг {fmt(case.debt)}, пени {fmt(case.penalty)} -> {folder}")
            for n in case.notes:
                print(f"     ! {n}")
        except Exception as e:  # ошибка по одному должнику не останавливает пакет
            failed += 1
            summary.append([title, "", "", "", "", "", "", "", "", "", f"ОШИБКА: {e}", ""])
            print(f"FAIL {title}: {e}", file=sys.stderr)
    out.mkdir(parents=True, exist_ok=True)
    with open(out / "сводка.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Файлы", "Должник", "ЛС", "Пакеты", "Долг", "Пени", "Цена иска", "Госпошлина", "Документ", "Суд",
                    "Замечания", "Папка"])
        w.writerows(summary)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
