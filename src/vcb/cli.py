from __future__ import annotations

import argparse
import csv
import sys
from datetime import date
from pathlib import Path

from .config import DATA, load_config
from .court.registry import CourtRegistry
from .debtors import read_registry
from .money import fmt
from .pipeline import build_case, safe, write_case
from .profiles import choose_profile, load_profile


def main(argv=None):
    ap = argparse.ArgumentParser(prog="vcb", description="Заявления в суд и карточки расчёта по ЖКУ из выгрузок 1С")
    ap.add_argument("input", help="выгрузка 1С (xls/xlsx) или папка с выгрузками (пакетный режим)")
    ap.add_argument("-o", "--out", default="out", help="папка результата")
    ap.add_argument("-r", "--registry", help="реестр должников (csv/xlsx): набор полей по каждому должнику")
    ap.add_argument("--profile", help="профиль взыскателя и представителя (yaml); без него - спросим")
    ap.add_argument("--region", default="perm_krai", help="регион для определения подсудности (файл data/courts/<регион>.yaml)")
    ap.add_argument("--courts", help="свой справочник судов (yaml)")
    ap.add_argument("--config", help="свои настройки расчёта (yaml)")
    ap.add_argument("--as-of", help="дата расчёта ГГГГ-ММ-ДД (по умолчанию сегодня)")
    ap.add_argument("--period-from"), ap.add_argument("--period-to")
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

    src = Path(a.input)
    cards = sorted(p for p in src.iterdir() if p.suffix.lower() in (".xls", ".xlsx")
                   and not p.name.startswith(("~", "реестр", "registry"))) if src.is_dir() else [src]
    rows = {}
    if a.registry:
        for r in read_registry(a.registry):
            if r.get("file"):
                rows[r["file"].strip().lower()] = r
    out = Path(a.out)
    summary, failed = [], 0
    for card in cards:
        try:
            case = build_case(card, rows.get(card.name.lower()), profile, cfg, as_of, registry, period)
            folder = out / safe(f"{case.account.number}_{case.debtor.fio}")
            write_case(case, folder)
            summary.append([card.name, case.debtor.fio, case.account.number, fmt(case.debt), fmt(case.penalty),
                            fmt(case.price), fmt(case.fee), case.doc_kind, case.court_name, "; ".join(case.notes),
                            str(folder)])
            print(f"OK   {card.name}: долг {fmt(case.debt)}, пени {fmt(case.penalty)} -> {folder}")
            for n in case.notes:
                print(f"     ! {n}")
        except Exception as e:  # ошибка по одному должнику не останавливает пакет
            failed += 1
            summary.append([card.name, "", "", "", "", "", "", "", "", f"ОШИБКА: {e}", ""])
            print(f"FAIL {card.name}: {e}", file=sys.stderr)
    out.mkdir(parents=True, exist_ok=True)
    with open(out / "сводка.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Файл", "Должник", "ЛС", "Долг", "Пени", "Цена иска", "Госпошлина", "Документ", "Суд",
                    "Замечания", "Папка"])
        w.writerows(summary)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
