"""Профили взыскателя и представителя; интерактивный ввод, если профиль не задан."""
from __future__ import annotations

import sys
from pathlib import Path

import yaml

from .config import ROOT
from .models import Party

PARTY_FIELDS = [("name", "Полное наименование (Общество с ограниченной ответственностью «…»)"),
                ("short_name", "Краткое наименование (ООО «…»)"), ("ogrn", "ОГРН"), ("inn", "ИНН"),
                ("address", "Адрес"), ("phone", "Телефон"), ("sign_name", "Подпись (И.О.Фамилия)")]
REP_FIELDS = [("name", "ФИО представителя"), ("address", "Адрес"), ("phone", "Телефон"),
              ("sign_name", "Подпись (И.О.Фамилия)")]


def load_profile(path: str | Path) -> dict:
    raw = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    return {
        "claimant": Party(**{k: v for k, v in (raw.get("claimant") or {}).items() if k in Party.__dataclass_fields__}),
        "representative": Party(**{k: v for k, v in (raw.get("representative") or {}).items()
                                   if k in Party.__dataclass_fields__}),
        "legal": raw.get("legal") or {},
    }


def _ask(fields) -> Party:
    p = Party()
    for key, title in fields:
        setattr(p, key, input(f"  {title}: ").strip())
    return p


def choose_profile(directory: Path | None = None, interactive: bool = True) -> dict:
    """Выбор сохранённого профиля или ввод вручную. Пустой ввод - поля останутся пустыми в документах."""
    empty = {"claimant": Party(), "representative": Party(), "legal": {}}
    if not interactive or not sys.stdin.isatty():
        return empty
    directory = directory or ROOT / "profiles"
    found = sorted(directory.glob("*.yaml")) if directory.exists() else []
    print("Взыскатель и представитель (необязательно).")
    for i, f in enumerate(found, 1):
        print(f"  {i}. профиль {f.stem}")
    print("  n. ввести вручную     Enter. пропустить (поля останутся пустыми)")
    ans = input("Выбор: ").strip().lower()
    if ans.isdigit() and 1 <= int(ans) <= len(found):
        return load_profile(found[int(ans) - 1])
    if ans == "n":
        print("Взыскатель:")
        claimant = _ask(PARTY_FIELDS)
        print("Представитель (Enter - пропустить):")
        rep = _ask(REP_FIELDS)
        save = input("Сохранить профиль под именем (Enter - не сохранять): ").strip()
        if save:
            directory.mkdir(exist_ok=True)
            from dataclasses import asdict
            (directory / f"{save}.yaml").write_text(yaml.safe_dump(
                {"claimant": {k: v for k, v in asdict(claimant).items() if v and k != "extra"},
                 "representative": {k: v for k, v in asdict(rep).items() if v and k != "extra"}},
                allow_unicode=True, sort_keys=False), encoding="utf-8")
        return {"claimant": claimant, "representative": rep, "legal": {}}
    return empty
