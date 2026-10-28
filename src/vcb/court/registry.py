"""Справочник судов и подсудность по адресу квартиры.

Файл data/courts/<регион>.yaml:

courts:
  - id: perm-kirovsky-rs
    level: district | magistrate
    name: Кировский районный суд г. Перми
    address: 614113, г. Пермь, ул. Кировоградская, д. 37
    territory:                       # адреса, подсудные суду / участку
      - city: Пермь
        streets:
          - {street: Кировоградская, houses: "1-99"}      # "all", "1-99", "2-40 чёт", "1-39 нечёт", "5,7,9"
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from ..config import DATA
from ..models import Debtor


@dataclass
class Court:
    id: str
    level: str
    name: str
    address: str = ""
    territory: list = field(default_factory=list)


def _norm(s: str) -> str:
    s = (s or "").lower().replace("ё", "е")
    s = re.sub(r"\b(ул|улица|пр|просп|проспект|пер|переулок|б-р|бульвар|ш|шоссе|г|город|пл|площадь)\b\.?", "", s)
    return re.sub(r"[^\w]+", " ", s).strip()


def _house_no(h: str) -> tuple[int, str]:
    m = re.match(r"\s*(\d+)\s*([а-яa-z]?)", (h or "").lower())
    return (int(m.group(1)), m.group(2)) if m else (0, "")


def house_matches(spec, house: str) -> bool:
    spec = str(spec or "all").lower().replace("ё", "е").strip()
    n, _ = _house_no(house)
    if spec in ("all", "все", ""):
        return True
    parity = "чет" if "чет" in spec and "нечет" not in spec else "нечет" if "нечет" in spec else None
    spec = re.sub(r"[а-я]+", "", spec).strip()
    if parity == "чет" and n % 2:
        return False
    if parity == "нечет" and not n % 2:
        return False
    for part in re.split(r"[;,]", spec):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = (int(x) for x in part.split("-", 1))
            if a <= n <= b:
                return True
        elif part.isdigit() and int(part) == n:
            return True
    return False


class CourtRegistry:
    def __init__(self, courts: list[Court], default_city: str = ""):
        self.courts = courts
        self.default_city = default_city

    @classmethod
    def load(cls, region: str, path: Path | None = None) -> "CourtRegistry":
        path = path or DATA / "courts" / f"{region}.yaml"
        raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        return cls([Court(**c) for c in raw.get("courts", [])], raw.get("default_city", ""))

    def find(self, debtor: Debtor, level: str) -> Court | None:
        city, street = _norm(debtor.city), _norm(debtor.street)
        for c in self.courts:
            if c.level != level:
                continue
            for t in c.territory:
                if _norm(t.get("city", "")) != city:
                    continue
                streets = t.get("streets")
                if not streets:  # город целиком
                    return c
                for s in streets:
                    if _norm(s["street"]) == street and house_matches(s.get("houses"), debtor.house):
                        return c
        return None
