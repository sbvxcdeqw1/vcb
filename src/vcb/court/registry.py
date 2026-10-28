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
    level: str  # district | magistrate
    name: str
    address: str = ""
    territory: list = field(default_factory=list)  # [{city, streets: [{street, houses}]}] - разобранные улицы
    parent: str | None = None  # id районного суда для участка мирового судьи
    judicial_district: str = ""
    judge: str = ""
    territory_text: str = ""  # описание границ как в источнике (для поиска, если улицы не разобраны)


def _norm(s: str) -> str:
    s = (s or "").lower().replace("ё", "е")
    s = re.sub(r"\b(ул|улица|пр|просп|проспект|пер|переулок|б-р|бульвар|ш|шоссе|г|город|пл|площадь)\b\.?", "", s)
    return re.sub(r"[^\w]+", " ", s).strip()


def _split_top(spec: str) -> list[str]:
    """Разбиение по запятым вне скобок."""
    out, depth, cur = [], 0, ""
    for ch in spec:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(depth - 1, 0)
        if ch in ",;" and depth == 0:
            out.append(cur.strip())
            cur = ""
        else:
            cur += ch
    out.append(cur.strip())
    return [x for x in out if x]


def _norm_house(h: str) -> str:
    h = (h or "").lower().replace("ё", "е").replace("\xa0", " ")
    h = re.sub(r"\b(д|дом|стр|строение)\.?\s*", "", h)
    h = re.sub(r"\s*(корп\.?|к)\s*(\d+)", r"/\2", h) if "/" not in h else h
    return re.sub(r"\s+", "", h)


def house_matches(spec, house: str) -> bool:
    """Диапазон «12-40» покрывает только дома без буквы и дроби; «35а», «27/1» - точное совпадение.
    «(четные)» / «(нечетные)» относятся к диапазону перед скобкой."""
    spec = str(spec or "all").lower().replace("ё", "е").strip()
    if spec in ("all", "все", ""):
        return True
    h = _norm_house(house)
    m = re.fullmatch(r"(\d+)", h)
    n = int(m.group(1)) if m else None
    for item in _split_top(spec):
        parity = None
        if "нечет" in item:
            parity = "odd"
        elif "чет" in item:
            parity = "even"
        core = re.sub(r"\(.*?\)|\b(не)?чет\w*", "", item).strip()
        if re.fullmatch(r"\d+\s*-\s*\d+", core):
            if n is None:
                continue
            a, b = (int(x) for x in re.split(r"\s*-\s*", core))
            if a <= n <= b and (parity is None or (n % 2 == 0) == (parity == "even")):
                return True
        elif _norm_house(core) == h:
            if n is None or parity is None or (n % 2 == 0) == (parity == "even"):
                return True
    return False


class CourtRegistry:
    def __init__(self, courts: list[Court], default_city: str = ""):
        self.courts = courts
        self.default_city = default_city
        self.notes: list[str] = []

    @classmethod
    def load(cls, region: str, path: Path | None = None) -> "CourtRegistry":
        path = path or DATA / "courts" / f"{region}.yaml"
        raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        return cls([Court(**c) for c in raw.get("courts", [])], raw.get("default_city", ""))

    def by_id(self, cid: str | None) -> Court | None:
        return next((c for c in self.courts if c.id == cid), None)

    def find(self, debtor: Debtor, level: str) -> Court | None:
        """Суд по адресу квартиры. Для районного суда сначала определяется участок мирового судьи.
        Причины неоднозначности складываются в `self.notes`."""
        self.notes = []
        ms = self._find_magistrate(debtor)
        if level == "magistrate":
            return ms
        if ms is not None and ms.parent:
            return self.by_id(ms.parent)
        cands = [c for c in self.courts if c.level == "district" and _norm(debtor.city) and
                 re.search(rf"\b{re.escape(_norm(debtor.city))}", _norm(c.address))]
        if len(cands) == 1:
            return cands[0]
        if cands:
            self.notes.append("несколько районных судов в населённом пункте: " + "; ".join(c.name for c in cands))
        return None

    def _find_magistrate(self, debtor: Debtor) -> Court | None:
        city, street = _norm(debtor.city), _norm(debtor.street)
        mags = [c for c in self.courts if c.level == "magistrate"]
        # 1. разобранные списки улиц и домов
        for c in mags:
            for t in c.territory:
                if _norm(t.get("city", "")) != city:
                    continue
                streets = t.get("streets")
                if not streets:
                    return c
                for sdef in streets:
                    if _norm(sdef["street"]) == street and house_matches(sdef.get("houses"), debtor.house):
                        return c
        # 2. описание границ текстом: населённый пункт (+ улица) должны встречаться в тексте
        if not city:
            return None
        rx_city = re.compile(rf"(?<![а-яё]){re.escape(city[:max(len(city) - 1, 4)])}", re.I)
        cands = [c for c in mags if rx_city.search(c.territory_text.lower().replace("ё", "е"))]
        if street and len(cands) > 1:
            rx_st = re.compile(rf"(?<![а-яё]){re.escape(street.split()[0][:max(len(street.split()[0]) - 2, 4)])}", re.I)
            narrowed = [c for c in cands if rx_st.search(c.territory_text.lower().replace("ё", "е"))]
            cands = narrowed or cands
        if len(cands) == 1:
            self.notes.append(f"участок определён по тексту описания границ — проверьте: {cands[0].name}")
            return cands[0]
        if cands:
            self.notes.append("не удалось однозначно определить участок мирового судьи; кандидаты: "
                              + "; ".join(c.name for c in cands[:6]))
        return None
