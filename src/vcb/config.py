from __future__ import annotations

import copy
from datetime import date
from decimal import Decimal
from pathlib import Path

import yaml

from .calc.penalty import PenaltyRules, RateTable
from .money import D

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data"


def _merge(a: dict, b: dict) -> dict:
    out = copy.deepcopy(a)
    for k, v in (b or {}).items():
        out[k] = _merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


def load_config(path: str | Path | None = None) -> dict:
    cfg = yaml.safe_load((DATA / "config.yaml").read_text(encoding="utf-8"))
    if path:
        cfg = _merge(cfg, yaml.safe_load(Path(path).read_text(encoding="utf-8")))
    return cfg


def _d(v) -> date:
    return v if isinstance(v, date) else date.fromisoformat(str(v))


def rules_from(cfg: dict, category: str = "services") -> tuple[PenaltyRules, RateTable]:
    """Правила пени. Для взносов на капремонт (`caprepair`) поверх общих накладываются penalty.caprepair."""
    p = dict(cfg["penalty"])
    if category == "caprepair":
        p.update(cfg["penalty"].get("caprepair") or {})
    rules = PenaltyRules(mode=p["mode"], divisor1=p["divisor1"], divisor2=p["divisor2"], tier1_days=p["tier1_days"],
                         due_day=p["due_day"], include_payment_day=p["include_payment_day"],
                         moratoria=[(_d(a), _d(b)) for a, b in p.get("moratoria") or []])
    rates = RateTable(default=D(p["key_rate"]))
    cap = p.get("rate_cap")
    if cap:
        rates.cap, rates.cap_from, rates.cap_to = D(cap["value"]), _d(cap.get("from")) if cap.get("from") else None, \
            _d(cap.get("to")) if cap.get("to") else None
    if p.get("key_rates_file"):
        f = Path(p["key_rates_file"])
        for line in (f if f.is_absolute() else DATA.parent / f).read_text(encoding="utf-8").splitlines():
            if line.strip() and not line.lower().startswith("дата"):
                d, r = line.replace(",", ".").split(";")[:2]
                rates.history.append((date.fromisoformat(d.strip()), Decimal(r.strip())))
    return rules, rates
