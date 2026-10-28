from __future__ import annotations

from decimal import Decimal

from ..money import D, q


def state_fee(price: Decimal, cfg: dict, *, level: str, order: bool) -> Decimal:
    f = cfg["fee"]
    fee = None
    for upto, fixed, rate, base in f["tiers"]:
        if upto is None or price <= D(upto):
            fee = D(fixed) + (price - D(base)) * D(rate)
            break
    fee = max(fee, D(f["min_district"] if level == "district" else f["min_magistrate"]))
    fee = min(fee, D(f["max"]))
    if order:
        fee *= D(f["order_share"])
    return q(fee)
