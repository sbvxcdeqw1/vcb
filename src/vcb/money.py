from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal

from num2words import num2words

CENT = Decimal("0.01")


def D(x) -> Decimal:
    if isinstance(x, Decimal):
        return x
    return Decimal(str(x).replace(" ", "").replace(",", "."))


def q(x: Decimal) -> Decimal:
    return x.quantize(CENT, rounding=ROUND_HALF_UP)


def fmt(x: Decimal) -> str:
    """64 613,90 - как в образцах."""
    return f"{q(x):,.2f}".replace(",", " ").replace(".", ",")


def _plural(n: int, forms: tuple[str, str, str]) -> str:
    n = abs(n) % 100
    if 10 < n < 20:
        return forms[2]
    n %= 10
    return forms[0] if n == 1 else forms[1] if 2 <= n <= 4 else forms[2]


def in_words(x: Decimal) -> str:
    """шестьдесят четыре тысячи ... рублей 90 копеек"""
    x = q(x)
    rub, kop = int(x), int((x - int(x)) * 100)
    words = num2words(rub, lang="ru")
    return f"{words} {_plural(rub, ('рубль', 'рубля', 'рублей'))} {kop:02d} {_plural(kop, ('копейка', 'копейки', 'копеек'))}"
