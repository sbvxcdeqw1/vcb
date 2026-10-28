from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

SERVICES = "services"  # содержание, текущий ремонт, коммунальные услуги
CAPREPAIR = "caprepair"  # взносы на капитальный ремонт
CATEGORY_TITLES = {
    SERVICES: "содержание жилья, текущий ремонт, коммунальные услуги",
    CAPREPAIR: "взносы на капитальный ремонт",
}

CHARGE, PAYMENT, PENALTY, OPENING = "charge", "payment", "penalty", "opening"


@dataclass
class Entry:
    """Строка выгрузки 1С. `purpose` - назначение (операция) как в выгрузке."""

    date: date
    kind: str  # CHARGE | PAYMENT | PENALTY
    category: str  # SERVICES | CAPREPAIR
    purpose: str
    amount: Decimal  # начисление/оплата; перерасчёт может быть отрицательным
    registrar: str = ""


@dataclass
class Account:
    """Лицевой счёт: шапка выгрузки + проводки."""

    number: str = ""
    owner: str = ""
    address: str = ""
    entries: list[Entry] = field(default_factory=list)
    source: str = ""
    warnings: list[str] = field(default_factory=list)
    saldo_start: Decimal | None = None  # «Сальдо на …» из карточки (услуги + пени вместе)
    saldo_start_date: date | None = None
    saldo_end: Decimal | None = None


@dataclass
class Party:
    """Взыскатель или представитель (все поля необязательны)."""

    name: str = ""
    short_name: str = ""
    ogrn: str = ""
    inn: str = ""
    address: str = ""
    phone: str = ""
    sign_name: str = ""  # «Д.Г.Фалеев»
    extra: dict = field(default_factory=dict)


@dataclass
class Debtor:
    """Данные должника/квартиры (набор полей)."""

    fio: str = ""
    city: str = ""
    street: str = ""
    house: str = ""
    flat: str = ""
    area: str = ""
    account: str = ""
    birth_date: str = ""
    birth_place: str = ""
    share: str = ""
    extra: dict = field(default_factory=dict)

    @property
    def street_full(self) -> str:
        """В выгрузках 1С тип улицы часто опущен («Кировоградская») - по умолчанию «ул.»."""
        st = self.street.strip()
        if st and not re.search(r"\b(ул|улица|пр|пр-кт|просп|проспект|пер|переулок|проезд|б-р|бульвар|ш|шоссе|пл|"
                                r"площадь|наб|набережная|тракт|мкр|микрорайон)\b\.?", st, re.I):
            return f"ул. {st}"
        return st

    @property
    def address(self) -> str:
        parts = [f"г. {self.city}" if self.city else "", self.street_full, f"д. {self.house}" if self.house else "",
                 f"кв. {self.flat}" if self.flat else ""]
        return ", ".join(p for p in parts if p)
