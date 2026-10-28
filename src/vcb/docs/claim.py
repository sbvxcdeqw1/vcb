"""Исковое заявление / заявление о выдаче судебного приказа (docx). Текст - по образцам заказчика."""
from __future__ import annotations

from pathlib import Path

import yaml
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from ..config import ROOT
from ..models import CAPREPAIR, CATEGORY_TITLES, SERVICES
from ..money import fmt, in_words
from .context import Case

BLANK = "________________"
STATIC = yaml.safe_load((ROOT / "templates" / "claim_static.yaml").read_text(encoding="utf-8"))

SUBJECT = {
    SERVICES: "по оплате за содержание, текущий ремонт общего имущества многоквартирного дома, предоставление "
              "коммунальных услуг",
    CAPREPAIR: "по оплате взносов на капитальный ремонт общего имущества в многоквартирном доме",
}
ATT_SERVICES = {SERVICES: "по услугам", CAPREPAIR: "по взносам на капремонт"}


def _doc() -> Document:
    d = Document()
    s = d.sections[0]
    s.left_margin, s.right_margin, s.top_margin, s.bottom_margin = Cm(3), Cm(1.5), Cm(1.5), Cm(1.5)
    st = d.styles["Normal"]
    st.font.name, st.font.size = "Times New Roman", Pt(12)
    st.paragraph_format.space_after = Pt(0)
    return d


def _p(d, text="", *, bold=False, align=None, indent=True, after=6, right_block=False):
    p = d.add_paragraph()
    if right_block:
        p.paragraph_format.left_indent = Cm(8.5)
    elif indent:
        p.paragraph_format.first_line_indent = Cm(1.25)
    p.paragraph_format.space_after = Pt(after)
    p.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    r.bold = bold
    return p


def _val(x, blank=BLANK):
    return x if x else blank


def _period(case):
    return f"с {case.period_from:%d.%m.%Y} года по {case.period_to:%d.%m.%Y} года"


def _money(x):
    return f"{fmt(x)} руб. ({in_words(x)})"


def write_claim(path, case: Case):
    order = case.doc_kind == "order"
    c, rep, dbt = case.claimant, case.representative, case.debtor
    cname = _val(c.short_name or c.name)
    fio_gen = dbt.extra.get("fio_gen") or dbt.fio
    asof = f"{case.as_of:%d.%m.%Y}г."
    d = _doc()
    # --- шапка
    _p(d, _val(case.court_name, "Суд (определите по адресу вручную)"), bold=True, right_block=True, after=0)
    _p(d, case.court_address, right_block=True, after=6)
    _p(d, f"Взыскатель: {cname}", bold=True, right_block=True, after=0)
    if c.ogrn or c.inn:
        _p(d, f"ОГРН {_val(c.ogrn, '________')}, ИНН {_val(c.inn, '________')}", right_block=True, after=0)
    _p(d, _val(c.address, "адрес взыскателя"), right_block=True, after=6)
    if rep.name:
        _p(d, "Представитель взыскателя:", bold=True, right_block=True, after=0)
        _p(d, rep.name, right_block=True, after=0)
        _p(d, rep.address, right_block=True, after=0)
        if rep.phone:
            _p(d, f"тел: {rep.phone}", right_block=True, after=6)
    _p(d, "Должник:" if order else "Ответчик:", bold=True, right_block=True, after=0)
    _p(d, _val(dbt.fio), bold=True, right_block=True, after=0)
    _p(d, _val(dbt.address), right_block=True, after=0)
    if order:
        _p(d, f"Дата рождения: {_val(dbt.birth_date, '________')}, место рождения: "
              f"{_val(dbt.birth_place, '________')}", right_block=True, after=0)
    _p(d, f"{'Размер требований' if order else 'Цена иска'}: {fmt(case.price)} руб.", right_block=True, after=0)
    _p(d, f"Госпошлина: {fmt(case.fee)} руб.", right_block=True, after=12)
    title = "ЗАЯВЛЕНИЕ о вынесении судебного приказа" if order else "ИСКОВОЕ ЗАЯВЛЕНИЕ"
    _p(d, title, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, after=0)
    _p(d, "о взыскании задолженности по оплате за жилое помещение и коммунальные услуги",
       align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, after=12)
    # --- обстоятельства
    full = c.name or cname
    house = f"{dbt.house}" if dbt.house else BLANK
    _p(d, f"{full} осуществляет управление многоквартирным домом № {house}, расположенным по адресу: "
          f"{_val(dbt.city and 'г. ' + dbt.city)}{', ' + dbt.street_full if dbt.street else ''}.")
    _p(d, f"Собственником жилого помещения, расположенного по адресу: {_val(dbt.address)} является: "
          f"{_val(dbt.fio)}.")
    for t in STATIC["housing_law"]:
        _p(d, t)
    _p(d, f"{'Должник' if order else 'Ответчик'} {_val(dbt.fio)} ненадлежащим образом исполнял(а) свои обязанности "
          f"по оплате за жилое помещение и коммунальные услуги: не в полном объеме осуществлял(а) оплату "
          f"предоставленных жилищно-коммунальных услуг, допустил(а) образование задолженности.")
    for cat, res in case.results.items():
        _p(d, f"За период {_period(case)} размер задолженности {SUBJECT[cat]} составил по состоянию на {asof} "
              f"{_money(res.debt)} (Выписка из лицевого счета прилагается).")
    for t in STATIC["penalty_law"]:
        _p(d, t)
    key = next(iter(case.results.values())).rate if case.results else None
    if key is not None:
        _p(d, f"Размер ключевой ставки Банка России для расчета пени на момент подачи заявления составляет - "
              f"{fmt(key)}% годовых;")
    for cat, res in case.results.items():
        _p(d, f"Размер пени за несвоевременную оплату {SUBJECT[cat].replace('по оплате ', '')} за период "
              f"{_period(case)} составил по состоянию на {asof} {_money(res.penalty)} (Расчет пени прилагается).")
    _p(d, f"{'Должник' if order else 'Ответчик'}, ненадлежащим образом выполняя свои обязанности по оплате "
          f"жилищно-коммунальных услуг, нарушает требования жилищного законодательства, а также права и законные "
          f"интересы {'взыскателя' if order else 'истца'}.")
    for t in STATIC["costs_law"][:1]:
        _p(d, t)
    if case.legal_costs:
        _p(d, f"Для подготовки заявления о взыскании задолженности по оплате за жилое помещение и коммунальные "
              f"услуги с {fio_gen} {cname} заключило с {_val(case.legal_executor)} договор "
              f"{_val(case.legal_contract)} на оказание юридических услуг, стоимость услуг по договору составила "
              f"{fmt(case.legal_costs)} руб.")
    for t in STATIC["costs_law"][1:]:
        _p(d, t)
    # --- требования
    if order:
        _p(d, "Требование является бесспорным: подтверждается документами о начислении платы и расчётом "
              "задолженности. Спора о праве нет. Руководствуясь ст. 30, 153, 154, 155, 158, 161 Жилищного кодекса "
              "РФ, ст.ст. 121, 122, 123, 124 Гражданского процессуального кодекса РФ,")
        _p(d, "ПРОШУ:", bold=True, indent=False)
        _p(d, f"1. Выдать судебный приказ о взыскании с должника {fio_gen} в пользу {full} задолженности по "
              f"оплате за жилое помещение и коммунальные услуги за период {_period(case)} по состоянию на {asof} в "
              f"размере {fmt(case.debt)} руб., пени в размере {fmt(case.penalty)} руб., всего {fmt(case.price)} руб.")
    else:
        _p(d, "На основании изложенного, руководствуясь ст. 30, 153, 154, 155, 158, 161 Жилищного Кодекса РФ, "
              "ст.ст. 88, 98, 131, 132 Гражданского процессуального кодекса РФ,")
        _p(d, "ПРОШУ:", bold=True, indent=False)
        _p(d, f"1. Взыскать с должника {fio_gen} в пользу {full} задолженность "
              f"за период {_period(case)} включительно по состоянию на {asof} в размере {fmt(case.debt)} руб., пени "
              f"в размере {fmt(case.penalty)} руб. с последующим доначислением по день фактической оплаты "
              f"задолженности, всего на общую сумму {fmt(case.price)} руб.")
    _p(d, f"2. Взыскать с должника {fio_gen} в пользу {full} расходы по оплате государственной пошлины в размере "
          f"{fmt(case.fee)} рублей.")
    if case.legal_costs:
        _p(d, f"3. Взыскать с должника {fio_gen} в пользу {full} судебные расходы в сумме "
              f"{fmt(case.legal_costs)} рублей.")
    # --- приложения
    _p(d, "Приложение:", bold=True, indent=False, after=0)
    n = 0
    items = []
    if not order:
        items.append("копия чека об отправке искового заявления ответчику")
    for cat in case.results:
        items.append(f"Карточка расчетов по лицевому счету {ATT_SERVICES[cat]} - 1 экз.")
        items.append(f"Расчет пени {ATT_SERVICES[cat]} - 1 экз.")
    items += ["Шахматка начислений и оплат - 1 экз.", "Копия Устава", "Копия выписки из ЕГРЮЛ",
              "Копия доверенности на представителя", "Копия платежного поручения об уплате государственной пошлины",
              "Копия свидетельства о регистрации права / выписка из ЕГРН"]
    if case.legal_costs:
        items += [f"Копия договора {_val(case.legal_contract)} - 1 экз.",
                  "Копия платежного поручения об оплате юридических услуг"]
    for t in items:
        n += 1
        _p(d, f"{n}. {t}", indent=False, after=0)
    _p(d, "", after=12)
    who = f"Представитель {cname}" if rep.name or rep.sign_name else f"Взыскатель {cname}"
    _p(d, f"{who}                                                      _________________ "
          f"{_val(rep.sign_name or c.sign_name, '________')}", indent=False)
    d.save(path)
