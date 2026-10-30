"""Исковое заявление / заявление о выдаче судебного приказа (docx). Текст - по образцам заказчика."""
from __future__ import annotations

import re
from pathlib import Path

import yaml
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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
THEME = {"gku": "по оплате за жилое помещение и коммунальные услуги",
         "kr": "по оплате взносов на капитальный ремонт общего имущества в многоквартирном доме",
         "both": "по оплате за жилое помещение и коммунальные услуги и по оплате взносов на капитальный ремонт"}
OBLIGATION = {
    "gku": "по оплате за жилое помещение и коммунальные услуги: не в полном объеме осуществлял(а) оплату "
           "предоставленных жилищно-коммунальных услуг",
    "kr": "по уплате взносов на капитальный ремонт общего имущества в многоквартирном доме: не в полном объеме "
          "осуществлял(а) уплату взносов",
    "both": "по оплате за жилое помещение и коммунальные услуги и по уплате взносов на капитальный ремонт: не в "
            "полном объеме осуществлял(а) оплату предоставленных жилищно-коммунальных услуг и уплату взносов",
}
DUTY = {"gku": "по оплате жилищно-коммунальных услуг", "kr": "по уплате взносов на капитальный ремонт",
        "both": "по оплате жилищно-коммунальных услуг и по уплате взносов на капитальный ремонт"}
ART169 = ("В соответствии со статьей 169 Жилищного кодекса Российской Федерации собственники помещений в "
          "многоквартирном доме обязаны уплачивать ежемесячные взносы на капитальный ремонт общего имущества в "
          "многоквартирном доме.")


FONT = "Times New Roman"
FONT_SIZE = Pt(14)  # ГОСТ Р 7.0.97-2016: шрифт 12-14 пт
LINE_SPACING = 1.15  # ГОСТ Р 7.0.97-2016: межстрочный интервал 1,0-1,5
INDENT = Cm(1.25)  # абзацный отступ
NBSP = "\u00a0"


def typo(text: str) -> str:
    """Неразрывные пробелы: разряды чисел, «руб.», «№ 5», «ст. 155», «д. 180», «кв. 24» не переносятся по строкам."""
    text = re.sub(r"(?<=\d) (?=\d{3}\b)", NBSP, text)
    text = re.sub(r"(?<=\d) (?=(?:руб|₽|%|экз|кв\.\s?м))", NBSP, text)
    text = re.sub(r"№ ?(?=\S)", "№" + NBSP, text)
    text = re.sub(r"\b(ст\.ст|ст|ч|п|пп|д|кв|ул|тел|стр)\.\s*(?=[\d№])", lambda m: m.group(1) + "." + NBSP, text)
    text = re.sub(r"(?<![А-Яа-яЁё])([ВвКкОоСсУуИи]) (?=\S)", lambda m: m.group(1) + NBSP, text)  # предлоги
    return text


def _set_fonts(style):
    style.font.name, style.font.size = FONT, FONT_SIZE
    rpr = style.element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.append(fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(attr), FONT)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "ru-RU")
    lang.set(qn("w:eastAsia"), "ru-RU")


def _doc() -> Document:
    """A4; поля по ГОСТ Р 7.0.97-2016: левое 30, правое 10, верхнее и нижнее 20 мм; номера страниц сверху по центру."""
    d = Document()
    s = d.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin, s.right_margin, s.top_margin, s.bottom_margin = Cm(3), Cm(1), Cm(2), Cm(2)
    s.header_distance = Cm(1)
    s.different_first_page_header_footer = True  # номер не ставится на первой странице
    st = d.styles["Normal"]
    _set_fonts(st)
    pf = st.paragraph_format
    pf.space_before = pf.space_after = Pt(0)
    pf.line_spacing = LINE_SPACING
    pf.widow_control = True
    # номер страницы
    para = s.header.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run()
    for kind, text in (("begin", None), (None, " PAGE "), ("end", None)):
        el = OxmlElement("w:fldChar" if kind else "w:instrText")
        if kind:
            el.set(qn("w:fldCharType"), kind)
        else:
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)
    run.font.size = Pt(12)
    return d


def _p(d, text="", *, bold=False, align=None, indent=True, after=0, before=0, keep_next=False, single=False,
       hanging=None, container=None):
    """Абзац основного текста: по ширине, абзацный отступ 1,25 см, интервал 1,15."""
    p = (container or d).add_paragraph()
    pf = p.paragraph_format
    if hanging is not None:
        pf.left_indent, pf.first_line_indent = hanging, -hanging
    elif indent:
        pf.first_line_indent = INDENT
    pf.space_before, pf.space_after = Pt(before), Pt(after)
    pf.keep_with_next = keep_next
    if single:
        pf.line_spacing = 1.0
    p.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(typo(text))
    r.bold = bold
    return p


def _header_block(d, groups: list[list[tuple[str, bool]]]):
    """Реквизиты адресата справа: невидимая таблица, правая колонка 8,5 см, одинарный интервал."""
    table = d.add_table(rows=1, cols=2)
    table.autofit = False
    left, right = table.rows[0].cells
    left.width, right.width = Cm(7.5), Cm(9.5)
    right.paragraphs[0]._p.getparent().remove(right.paragraphs[0]._p)
    for gi, group in enumerate(groups):
        for li, (text, bold) in enumerate(group):
            p = right.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(6 if li == len(group) - 1 and gi < len(groups) - 1 else 0)
            p.add_run(typo(text)).bold = bold
    return table


def _signature(d, left_text: str, right_text: str):
    table = d.add_table(rows=1, cols=2)
    table.autofit = False
    a, b = table.rows[0].cells
    a.width, b.width = Cm(9), Cm(8)
    pa, pb = a.paragraphs[0], b.paragraphs[0]
    pa.paragraph_format.line_spacing = pb.paragraph_format.line_spacing = 1.0
    pa.add_run(typo(left_text))
    pb.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pb.add_run(typo(right_text))
    trpr = table.rows[0]._tr.get_or_add_trPr()
    trpr.append(OxmlElement("w:cantSplit"))
    return table


def _val(x, blank=BLANK):
    return x if x else blank


def _period(res):
    return f"с {res.period_from:%d.%m.%Y} года по {res.period_to:%d.%m.%Y} года"


def _money(x):
    return f"{fmt(x)} руб. ({in_words(x)})"


def write_claim(path, case: Case):
    order = case.doc_kind == "order"
    c, rep, dbt = case.claimant, case.representative, case.debtor
    cname = _val(c.short_name or c.name)
    fio_gen = dbt.extra.get("fio_gen") or dbt.fio
    asof = f"{case.as_of:%d.%m.%Y}г."
    cats = sorted(case.results, key=lambda k: 0 if k == SERVICES else 1)  # сначала ЖКУ, потом капремонт
    kinds = "both" if len(cats) > 1 else ("gku" if cats[0] == SERVICES else "kr")
    d = _doc()
    # --- шапка
    groups = [[(_val(case.court_name, "Суд (определите по адресу вручную)"), True), (case.court_address, False)],
              [(f"Взыскатель: {cname}", True)]
              + ([(f"ОГРН {_val(c.ogrn, '________')}", False), (f"ИНН {_val(c.inn, '________')}", False)]
                 if c.ogrn or c.inn else [])
              + [(_val(c.address, "адрес взыскателя"), False)]]
    if rep.name:
        groups.append([("Представитель взыскателя:", True), (rep.name, False)]
                      + ([(rep.address, False)] if rep.address else [])
                      + ([(f"тел. {rep.phone}", False)] if rep.phone else []))
    groups.append([("Должник:" if order else "Ответчик:", True), (_val(dbt.fio), True), (_val(dbt.address), False)]
                  + ([(f"Дата рождения: {_val(dbt.birth_date, '________')}", False),
                      (f"Место рождения: {_val(dbt.birth_place, '________')}", False)] if order else []))
    groups.append([(f"{'Размер требований' if order else 'Цена иска'}: {fmt(case.price)} руб.", False),
                   (f"Госпошлина: {fmt(case.fee)} руб.", False)])
    _header_block(d, groups)
    title = "ЗАЯВЛЕНИЕ о вынесении судебного приказа" if order else "ИСКОВОЕ ЗАЯВЛЕНИЕ"
    _p(d, title, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, before=18, keep_next=True)
    _p(d, f"о взыскании задолженности {THEME[kinds]}", align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, after=12,
       keep_next=True)
    # --- обстоятельства
    full = c.name or cname
    house = f"{dbt.house}" if dbt.house else BLANK
    _p(d, f"{full} осуществляет управление многоквартирным домом № {house}, расположенным по адресу: "
          f"{_val(dbt.city and 'г. ' + dbt.city)}{', ' + dbt.street_full if dbt.street else ''}.")
    _p(d, f"Собственником жилого помещения, расположенного по адресу: {_val(dbt.address)} является: "
          f"{_val(dbt.fio)}.")
    for t in STATIC["housing_law"]:
        _p(d, t)
    if CAPREPAIR in cats:
        _p(d, STATIC.get("caprepair_law", ART169))
    who_is = "Должник" if order else "Ответчик"
    _p(d, f"{who_is} {_val(dbt.fio)} ненадлежащим образом исполнял(а) свои обязанности {OBLIGATION[kinds]}, "
          f"допустил(а) образование задолженности.")
    for cat in cats:
        res = case.results[cat]
        _p(d, f"За период {_period(res)} размер задолженности {SUBJECT[cat]} составил по состоянию на {asof} "
              f"{_money(res.debt)} (Выписка из лицевого счета прилагается).")
    law = STATIC["penalty_law"]
    for t in ((law[0], law[1]) if SERVICES in cats else ()) + ((law[2],) if CAPREPAIR in cats else ()) + (law[3],):
        _p(d, t)
    key = case.results[cats[0]].rate
    _p(d, f"Размер ключевой ставки Банка России для расчета пени на момент подачи заявления составляет - "
          f"{fmt(key)}% годовых;")
    for cat in cats:
        res = case.results[cat]
        _p(d, f"Размер пени за несвоевременную оплату {SUBJECT[cat].replace('по оплате ', '')} за период "
              f"{_period(res)} составил по состоянию на {asof} {_money(res.penalty)} (Расчет пени прилагается).")
    _p(d, f"{who_is}, ненадлежащим образом выполняя свои обязанности {DUTY[kinds]}, нарушает требования "
          f"жилищного законодательства, а также права и законные интересы {'взыскателя' if order else 'истца'}.")
    for t in STATIC["costs_law"][:1]:
        _p(d, t)
    if case.legal_costs:
        _p(d, f"Для подготовки заявления о взыскании задолженности {THEME[kinds]} с {fio_gen} {cname} заключило "
              f"с {_val(case.legal_executor)} договор {_val(case.legal_contract)} на оказание юридических услуг, "
              f"стоимость услуг по договору составила {fmt(case.legal_costs)} руб.")
    for t in STATIC["costs_law"][1:]:
        _p(d, t)
    # --- требования
    articles = "30, 153, 154, 155, 158, 161" + (", 169" if CAPREPAIR in cats else "")
    parts = "; ".join(f"{SUBJECT[cat].replace('по оплате ', 'по оплате ', 1)} за период {_period(case.results[cat])} "
                      f"в размере {fmt(case.results[cat].debt)} руб., пени в размере "
                      f"{fmt(case.results[cat].penalty)} руб." for cat in cats)
    if order:
        _p(d, f"Требование является бесспорным: подтверждается документами о начислении платы и расчётом "
              f"задолженности. Спора о праве нет. Руководствуясь ст. {articles} Жилищного кодекса РФ, "
              f"ст.ст. 121, 122, 123, 124 Гражданского процессуального кодекса РФ,")
        _p(d, "ПРОШУ:", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, before=6, keep_next=True)
        _p(d, f"1. Выдать судебный приказ о взыскании с должника {fio_gen} в пользу {full} задолженности: "
              f"{parts} Всего по состоянию на {asof} {fmt(case.price)} руб.")
    else:
        _p(d, f"На основании изложенного, руководствуясь ст. {articles} Жилищного Кодекса РФ, "
              f"ст.ст. 88, 98, 131, 132 Гражданского процессуального кодекса РФ,")
        _p(d, "ПРОШУ:", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, before=6, keep_next=True)
        _p(d, f"1. Взыскать с должника {fio_gen} в пользу {full} задолженность: {parts} Пени — с последующим "
              f"доначислением по день фактической оплаты задолженности; всего по состоянию на {asof} на общую сумму "
              f"{fmt(case.price)} руб.")
    _p(d, f"2. Взыскать с должника {fio_gen} в пользу {full} расходы по оплате государственной пошлины в размере "
          f"{fmt(case.fee)} рублей.")
    if case.legal_costs:
        _p(d, f"3. Взыскать с должника {fio_gen} в пользу {full} судебные расходы в сумме "
              f"{fmt(case.legal_costs)} рублей.")
    # --- приложения
    _p(d, "Приложение:", bold=True, indent=False, before=12, keep_next=True)
    items = []
    if not order:
        items.append("копия чека об отправке искового заявления ответчику")
    for cat in cats:
        items += [f"Карточка расчетов по лицевому счету {ATT_SERVICES[cat]} - 1 экз.",
                  f"Расчет пени {ATT_SERVICES[cat]} - 1 экз.",
                  f"Шахматка начислений и оплат {ATT_SERVICES[cat]} - 1 экз."]
    items += ["Копия Устава", "Копия выписки из ЕГРЮЛ", "Копия доверенности на представителя",
              "Копия платежного поручения об уплате государственной пошлины",
              "Копия свидетельства о регистрации права / выписка из ЕГРН"]
    if CAPREPAIR in cats:
        items.append("Копия протокола общего собрания собственников о формировании фонда капитального ремонта")
    if case.legal_costs:
        items += [f"Копия договора {_val(case.legal_contract)} - 1 экз.",
                  "Копия платежного поручения об оплате юридических услуг"]
    for n, t in enumerate(items, 1):
        _p(d, f"{n}. {t[:1].upper() + t[1:]}", indent=False, hanging=Cm(0.9), single=True, after=2, keep_next=(n == len(items)))
    who = f"Представитель {cname}" if rep.name or rep.sign_name else f"Взыскатель {cname}"
    _p(d, "", indent=False, after=6, keep_next=True)
    _signature(d, who, f"_______________ {_val(rep.sign_name or c.sign_name, '________')}")
    d.save(path)
