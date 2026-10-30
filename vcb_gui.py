#!/usr/bin/env python3
"""Окно программы vcb: запускается двойным щелчком (или `python vcb_gui.py`), других файлов запускать не нужно.

Два режима: «Один должник» (данные вводятся прямо в окне) и «Пакет должников» (папка выгрузок + реестр).
Реквизиты взыскателя, представителя и должников запоминаются в истории (~/.vcb/history.json).

Нужен Python 3.10-3.13 (на Windows - с галочкой «tcl/tk and IDLE»). Недостающие библиотеки программа
предложит установить сама при первом запуске.
"""
from __future__ import annotations

import contextlib
import csv
import io
import json
import os
import queue
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "src"))

REQUIRED = {"openpyxl": "openpyxl", "docx": "python-docx", "yaml": "pyyaml", "xlrd": "xlrd",
            "num2words": "num2words", "customtkinter": "customtkinter"}


def ensure_dependencies() -> None:
    """Проверяет библиотеки; при отсутствии предлагает поставить их через pip."""
    missing = []
    for module, package in REQUIRED.items():
        try:
            __import__(module)
        except ImportError:
            missing.append(package)
    if not missing:
        return
    root = tk.Tk()
    root.withdraw()
    note = ""
    if sys.version_info >= (3, 14):
        note = (f"\n\nВнимание: у вас Python {sys.version_info.major}.{sys.version_info.minor} - для такой новой версии "
                "у части библиотек (lxml, pyyaml) ещё нет готовых сборок, установка может не пройти. "
                "Надёжнее поставить Python 3.12 или 3.13 с python.org.")
    if not messagebox.askyesno("vcb", "Для работы нужны библиотеки:\n  " + ", ".join(missing) +
                               "\n\nУстановить сейчас? (нужен интернет, займёт около минуты)" + note):
        raise SystemExit(1)
    proc = subprocess.run([sys.executable, "-m", "pip", "install", *missing], capture_output=True, text=True,
                          encoding="utf-8", errors="replace")
    if proc.returncode != 0:
        tail = "\n".join((proc.stdout + "\n" + proc.stderr).strip().splitlines()[-12:])
        messagebox.showerror("vcb", "Не удалось установить библиотеки.\n\nПоследние строки отчёта pip:\n" + tail +
                             "\n\nЧто делать: установите Python 3.12 или 3.13 (python.org) и запустите программу им. "
                             "Либо выполните вручную:\npip install " + " ".join(missing))
        raise SystemExit(1)
    root.destroy()


# ----------------------------------------------------------------------------- логика без окна
def parse_date(text: str) -> str | None:
    """дд.мм.гггг или гггг-мм-дд -> гггг-мм-дд; пусто -> None."""
    text = text.strip()
    if not text:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError(f"не понимаю дату «{text}», нужен формат дд.мм.гггг")


def build_argv(v: dict) -> list[str]:
    """Значения формы -> аргументы командной строки vcb (вынесено отдельно, чтобы тестировать без окна)."""
    if not v.get("source") and not v.get("cap_source"):
        raise ValueError("укажите шахматку ЖКУ и/или шахматку капитального ремонта")
    argv = ([v["source"]] if v.get("source") else []) + [
        "-o", v.get("out") or "out", "--no-interactive", "--region", v.get("region") or "perm_krai"]
    if v.get("cap_source"):
        argv += ["--cap-input", v["cap_source"]]
    for key, flag in (("cap_payments", "--cap-payments"), ("cap_exclude_services", "--cap-exclude-services"),
                      ("cap_overrides", "--cap-payment-overrides")):
        if v.get(key):
            argv += [flag, v[key]]
    for key, flag in (("payments", "--payments"), ("registry", "--registry"), ("profile", "--profile"),
                      ("overrides", "--payment-overrides")):
        if v.get(key):
            argv += [flag, v[key]]
    as_of = parse_date(v.get("as_of", "")) or date.today().isoformat()
    argv += ["--as-of", as_of]
    for key, flag in (("period_from", "--period-from"), ("period_to", "--period-to")):
        d = parse_date(v.get(key, ""))
        if d:
            argv += [flag, d]
    if v.get("doc") in ("order", "claim"):
        argv += ["--doc", v["doc"]]
    if v.get("exclude_services"):
        argv += ["--exclude-services", v["exclude_services"]]
    if v.get("no_claim"):
        argv.append("--no-claim")
    if v.get("allocation_report"):
        argv.append("--allocation-report")
    return argv


DEBTOR_FIELDS = [
    ("fio", "ФИО должника", "Иванов Иван Иванович"),
    ("fio_gen", "ФИО в родительном падеже", "Иванова Ивана Ивановича (для «взыскать с …»)"),
    ("city", "Населённый пункт", "Пермь"),
    ("street", "Улица", "Ленина"),
    ("house", "Дом", "10"),
    ("flat", "Квартира", "5"),
    ("area", "Площадь, кв. м", "45,2"),
    ("birth_date", "Дата рождения", "дд.мм.гггг (для судебного приказа)"),
    ("birth_place", "Место рождения", "для судебного приказа"),
    ("court", "Участок мирового судьи", "если сам не определился, например: № 3 Кировского района"),
]
CLAIMANT_FIELDS = [
    ("name", "Полное наименование", "Общество с ограниченной ответственностью «…»"),
    ("short_name", "Краткое наименование", "ООО «…»"),
    ("ogrn", "ОГРН", ""), ("inn", "ИНН", ""),
    ("address", "Адрес", "индекс, город, улица, дом"), ("phone", "Телефон", ""),
    ("sign_name", "Подпись в документах", "И.О.Фамилия"),
]
REP_FIELDS = [
    ("name", "ФИО представителя", "Иванов Иван Иванович"), ("sign_name", "Подпись в документах", "И.О.Фамилия"),
    ("address", "Адрес", ""), ("phone", "Телефон", ""),
]
LEGAL_FIELDS = [
    ("executor", "Исполнитель по договору", "ООО «Юрист»"), ("contract", "Договор", "№ 00-00/юр от 01.01.2026г."),
    ("costs", "Расходы на юриста, ₽", "15000"),
]
PACKAGES = [("gku", "Жилищно-коммунальные услуги"), ("kr", "Капитальный ремонт")]
DOC_CHOICES = {"Определить автоматически (по сумме долга)": "auto", "Судебный приказ": "order",
               "Исковое заявление": "claim"}
REGISTRY_COLUMNS = ["file", "fio", "fio_gen", "city", "street", "house", "flat", "area", "birth_date",
                    "birth_place", "court", "period_from", "period_to"]


def write_registry_csv(path: Path, rows: list[dict]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=REGISTRY_COLUMNS, delimiter=";", extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def write_profile_yaml(path: Path, claimant: dict, rep: dict, legal: dict) -> None:
    import yaml

    out = {"claimant": {k: v for k, v in claimant.items() if v}, "representative": {k: v for k, v in rep.items() if v}}
    legal = {k: v for k, v in legal.items() if v}
    if legal:
        if "costs" in legal:
            try:
                legal["costs"] = float(str(legal["costs"]).replace(" ", "").replace(",", "."))
            except ValueError:
                raise ValueError("«Расходы на юриста» должны быть числом, например 15000") from None
        out["legal"] = legal
    path.write_text(yaml.safe_dump(out, allow_unicode=True, sort_keys=False), encoding="utf-8")


class History:
    """Запоминает введённые реквизиты: {вид: {подпись: {поле: значение}}}, последние сверху."""

    LIMIT = 60

    def __init__(self, path: Path | None = None):
        self.path = path or Path.home() / ".vcb" / "history.json"
        try:
            self.data = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            self.data = {}

    def labels(self, kind: str) -> list[str]:
        return list(reversed(list(self.data.get(kind, {}))))

    def get(self, kind: str, label: str) -> dict:
        return dict(self.data.get(kind, {}).get(label, {}))

    def remember(self, kind: str, label: str, values: dict) -> None:
        label = label.strip()
        if not label or not any(values.values()):
            return
        box = self.data.setdefault(kind, {})
        box.pop(label, None)
        box[label] = {k: v for k, v in values.items() if v}
        while len(box) > self.LIMIT:
            box.pop(next(iter(box)))
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(self.data, ensure_ascii=False, indent=1), encoding="utf-8")
        except OSError:
            pass


# ----------------------------------------------------------------------------- окно
COLORS = {"bg": "#1f2633", "card": "#2a3345", "field": "#35405a", "line": "#46536e", "text": "#eef1f7",
          "muted": "#a9b3c8", "accent": "#5aa2ff", "accent_hover": "#4690ee", "ok": "#5fd3a0", "warn": "#f3bd6a",
          "hint": "#8590a8"}
FONT_FAMILY = "Segoe UI"  # заменяется на Montserrat, если шрифт удалось загрузить

HELP = {
    "gku.source_file": "Excel-файл «Шахматка» из 1С по лицевому счёту должника (услуги по строкам, месяцы по столбцам). "
                       "Из него берутся начисления; ФИО, адрес и участок суда подставятся в раздел «Должник» сами.",
    "kr.source_file": "Такая же «Шахматка» из 1С, но по взносам на капитальный ремонт. Начисления берутся отсюда.",
    "gku.source_dir": "Папка, где лежат шахматки ЖКУ — по одному файлу на должника. Каждый файл — отдельный расчёт.",
    "kr.source_dir": "Папка с шахматками капитального ремонта. С файлами ЖКУ они сопоставляются по адресу квартиры.",
    "gku.payments": "Реестр банковских документов с оплатами (файл или папка). Подойдёт общий реестр банка: платежи "
                    "отберутся по лицевому счёту. Оплаты разносятся по периодам из назначения платежа.",
    "kr.payments": "Реестр оплат взносов на капитальный ремонт. Платежи по исполнительному производству программа "
                   "в зачёт не берёт; решение можно изменить галочками ниже.",
    "registry": "Необязательная таблица (csv/xlsx) с ФИО, адресами и периодами должников пакета. Пустые поля берутся "
                "из шапки выгрузок. Шаблон создаётся кнопкой ниже.",
    "out": "Папка, куда сохранятся шахматка, карточка расчёта, расчёт пени и заявление. Для каждого должника создаётся "
           "своя подпапка.",
    "debtor.fio": "Фамилия, имя, отчество должника. Попадает в заявление, шапки карточки и шахматки.",
    "debtor.fio_gen": "ФИО в родительном падеже для фраз «взыскать с …». Программа составляет его по правилам — "
                      "проверьте окончания.",
    "debtor.city": "Населённый пункт квартиры. Нужен для определения суда и участка мирового судьи.",
    "debtor.street": "Улица без слова «ул.». Вместе с домом определяет участок мирового судьи и попадает в заявление.",
    "debtor.house": "Номер дома (с литерой или корпусом, если есть). Определяет участок мирового судьи.",
    "debtor.flat": "Номер квартиры. Выводится в адресе должника в документах.",
    "debtor.area": "Площадь помещения. Необязательно: указывается в заявлении, если известна.",
    "debtor.birth_date": "Нужна для заявления о выдаче судебного приказа (ст. 124 ГПК РФ). В иске не используется.",
    "debtor.birth_place": "Нужно для заявления о выдаче судебного приказа. В иске не используется.",
    "debtor.court": "Участок мирового судьи, например «№ 3 Кировского района». Программа определяет его по адресу; "
                    "если не смогла однозначно — впишите сами. От участка зависит шапка заявления.",
    "claimant.name": "Полное наименование организации-взыскателя — идёт в заявление («в пользу …») и в шапку.",
    "claimant.short_name": "Краткое наименование — подписывает шахматку и карточку расчёта.",
    "claimant.ogrn": "ОГРН взыскателя — выводится в реквизитах заявления.",
    "claimant.inn": "ИНН взыскателя — выводится в реквизитах заявления.",
    "claimant.address": "Адрес взыскателя для шапки заявления и почтовых отправлений.",
    "claimant.phone": "Контактный телефон взыскателя в заявлении.",
    "claimant.sign_name": "Как подписывать от имени взыскателя, например «И.О.Фамилия». Попадает в подпись документов.",
    "rep.name": "ФИО представителя взыскателя. Указывается в заявлении и в подписи под таблицами Excel.",
    "rep.sign_name": "Подпись представителя в формате «И.О.Фамилия» — внизу каждого Excel-документа и заявления.",
    "rep.address": "Адрес представителя для заявления.",
    "rep.phone": "Телефон представителя для заявления.",
    "legal.executor": "Организация или лицо, оказавшее юридические услуги по подготовке заявления.",
    "legal.contract": "Номер и дата договора на юридические услуги — для требования о возмещении судебных расходов.",
    "legal.costs": "Сумма расходов на юриста, ₽. Если заполнить, в заявлении появится требование о возмещении.",
    "par.as_of": "Дата, на которую считаются долг и пени (обычно сегодня). Формат дд.мм.гггг.",
    "par.period_from": "Начало расчётного периода. Начисления и оплаты до этой даты не учитываются. "
                       "Пусто — с первого начисления выгрузки.",
    "par.period_to": "Конец расчётного периода. Пусто — до последнего начисления выгрузки.",
    "par.doc": "Судебный приказ или исковое заявление. «Автоматически» — приказ при сумме до 500 000 ₽, иначе иск.",
    "par.region": "Регион, по справочнику судов которого определяется подсудность.",
}


class Tooltip:
    """Всплывающая подсказка у виджета: появляется при наведении и пропадает при уходе курсора."""

    def __init__(self, widget, text: str, delay: int = 350):
        self.widget, self.text, self.delay, self.tip, self.job = widget, text, delay, None, None
        widget.bind("<Enter>", self._enter, add="+")
        widget.bind("<Leave>", self._leave, add="+")
        widget.bind("<ButtonPress>", self._leave, add="+")

    def _enter(self, _e=None):
        self._cancel()
        self.job = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self.job:
            self.widget.after_cancel(self.job)
            self.job = None

    def _show(self):
        if self.tip or not self.text:
            return
        x, y = self.widget.winfo_pointerx() + 14, self.widget.winfo_pointery() + 18
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.attributes("-topmost", True)
        tk.Label(self.tip, text=self.text, justify="left", wraplength=340, bg=COLORS["field"], fg=COLORS["text"],
                 padx=12, pady=8, relief="solid", bd=1, font=(FONT_FAMILY, 10)).pack()
        self.tip.update_idletasks()
        sw, sh = self.widget.winfo_screenwidth(), self.widget.winfo_screenheight()
        x = min(x, sw - self.tip.winfo_width() - 10)
        y = y if y + self.tip.winfo_height() < sh - 10 else y - self.tip.winfo_height() - 36
        self.tip.wm_geometry(f"+{x}+{y}")

    def _leave(self, _e=None):
        self._cancel()
        if self.tip:
            self.tip.destroy()
            self.tip = None


class App:
    def __init__(self):
        import customtkinter as ctk

        self.ctk = ctk
        ctk.set_appearance_mode("dark")
        self._load_fonts(ctk)
        self.win = ctk.CTk(fg_color=COLORS["bg"])
        from vcb import __version__  # noqa: PLC0415

        self.win.title(f"vcb {__version__} — расчёт задолженности по ЖКУ и заявление в суд")
        self.win.geometry("1320x980")
        self.win.minsize(1100, 720)
        from vcb.config import DATA, ROOT  # noqa: PLC0415

        self.root_dir, self.data_dir = ROOT, DATA
        self.history = History()
        self.vars: dict[str, tk.StringVar] = {}
        self.entries: dict[str, object] = {}
        self._results: queue.Queue = queue.Queue()
        self.mode = tk.StringVar(value="Один должник")
        self.auto_vals: dict[str, str] = {}
        self._build()
        self._install_clipboard()
        self._load_last()


    # ---- шрифт Montserrat (лежит в папке fonts рядом с программой)
    @staticmethod
    def _load_fonts(ctk):
        global FONT_FAMILY
        loaded = False
        for name in ("Montserrat-Regular.ttf", "Montserrat-Bold.ttf"):
            f = HERE / "fonts" / name
            if f.is_file():
                try:
                    loaded = bool(ctk.FontManager.load_font(str(f))) or loaded
                except Exception:  # noqa: BLE001
                    pass
        if loaded:
            FONT_FAMILY = "Montserrat"
            ctk.ThemeManager.theme["CTkFont"]["family"] = "Montserrat"

    # ---- буфер обмена: Ctrl+C/V/X/A работают в любой раскладке + меню по правому щелчку
    def _install_clipboard(self):
        win = self.win
        latin = {"v", "c", "x", "a"}
        by_keycode = {86: "<<Paste>>", 67: "<<Copy>>", 88: "<<Cut>>", 65: "select"}   # Windows: код клавиши
        by_keysym = {"cyrillic_em": "<<Paste>>", "cyrillic_es": "<<Copy>>", "cyrillic_che": "<<Cut>>",
                     "cyrillic_ef": "select"}                                           # Linux/Mac: русские символы

        def editable(widget):
            return isinstance(widget, (tk.Entry, tk.Text))

        def select_all(widget):
            if isinstance(widget, tk.Entry):
                widget.select_range(0, "end")
                widget.icursor("end")
            else:
                widget.tag_add("sel", "1.0", "end")

        def on_key(e):
            if not editable(e.widget) or e.keysym.lower() in latin:   # латиница работает штатно
                return None
            action = by_keysym.get(e.keysym.lower()) or (by_keycode.get(e.keycode) if sys.platform == "win32" else None)
            if not action:
                return None
            if action == "select":
                select_all(e.widget)
            else:
                e.widget.event_generate(action)
            return "break"

        menu = tk.Menu(win, tearoff=0, bg=COLORS["field"], fg=COLORS["text"], activebackground=COLORS["accent"],
                       activeforeground="#ffffff", bd=0, font=(FONT_FAMILY, 10))

        def popup(e):
            w = e.widget
            if not editable(w):
                return None
            w.focus_set()
            menu.delete(0, "end")
            readonly = str(w.cget("state")) == "disabled"
            for label, ev in (("Вырезать", "<<Cut>>"), ("Копировать", "<<Copy>>"), ("Вставить", "<<Paste>>")):
                menu.add_command(label=label, command=lambda ev=ev, w=w: w.event_generate(ev),
                                 state="disabled" if readonly and ev != "<<Copy>>" else "normal")
            menu.add_separator()
            menu.add_command(label="Выделить всё", command=lambda w=w: select_all(w))
            menu.tk_popup(e.x_root, e.y_root)
            return "break"

        self._on_key = on_key
        win.bind_all("<Control-KeyPress>", on_key, add="+")
        win.bind_all("<Button-3>", popup, add="+")
        if sys.platform == "darwin":
            win.bind_all("<Button-2>", popup, add="+")

    # ---- подписи полей с подсказкой
    def label(self, parent, text: str, help_text: str = ""):
        ctk = self.ctk
        row = ctk.CTkFrame(parent, fg_color="transparent")
        lab = ctk.CTkLabel(row, text=text, text_color=COLORS["muted"], anchor="w", font=ctk.CTkFont(size=12))
        lab.pack(side="left", padx=(2, 0))
        if help_text:
            icon = ctk.CTkLabel(row, text=" ⓘ", text_color=COLORS["accent"], cursor="question_arrow",
                                font=ctk.CTkFont(size=12))
            icon.pack(side="left")
            for widget in (icon, lab):
                Tooltip(widget, help_text)
        return row

    # ---- левая панель: разделы и чек-лист заполненности
    def _build_nav(self, parent):
        ctk = self.ctk
        self.nav = ctk.CTkFrame(parent, width=270, corner_radius=18, fg_color=COLORS["card"], border_width=1,
                                border_color=COLORS["line"])
        self.nav.grid(row=0, column=0, sticky="ns", padx=(18, 0), pady=(18, 18))
        self.nav.grid_propagate(False)
        ctk.CTkLabel(self.nav, text="Разделы", font=ctk.CTkFont(size=16, weight="bold"), text_color=COLORS["text"],
                     anchor="w").pack(fill="x", padx=18, pady=(18, 2))
        ctk.CTkLabel(self.nav, text="Щёлкните, чтобы перейти. Значок показывает, заполнен ли раздел.",
                     text_color=COLORS["hint"], anchor="w", justify="left", wraplength=230,
                     font=ctk.CTkFont(size=11)).pack(fill="x", padx=18, pady=(0, 10))
        self.nav_rows: list[dict] = []
        self.nav_foot = ctk.CTkLabel(self.nav, text="", text_color=COLORS["muted"], anchor="w", justify="left",
                                     wraplength=230, font=ctk.CTkFont(size=12))
        self.nav_foot.pack(side="bottom", fill="x", padx=18, pady=(8, 16))

    def _fill_nav(self):
        ctk = self.ctk
        items = [("docs", "1. Исходные документы", self.src_card), ("debtor", "2. Должник", self.debtor_card),
                 ("claimant", "3. Взыскатель", self.claimant_card), ("rep", "4. Представитель", self.rep_card),
                 ("legal", "5. Юридические расходы", self.legal_card), ("par", "6. Параметры расчёта", self.par_card)]
        for key, title, card in items:
            row = ctk.CTkFrame(self.nav, fg_color="transparent", corner_radius=12, cursor="hand2")
            row.pack(fill="x", padx=10, pady=2)
            icon = ctk.CTkLabel(row, text="○", width=26, font=ctk.CTkFont(size=16, weight="bold"),
                                text_color=COLORS["hint"])
            icon.grid(row=0, column=0, rowspan=2, padx=(8, 2), pady=6)
            name = ctk.CTkLabel(row, text=title, anchor="w", text_color=COLORS["text"], font=ctk.CTkFont(size=13))
            name.grid(row=0, column=1, sticky="ew", pady=(6, 0))
            detail = ctk.CTkLabel(row, text="", anchor="w", justify="left", wraplength=190, text_color=COLORS["hint"],
                                  font=ctk.CTkFont(size=11))
            detail.grid(row=1, column=1, sticky="ew", pady=(0, 6))
            row.columnconfigure(1, weight=1)
            for wdg in (row, icon, name, detail):
                wdg.bind("<Button-1>", lambda e, c=card: self.goto(c))
                wdg.bind("<Enter>", lambda e, r=row: r.configure(fg_color=COLORS["field"]))
                wdg.bind("<Leave>", lambda e, r=row: r.configure(fg_color="transparent"))
            self.nav_rows.append({"key": key, "row": row, "icon": icon, "detail": detail})
        self.nav_job = None
        watched = [k for k in self.vars if k not in ("out",)] + ["out"]
        for k in watched:
            self.vars[k].trace_add("write", lambda *a: self._schedule_nav())
        self._update_nav()

    def goto(self, card):
        self.win.update_idletasks()
        canvas = self.scroll._parent_canvas
        total = max(self.scroll.winfo_height(), 1)
        canvas.yview_moveto(max(0.0, min(1.0, (card.winfo_y() - 6) / total)))

    def _schedule_nav(self):
        if getattr(self, "nav_job", None):
            self.win.after_cancel(self.nav_job)
        self.nav_job = self.win.after(250, self._update_nav)

    def _checklist(self) -> dict:
        v = lambda k: self.vars[k].get().strip()
        single = self.mode.get() == "Один должник"
        res: dict[str, tuple[str, str]] = {}
        packs = [(t, k) for k, t in (("gku", "ЖКУ"), ("kr", "КР")) if self.source(k)]
        if not packs:
            res["docs"] = ("todo", "нужна шахматка ЖКУ или капремонта")
        else:
            no_reg = [t for t, k in packs if not v(f"{k}.payments")]
            detail = " · ".join(f"{t} {'✓' if t not in no_reg else '—'}" for t, _ in packs)
            res["docs"] = ("warn", f"нет реестра оплат: {', '.join(no_reg)}") if no_reg else \
                ("ok", f"{detail}, реестры выбраны")
        if single:
            need = [("fio", "ФИО"), ("city", "город"), ("street", "улица"), ("house", "дом"), ("flat", "кв.")]
            miss = [t for k, t in need if not v(f"debtor.{k}")]
            extra = [] if v("debtor.court") else ["участок суда"]
            if not miss:
                res["debtor"] = ("ok", ("не определён: " + ", ".join(extra)) if extra else "данные заполнены")
            else:
                res["debtor"] = ("warn" if len(miss) < len(need) else "todo", "не хватает: " + ", ".join(miss))
        else:
            res["debtor"] = ("opt", "берётся из выгрузок и реестра должников")
        c = v("claimant.name") or v("claimant.short_name")
        res["claimant"] = ("ok", "заполнен" + ("" if v("claimant.sign_name") else ", нет подписи")) if c else \
            ("opt", "необязательно")
        res["rep"] = ("ok", "заполнен" + ("" if v("rep.sign_name") else ", нет подписи")) if v("rep.name") else \
            ("opt", "необязательно")
        res["legal"] = ("ok", "расходы указаны") if v("legal.costs") else ("opt", "необязательно")
        bad = []
        if not parse_date(v("par.as_of")):
            bad.append("дата расчёта")
        if not v("out"):
            bad.append("папка результата")
        res["par"] = ("todo", "укажите: " + ", ".join(bad)) if bad else ("ok", "задано")
        return res

    def _update_nav(self):
        self.nav_job = None
        res = self._checklist()
        single = self.mode.get() == "Один должник"
        style = {"ok": ("✓", COLORS["ok"]), "warn": ("!", COLORS["warn"]), "todo": ("○", COLORS["accent"]),
                 "opt": ("–", COLORS["hint"])}
        for r in self.nav_rows:
            state, detail = res[r["key"]]
            sym, color = style[state]
            r["icon"].configure(text=sym, text_color=color)
            r["detail"].configure(text=detail, text_color=COLORS["warn"] if state == "warn" else COLORS["hint"])
        required = [k for k in ("docs", "debtor", "par") if k != "debtor" or single]
        done = sum(res[k][0] in ("ok", "warn") and not (k == "debtor" and res[k][0] == "warn") for k in required)
        ready = all(res[k][0] == "ok" or (k == "docs" and res[k][0] == "warn") for k in required)
        self.nav_foot.configure(
            text=f"Обязательных разделов заполнено: {done} из {len(required)}. "
                 + ("Можно считать." if ready else "Заполните отмеченные ○ и !."),
            text_color=COLORS["ok"] if ready else COLORS["muted"])

    # ---- предзаполнение из файлов
    def _schedule_autofill(self):
        if getattr(self, "af_job", None):
            self.win.after_cancel(self.af_job)
        self.af_job = self.win.after(800, self.autofill)

    def autofill(self):
        """Берёт ФИО, адрес и участок суда из шахматки/реестра и подставляет в пустые поля должника.
        Поля, которые оператор уже изменил руками, не затираются."""
        from vcb.autofill import autofill  # noqa: PLC0415

        self.af_job = None
        if self.mode.get() != "Один должник":
            return
        chess = next((self.vars[f"{k}.source_file"].get().strip() for k in ("gku", "kr")
                      if self.vars[f"{k}.source_file"].get().strip()), "")
        reg = next((self.vars[f"{k}.payments"].get().strip() for k in ("gku", "kr")
                    if self.vars[f"{k}.payments"].get().strip()), "")
        try:
            values, notes = autofill(chess, reg, self.regions.get(self.vars["par.region"].get(), "perm_krai"))
        except Exception as e:  # noqa: BLE001
            self.auto_hint.configure(text=f"Не удалось прочитать данные должника из файлов: {e}")
            return
        filled = []
        for key, *_ in DEBTOR_FIELDS:
            var = self.vars[f"debtor.{key}"]
            cur, last = var.get().strip(), self.auto_vals.get(key)
            new = values.get(key, "")
            if new and (not cur or cur == last):
                var.set(new)
                self.auto_vals[key] = new
                filled.append(key)
            elif not new and cur and cur == last:
                var.set("")
                self.auto_vals.pop(key, None)
        names = {k: t for k, t, _ in DEBTOR_FIELDS}
        text = ("Подставлено из файлов: " + ", ".join(names[k] for k in filled) + ". ") if filled else ""
        self.auto_hint.configure(text=(text + "; ".join(notes)).strip())

    # ---- элементы
    def var(self, key: str, value: str = "") -> tk.StringVar:
        self.vars[key] = tk.StringVar(value=value)
        return self.vars[key]

    def card(self, parent, title: str, hint: str = ""):
        ctk = self.ctk
        frame = ctk.CTkFrame(parent, corner_radius=18, fg_color=COLORS["card"], border_width=1,
                             border_color=COLORS["line"])
        frame.pack(fill="x", padx=6, pady=7)
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=16, weight="bold"), text_color=COLORS["text"],
                     anchor="w").pack(fill="x", padx=20, pady=(16, 0 if hint else 8))
        if hint:
            ctk.CTkLabel(frame, text=hint, text_color=COLORS["muted"], anchor="w", justify="left", wraplength=880
                         ).pack(fill="x", padx=20, pady=(2, 8))
        body = ctk.CTkFrame(frame, fg_color="transparent")
        body.pack(fill="x", padx=14, pady=(0, 16))
        return frame, body

    def field(self, parent, prefix: str, key: str, label: str, hint: str = "", row: int = 0, col: int = 0,
              span: int = 1):
        ctk = self.ctk
        cell = ctk.CTkFrame(parent, fg_color="transparent")
        cell.grid(row=row, column=col, columnspan=span, sticky="ew", padx=6, pady=4)
        self.label(cell, label, HELP.get(f"{prefix}.{key}", "")).pack(fill="x", padx=2)
        ent = ctk.CTkEntry(cell, textvariable=self.var(f"{prefix}.{key}"), placeholder_text=hint, height=38,
                           corner_radius=12, fg_color=COLORS["field"], border_color=COLORS["line"],
                           text_color=COLORS["text"], placeholder_text_color=COLORS["hint"])
        ent.pack(fill="x")
        self.entries[f"{prefix}.{key}"] = ent
        return ent

    def grid_fields(self, body, prefix: str, fields: list, cols: int = 2):
        for i in range(cols):
            body.columnconfigure(i, weight=1, uniform=prefix)
        for i, (key, label, hint) in enumerate(fields):
            self.field(body, prefix, key, label, hint, row=i // cols, col=i % cols)

    def pick_row(self, body, key: str, label: str, button: str, kind: str, hint: str = "", row: int = 0):
        ctk = self.ctk
        cell = ctk.CTkFrame(body, fg_color="transparent")
        cell.grid(row=row, column=0, sticky="ew", padx=6, pady=4)
        body.columnconfigure(0, weight=1)
        cell.columnconfigure(0, weight=1)
        self.label(cell, label, HELP.get(key, "")).grid(row=0, column=0, columnspan=2, sticky="ew", padx=2)
        ent = ctk.CTkEntry(cell, textvariable=self.var(key), placeholder_text=hint, height=38, corner_radius=12,
                           fg_color=COLORS["field"], border_color=COLORS["line"], text_color=COLORS["text"],
                           placeholder_text_color=COLORS["hint"])
        ent.grid(row=1, column=0, sticky="ew")
        btns = ctk.CTkFrame(cell, fg_color="transparent")
        btns.grid(row=1, column=1, padx=(8, 0))
        for text, k in ((("Файл…", "file"), ("Папка…", "dir")) if kind == "both" else ((button, kind),)):
            ctk.CTkButton(btns, text=text, width=100 if kind == "both" else 130, height=38, corner_radius=12,
                          fg_color=COLORS["field"], hover_color=COLORS["line"], border_width=1,
                          border_color=COLORS["line"], command=lambda k=k: self.pick(key, k)).pack(side="left", padx=2)
        self.entries[key] = ent

    def history_menu(self, body, kind: str, prefix: str, fields: list, row: int):
        """Выпадающий список «Выбрать из истории» над полями блока."""
        ctk = self.ctk
        cell = ctk.CTkFrame(body, fg_color="transparent")
        cell.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=(0, 6))
        ctk.CTkLabel(cell, text="Выбрать из истории:", text_color=COLORS["muted"]).pack(side="left", padx=(2, 8))
        menu = ctk.CTkOptionMenu(cell, values=["— пусто —"], width=420, height=34, corner_radius=12,
                                 fg_color=COLORS["field"], button_color=COLORS["line"],
                                 button_hover_color=COLORS["accent"], dropdown_fg_color=COLORS["field"],
                                 command=lambda lab, k=kind, p=prefix, f=fields: self.fill_from_history(k, p, f, lab))
        menu.pack(side="left")
        ctk.CTkButton(cell, text="Очистить поля", width=120, height=34, corner_radius=12, fg_color="transparent",
                      hover_color=COLORS["field"], text_color=COLORS["muted"], border_width=1,
                      border_color=COLORS["line"],
                      command=lambda p=prefix, f=fields: [self.vars[f"{p}.{k}"].set("") for k, *_ in f]
                      ).pack(side="left", padx=8)
        self.menus = getattr(self, "menus", {})
        self.menus[kind] = menu
        self.refresh_menu(kind)

    def refresh_menu(self, kind: str):
        labels = self.history.labels(kind)
        self.menus[kind].configure(values=labels or ["— пока ничего не сохранено —"])
        self.menus[kind].set("— выберите —" if labels else "— пока ничего не сохранено —")

    def fill_from_history(self, kind: str, prefix: str, fields: list, label: str):
        data = self.history.get(kind, label)
        if not data:
            return
        for key, *_ in fields:
            self.vars[f"{prefix}.{key}"].set(data.get(key, ""))

    # ---- сборка окна
    def _build(self):
        ctk = self.ctk
        root = self.win
        root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(0, weight=1)
        self._build_nav(root)
        w = ctk.CTkFrame(root, fg_color="transparent")
        w.grid(row=0, column=1, sticky="nsew")
        w.grid_columnconfigure(0, weight=1)
        w.grid_rowconfigure(1, weight=1)
        head = ctk.CTkFrame(w, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=24, pady=(18, 4))
        ctk.CTkLabel(head, text="Расчёт задолженности и заявление в суд", font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=COLORS["text"]).pack(side="left")
        self.mode_switch = ctk.CTkSegmentedButton(
            head, values=["Один должник", "Пакет должников"], variable=self.mode, command=self._switch_mode,
            height=38, corner_radius=12, fg_color=COLORS["card"], selected_color=COLORS["accent"],
            selected_hover_color=COLORS["accent_hover"], unselected_color=COLORS["card"],
            unselected_hover_color=COLORS["field"], font=ctk.CTkFont(size=14, weight="bold"))
        self.mode_switch.pack(side="right")

        self.scroll = ctk.CTkScrollableFrame(w, fg_color="transparent", corner_radius=0)
        self.scroll.grid(row=1, column=0, sticky="nsew", padx=18, pady=0)
        s = self.scroll

        # 1. входящие документы: два пакета - ЖКУ и капитальный ремонт
        self.src_card, body = self.card(s, "1. Исходные документы", "Заполните один пакет или оба. В каждом — шахматка "
                                        "из 1С и реестр оплат из банка; начисления берутся из шахматки, оплаты — из "
                                        "реестра. Если заполнены оба пакета, заявление и госпошлина строятся по общей "
                                        "сумме.")
        self.pk: dict[str, dict] = {}
        self.tabs = ctk.CTkTabview(body, corner_radius=14, fg_color=COLORS["bg"], border_width=1,
                                   border_color=COLORS["line"], segmented_button_fg_color=COLORS["card"],
                                   segmented_button_selected_color=COLORS["accent"],
                                   segmented_button_selected_hover_color=COLORS["accent_hover"],
                                   segmented_button_unselected_color=COLORS["card"],
                                   segmented_button_unselected_hover_color=COLORS["field"])
        self.tabs.pack(fill="x", padx=6, pady=(0, 6))
        for key, title in PACKAGES:
            self._build_package(self.tabs.add(title), key)
        self.batch_extra = ctk.CTkFrame(body, fg_color="transparent")
        self.pick_row(self.batch_extra, "registry", "Реестр должников (необязательно)", "Выбрать файл…", "file",
                      "Таблица с ФИО, адресами и периодами; шаблон можно создать кнопкой ниже")
        batch_btns = ctk.CTkFrame(self.batch_extra, fg_color="transparent")
        batch_btns.grid(row=2, column=0, sticky="w", padx=6, pady=(2, 4))
        ctk.CTkButton(batch_btns, text="Создать шаблон реестра по папке с шахматками", height=34, corner_radius=12,
                      fg_color=COLORS["field"], hover_color=COLORS["line"], border_width=1,
                      border_color=COLORS["line"], command=self.make_registry_template).pack(side="left")
        ctk.CTkLabel(batch_btns, text="  Шахматки ЖКУ и капремонта одного должника сопоставляются по адресу квартиры.",
                     text_color=COLORS["muted"]).pack(side="left")

        # 2. должник (только режим «один»)
        self.debtor_card, body = self.card(s, "2. Должник", "Данные из выгрузки подставятся сами; "
                                           "заполните то, чего в ней нет. Введённое сохранится в истории.")
        self.history_menu(body, "debtor", "debtor", DEBTOR_FIELDS, 0)
        grid = ctk.CTkFrame(body, fg_color="transparent")
        grid.grid(row=1, column=0, columnspan=2, sticky="ew")
        body.columnconfigure((0, 1), weight=1)
        self.grid_fields(grid, "debtor", DEBTOR_FIELDS)
        self.auto_hint = ctk.CTkLabel(body, text="", text_color=COLORS["ok"], anchor="w", justify="left",
                                      wraplength=820, font=ctk.CTkFont(size=12))
        self.auto_hint.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 0))

        # 3. взыскатель
        self.claimant_card, body = self.card(s, "3. Взыскатель", "Можно не заполнять — в документах останутся пустые поля.")
        self.history_menu(body, "claimant", "claimant", CLAIMANT_FIELDS, 0)
        grid = ctk.CTkFrame(body, fg_color="transparent")
        grid.grid(row=1, column=0, columnspan=2, sticky="ew")
        body.columnconfigure((0, 1), weight=1)
        self.grid_fields(grid, "claimant", CLAIMANT_FIELDS)

        # 4. представитель
        self.rep_card, body = self.card(s, "4. Представитель взыскателя")
        self.history_menu(body, "rep", "rep", REP_FIELDS, 0)
        grid = ctk.CTkFrame(body, fg_color="transparent")
        grid.grid(row=1, column=0, columnspan=2, sticky="ew")
        body.columnconfigure((0, 1), weight=1)
        self.grid_fields(grid, "rep", REP_FIELDS)

        # 5. расходы на юриста
        self.legal_card, body = self.card(s, "5. Расходы на юридические услуги (необязательно)",
                                          "Если заполнить, в заявлении появится требование о взыскании судебных расходов.")
        self.grid_fields(body, "legal", LEGAL_FIELDS, cols=3)

        # 6. параметры
        self.par_card, body = self.card(s, "6. Параметры расчёта")
        body.columnconfigure((0, 1, 2), weight=1, uniform="par")
        today = date.today().strftime("%d.%m.%Y")
        self.field(body, "par", "as_of", "Рассчитать на дату", "дд.мм.гггг", 0, 0).insert(0, today)
        self.field(body, "par", "period_from", "Расчётный период с", "пусто — с начала выгрузки", 0, 1)
        self.field(body, "par", "period_to", "Расчётный период по", "пусто — до конца выгрузки", 0, 2)
        self.regions = {}
        for p in sorted((self.data_dir / "courts").glob("*.yaml")):
            name = p.stem
            try:
                import yaml  # noqa: PLC0415

                name = (yaml.safe_load(p.read_text(encoding="utf-8")) or {}).get("region_name") or p.stem
            except Exception:  # noqa: BLE001
                pass
            self.regions[name] = p.stem
        for col, (label, values, key) in enumerate((
                ("Какой документ подготовить", list(DOC_CHOICES), "doc"),
                ("Регион (для определения суда)", list(self.regions) or ["Пермский край"], "region"))):
            cell = ctk.CTkFrame(body, fg_color="transparent")
            cell.grid(row=1, column=col, sticky="ew", padx=6, pady=(10, 4))
            self.label(cell, label, HELP.get(f"par.{key}", "")).pack(fill="x", padx=2)
            menu = ctk.CTkOptionMenu(cell, values=values, variable=self.var(f"par.{key}", values[0]), height=38,
                                     corner_radius=12, fg_color=COLORS["field"], button_color=COLORS["line"],
                                     button_hover_color=COLORS["accent"], dropdown_fg_color=COLORS["field"])
            menu.pack(fill="x")
        opts = ctk.CTkFrame(body, fg_color="transparent")
        opts.grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=(10, 2))
        self.claim_var = tk.BooleanVar(value=True)
        self.alloc_var = tk.BooleanVar(value=False)
        for text, var in (("Подготовить заявление в суд (иск или судебный приказ)", self.claim_var),
                          ("Сохранить отчёт о разноске оплат (для проверки)", self.alloc_var)):
            ctk.CTkCheckBox(opts, text=text, variable=var, corner_radius=6, checkmark_color="#ffffff",
                            fg_color=COLORS["accent"], border_color=COLORS["line"]).pack(anchor="w", pady=3)
        out = ctk.CTkFrame(body, fg_color="transparent")
        out.grid(row=3, column=0, columnspan=3, sticky="ew")
        self.pick_row(out, "out", "Папка для готовых документов", "Выбрать папку…", "dir")
        self.vars["out"].set(str(Path.home() / "vcb_результаты"))

        # низ: кнопки и ход расчёта
        bottom = ctk.CTkFrame(w, fg_color=COLORS["card"], corner_radius=18, border_width=1,
                              border_color=COLORS["line"])
        bottom.grid(row=2, column=0, sticky="ew", padx=24, pady=(8, 18))
        bar = ctk.CTkFrame(bottom, fg_color="transparent")
        bar.pack(fill="x", padx=16, pady=(14, 6))
        self.run_btn = ctk.CTkButton(bar, text="Рассчитать и подготовить документы", height=44, corner_radius=14,
                                     font=ctk.CTkFont(size=15, weight="bold"), fg_color=COLORS["accent"],
                                     hover_color=COLORS["accent_hover"], command=self.run)
        self.run_btn.pack(side="left")
        ctk.CTkButton(bar, text="Открыть папку с результатами", height=44, corner_radius=14, fg_color=COLORS["field"],
                      hover_color=COLORS["line"], border_width=1, border_color=COLORS["line"],
                      command=self.open_out).pack(side="left", padx=10)
        self.progress = ctk.CTkProgressBar(bar, mode="indeterminate", width=140, progress_color=COLORS["accent"])
        self.status = ctk.CTkLabel(bar, text="", text_color=COLORS["muted"])
        self.status.pack(side="right")
        self.log = ctk.CTkTextbox(bottom, height=220, corner_radius=12, fg_color=COLORS["bg"],
                                  text_color=COLORS["text"], font=("Consolas", 14), border_width=1,
                                  border_color=COLORS["line"])
        self.log.pack(fill="x", padx=16, pady=(0, 16))
        self.log.insert("end", "Здесь появится ход расчёта и замечания по каждому должнику.\n")
        self.log.configure(state="disabled")
        for text, widget in (("Заявление готовится вместе с расчётами. Снимите галочку, если нужны только Excel-документы.",
                              opts.winfo_children()[0]),
                             ("Для проверки: какой платёж на какие начисления разнесён и что исключено.",
                              opts.winfo_children()[1])):
            Tooltip(widget, text)
        self._fill_nav()
        for k in ("gku.source_file", "kr.source_file", "gku.payments", "kr.payments", "par.region"):
            self.vars[k].trace_add("write", lambda *a: self._schedule_autofill())
        self._switch_mode(self.mode.get())

    def _build_package(self, tab, key: str):
        """Содержимое вкладки пакета: шахматка, реестр оплат, услуги с галочками."""
        ctk = self.ctk
        single = ctk.CTkFrame(tab, fg_color="transparent")
        batch = ctk.CTkFrame(tab, fg_color="transparent")
        label = "ЖКУ" if key == "gku" else "капитального ремонта"
        self.pick_row(single, f"{key}.source_file", f"Шахматка {label} (Excel-файл из 1С)", "Выбрать файл…", "file",
                      ".xls или .xlsx — «Шахматка» по лицевому счёту должника")
        self.pick_row(batch, f"{key}.source_dir", f"Папка с шахматками {label} (по одному файлу на должника)",
                      "Выбрать папку…", "dir", "Шахматки одного должника по ЖКУ и капремонту сопоставляются по адресу")
        pay_holder = ctk.CTkFrame(tab, fg_color="transparent")
        self.pick_row(pay_holder, f"{key}.payments", f"Реестр оплат {label}", "", "both",
                      "Реестр банковских документов (файл или папка). Общий реестр банка подойдёт: платежи отберутся "
                      "по лицевому счёту")
        svc_hint = ctk.CTkLabel(tab, text="Выберите шахматку — здесь появятся её услуги; снимите галочку с тех, "
                                          "что не нужно включать в расчёт.", text_color=COLORS["muted"], anchor="w",
                                justify="left", wraplength=860)
        svc_grid = ctk.CTkFrame(tab, fg_color="transparent")
        svc_btns = ctk.CTkFrame(tab, fg_color="transparent")
        pays_hint = ctk.CTkLabel(tab, text="", text_color=COLORS["muted"], anchor="w", justify="left", wraplength=860)
        pays_list = ctk.CTkScrollableFrame(tab, height=240, corner_radius=12, fg_color=COLORS["card"], border_width=1,
                                           border_color=COLORS["line"])
        pays_btns = ctk.CTkFrame(tab, fg_color="transparent")
        self.pk[key] = {"single": single, "batch": batch, "pay": pay_holder, "hint": svc_hint, "grid": svc_grid,
                        "btns": svc_btns, "vars": {}, "names": {}, "job": None, "pays_hint": pays_hint,
                        "pays_list": pays_list, "pays_btns": pays_btns, "pays": [], "manual": {}, "pjob": None}
        for text, mode in (("Зачесть все", "all"), ("Не зачитывать все", "none"), ("Как решила программа", "auto")):
            ctk.CTkButton(pays_btns, text=text, height=30, corner_radius=10, fg_color=COLORS["field"],
                          hover_color=COLORS["line"], border_width=1, border_color=COLORS["line"], width=170,
                          command=lambda m=mode, k=key: self.set_all_payments(k, m)).pack(side="left", padx=4)
        for pkey in (f"{key}.payments", f"{key}.source_file"):
            self.vars[pkey].trace_add("write", lambda *a, k=key: self._schedule_payments(k))
        for text, val in (("Отметить все", True), ("Снять все", False)):
            ctk.CTkButton(svc_btns, text=text, height=30, corner_radius=10, fg_color=COLORS["field"],
                          hover_color=COLORS["line"], border_width=1, border_color=COLORS["line"], width=130,
                          command=lambda v=val, k=key: [x.set(v) for x in self.pk[k]["vars"].values()]
                          ).pack(side="left", padx=4)
        for name in ("single", "batch"):
            self.vars[f"{key}.source_{'file' if name == 'single' else 'dir'}"].trace_add(
                "write", lambda *a, k=key: self._schedule_services(k))
        single.pack(fill="x")
        pay_holder.pack(fill="x")
        pays_hint.pack(fill="x", padx=8, pady=(6, 2))
        svc_hint.pack(fill="x", padx=8, pady=(12, 2))
        svc_grid.pack(fill="x")
        svc_btns.pack(fill="x", pady=(6, 0))

    def _switch_mode(self, mode: str):
        single = mode == "Один должник"
        for key, pk in self.pk.items():
            if single:
                pk["batch"].pack_forget()
                pk["single"].pack(fill="x", before=pk["pay"])
            else:
                pk["single"].pack_forget()
                pk["batch"].pack(fill="x", before=pk["pay"])
            self.refresh_services(key)
            self.refresh_payments(key)
        if hasattr(self, "nav_rows"):
            self._update_nav()
        if single:
            self.batch_extra.pack_forget()
            self.debtor_card.pack(fill="x", padx=6, pady=7, before=self.claimant_card)
        else:
            self.batch_extra.pack(fill="x")
            self.debtor_card.pack_forget()

    # ---- платежи реестра: ручной зачёт
    def _schedule_payments(self, key: str):
        pk = self.pk[key]
        if pk["pjob"]:
            self.win.after_cancel(pk["pjob"])
        pk["pjob"] = self.win.after(600, lambda: self.refresh_payments(key))

    @staticmethod
    def _show_pays(pk: dict, show: bool):
        for w in (pk["pays_list"], pk["pays_btns"]):
            w.pack_forget()
        if show:
            pk["pays_list"].pack(fill="x", padx=6, before=pk["hint"])
            pk["pays_btns"].pack(fill="x", pady=(6, 4), before=pk["hint"])

    def refresh_payments(self, key: str):
        """Читает платежи реестра, показывает решение программы и даёт поставить/снять галочку «в зачёт»."""
        from vcb.config import load_config  # noqa: PLC0415
        from vcb.money import fmt  # noqa: PLC0415
        from vcb.parsers.onec import parse_any  # noqa: PLC0415
        from vcb.parsers.payments import KIND_TITLES, classify, filter_by_account, parse_registers  # noqa: PLC0415

        pk = self.pk[key]
        pk["pjob"] = None
        for w in pk["pays_list"].winfo_children():
            w.destroy()
        pk["pays"] = []
        single = self.mode.get() == "Один должник"
        path = self.vars[f"{key}.payments"].get().strip()
        if not single:
            pk["pays_hint"].configure(text="В пакетном режиме платежи отбирает программа (исполнительное производство "
                                           "не зачитывается). Ручной выбор платежей — в режиме «Один должник».")
            self._show_pays(pk, False)
            return
        if not path or not Path(path).exists():
            pk["pays_hint"].configure(text="Выберите реестр оплат — здесь появятся его платежи; программа сама решит, "
                                           "какие зачесть, а вы сможете исправить это галочками.")
            self._show_pays(pk, False)
            return
        try:
            pays = parse_registers(path)
            chess = self.vars[f"{key}.source_file"].get().strip()
            number = parse_any(chess).number if chess and Path(chess).is_file() else ""
            pays, _ = filter_by_account(pays, number)
            classify(pays, load_config())
        except Exception as e:  # noqa: BLE001
            pk["pays_hint"].configure(text=f"Не удалось прочитать реестр оплат: {e}")
            self._show_pays(pk, False)
            return
        self._show_pays(pk, True)
        seen: dict[tuple, int] = {}
        for p in pays:
            k2 = (p.date, p.amount)
            seen[k2] = seen.get(k2, 0) + 1
            n = seen[k2]
            default = p.include
            var = tk.BooleanVar(value=pk["manual"].get((p.date, p.amount, n), default))
            var.trace_add("write", lambda *a, k=key: self._update_pay_summary(k))
            row = self.ctk.CTkFrame(pk["pays_list"], fg_color="transparent")
            row.pack(fill="x", pady=2)
            self.ctk.CTkCheckBox(row, text=f"{p.date:%d.%m.%Y}   {fmt(p.amount):>13} ₽   ·   {KIND_TITLES[p.kind]}",
                                 variable=var, corner_radius=6, checkmark_color="#ffffff", fg_color=COLORS["accent"],
                                 border_color=COLORS["line"], text_color=COLORS["text"]).pack(anchor="w")
            self.ctk.CTkLabel(row, text=p.info[:110], text_color=COLORS["muted"], anchor="w", justify="left",
                              font=self.ctk.CTkFont(size=11)).pack(fill="x", padx=34)
            pk["pays"].append({"payment": p, "var": var, "default": default, "n": n})
        self._update_pay_summary(key)

    def _update_pay_summary(self, key: str):
        from vcb.money import fmt  # noqa: PLC0415

        pk = self.pk[key]
        rows = pk["pays"]
        if not rows:
            return
        on = [r for r in rows if r["var"].get()]
        changed = sum(1 for r in rows if r["var"].get() != r["default"])
        total = sum((r["payment"].amount for r in rows), Decimal(0))
        zach = sum((r["payment"].amount for r in on), Decimal(0))
        for r in rows:
            p = r["payment"]
            pk["manual"][(p.date, p.amount, r["n"])] = r["var"].get()
        pk["pays_hint"].configure(text=(
            f"Платежей в реестре: {len(rows)} на {fmt(total)} ₽. В зачёт: {len(on)} на {fmt(zach)} ₽. "
            + (f"Изменено вручную: {changed}. " if changed else "Решение программы не менялось. ")
            + "Платежи вне расчётного периода в расчёт не входят, даже если отмечены."))

    def set_all_payments(self, key: str, mode: str):
        for r in self.pk[key]["pays"]:
            r["var"].set(r["default"] if mode == "auto" else mode == "all")
        if mode == "auto":
            self.pk[key]["manual"].clear()

    def payment_overrides_csv(self, key: str, folder: Path) -> str:
        """Решения оператора, отличающиеся от решения программы, - в csv для `--payment-overrides`."""
        rows = [r for r in self.pk[key]["pays"] if r["var"].get() != r["default"]]
        if not rows:
            return ""
        path = folder / f"overrides_{key}.csv"
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["дата", "сумма", "решение", "комментарий", "номер"])
            for r in rows:
                p = r["payment"]
                w.writerow([f"{p.date:%d.%m.%Y}", str(p.amount), "зачесть" if r["var"].get() else "исключить",
                            "вручную в окне", r["n"]])
        return str(path)

    # ---- услуги шахматки
    def _schedule_services(self, key: str):
        pk = self.pk[key]
        if pk["job"]:
            self.win.after_cancel(pk["job"])
        pk["job"] = self.win.after(500, lambda: self.refresh_services(key))

    def source(self, key: str) -> str:
        single = self.mode.get() == "Один должник"
        return self.vars[f"{key}.source_file" if single else f"{key}.source_dir"].get().strip()

    def _chess_files(self, key: str) -> list[Path]:
        raw = self.source(key)
        path = Path(raw) if raw else None
        if not path or not path.exists():
            return []
        if path.is_file():
            return [path]
        return sorted(f for f in path.iterdir() if f.suffix.lower() in (".xls", ".xlsx") and not f.name.startswith("~"))

    def refresh_services(self, key: str):
        """Читает услуги из выбранных шахматок пакета и показывает по галочке на каждую."""
        from vcb.money import fmt  # noqa: PLC0415
        from vcb.parsers.chess import list_services  # noqa: PLC0415
        from vcb.pipeline import norm_service  # noqa: PLC0415

        pk = self.pk[key]
        pk["job"] = None
        totals: dict[str, float] = {}
        names: dict[str, str] = {}
        files = self._chess_files(key)
        for f in files:
            try:
                for name, total in list_services(f):
                    k = norm_service(name)
                    names.setdefault(k, name)
                    totals[k] = totals.get(k, 0) + float(total)
            except Exception:  # noqa: BLE001
                continue
        previous = {k: v.get() for k, v in pk["vars"].items()}
        for w in pk["grid"].winfo_children():
            w.destroy()
        pk["vars"], pk["names"] = {}, names
        if not names:
            pk["hint"].configure(text="Выберите шахматку — здесь появятся её услуги; снимите галочку с тех, что не "
                                      "нужно включать в расчёт." if not files else
                                 "В выбранном файле не найдено услуг: это точно «Шахматка» из 1С?")
            return
        pk["hint"].configure(text=f"Найдено услуг: {len(names)}. Отмеченные участвуют в расчёте.")
        pk["grid"].columnconfigure((0, 1), weight=1, uniform=f"svc{key}")
        for i, (k, name) in enumerate(names.items()):
            var = tk.BooleanVar(value=previous.get(k, True))
            pk["vars"][k] = var
            amount = totals[k]
            label = f"{name}   ·   {fmt(Decimal(str(round(amount, 2))))} ₽" if amount else name
            self.ctk.CTkCheckBox(pk["grid"], text=label, variable=var, corner_radius=6, checkmark_color="#ffffff",
                                 fg_color=COLORS["accent"], border_color=COLORS["line"], text_color=COLORS["text"]
                                 ).grid(row=i // 2, column=i % 2, sticky="w", padx=8, pady=4)

    def excluded_services(self, key: str) -> str:
        pk = self.pk[key]
        return "; ".join(pk["names"][k] for k, v in pk["vars"].items() if not v.get())

    # ---- действия
    def pick(self, key: str, kind: str):
        if kind == "dir":
            path = filedialog.askdirectory(title="Выберите папку")
        else:
            path = filedialog.askopenfilename(title="Выберите файл", filetypes=[
                ("Таблицы", "*.xls *.xlsx *.csv"), ("Все файлы", "*.*")])
        if path:
            self.vars[key].set(path)

    def values(self, prefix: str, fields: list) -> dict:
        return {k: self.vars[f"{prefix}.{k}"].get().strip() for k, *_ in fields}

    def make_registry_template(self):
        folder = self.vars["gku.source_dir"].get().strip() or self.vars["kr.source_dir"].get().strip()
        if not folder or not Path(folder).is_dir():
            messagebox.showinfo("vcb", "Сначала выберите папку с шахматками.")
            return
        files = sorted(p.name for p in Path(folder).iterdir() if p.suffix.lower() in (".xls", ".xlsx"))
        target = filedialog.asksaveasfilename(title="Сохранить шаблон реестра", defaultextension=".csv",
                                              initialfile="реестр_должников.csv", initialdir=folder,
                                              filetypes=[("CSV (открывается в Excel)", "*.csv")])
        if not target:
            return
        write_registry_csv(Path(target), [{"file": f} for f in files])
        self.vars["registry"].set(target)
        messagebox.showinfo("vcb", "Шаблон сохранён. Откройте его в Excel, заполните ФИО, адрес, площадь "
                                   "и (по желанию) период по каждому должнику, сохраните как CSV и нажмите «Рассчитать».")

    def write(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def run(self):
        single = self.mode.get() == "Один должник"
        tmp = Path(tempfile.mkdtemp(prefix="vcb_"))
        try:
            packs = {}
            for key, title in PACKAGES:
                src = self.source(key)
                if not src:
                    continue
                if single and not Path(src).is_file():
                    raise ValueError(f"{title}: выбранный файл шахматки не найден")
                if not single and not Path(src).is_dir():
                    raise ValueError(f"{title}: выбранная папка с шахматками не найдена")
                pay = self.vars[f"{key}.payments"].get().strip()
                if pay and not Path(pay).exists():
                    raise ValueError(f"{title}: выбранный реестр оплат не найден")
                packs[key] = (src, pay, self.excluded_services(key), title)
            if not packs:
                raise ValueError("укажите шахматку ЖКУ и/или шахматку капитального ремонта")
            missing = [v_[3] for v_ in packs.values() if not v_[1]]
            if missing and not messagebox.askyesno(
                    "vcb", "Не выбран реестр оплат: " + ", ".join(missing) + ".\n\nОплаты по этому пакету будут взяты "
                           "из самой шахматки — без назначения платежа и без отбора платежей по исполнительному "
                           "производству. Продолжить?"):
                return
            if single:
                for key in packs:
                    ov = self.payment_overrides_csv(key, tmp)
                    packs[key] = (*packs[key][:3], packs[key][3], ov)
            gku, kr = packs.get("gku"), packs.get("kr")
            source = (gku or kr)[0]
            ov = lambda t: t[4] if t and len(t) > 4 else ""
            v = {"source": gku[0] if gku else "", "payments": gku[1] if gku else "",
                 "exclude_services": gku[2] if gku else "", "overrides": ov(gku),
                 "cap_source": kr[0] if kr else "", "cap_payments": kr[1] if kr else "",
                 "cap_exclude_services": kr[2] if kr else "", "cap_overrides": ov(kr),
                 "out": self.vars["out"].get().strip(),
                 "as_of": self.vars["par.as_of"].get(), "period_from": self.vars["par.period_from"].get(),
                 "period_to": self.vars["par.period_to"].get(),
                 "doc": DOC_CHOICES.get(self.vars["par.doc"].get(), "auto"),
                 "region": self.regions.get(self.vars["par.region"].get(), "perm_krai"),
                 "no_claim": not self.claim_var.get(), "allocation_report": self.alloc_var.get()}
            claimant = self.values("claimant", CLAIMANT_FIELDS)
            rep = self.values("rep", REP_FIELDS)
            legal = self.values("legal", LEGAL_FIELDS)
            if any(claimant.values()) or any(rep.values()) or any(legal.values()):
                write_profile_yaml(tmp / "profile.yaml", claimant, rep, legal)
                v["profile"] = str(tmp / "profile.yaml")
            debtor = {}
            if single:
                debtor = self.values("debtor", DEBTOR_FIELDS)
                if any(debtor.values()) and source:
                    write_registry_csv(tmp / "registry.csv", [{**debtor, "file": Path(source).name}])
                    v["registry"] = str(tmp / "registry.csv")
            elif self.vars["registry"].get().strip():
                v["registry"] = self.vars["registry"].get().strip()
            argv = build_argv(v)
        except (ValueError, OSError) as e:
            messagebox.showerror("vcb", str(e))
            return
        # запоминаем введённое
        self.history.remember("claimant", claimant.get("short_name") or claimant.get("name", ""), claimant)
        self.history.remember("rep", rep.get("name", ""), rep)
        if single and debtor.get("fio"):
            addr = ", ".join(x for x in (debtor.get("city"), debtor.get("street"), debtor.get("house"),
                                         debtor.get("flat")) if x)
            self.history.remember("debtor", f"{debtor['fio']} — {addr}" if addr else debtor["fio"], debtor)
        for kind in ("claimant", "rep", "debtor"):
            self.refresh_menu(kind)
        self.run_btn.configure(state="disabled", text="Считаю…")
        self.progress.pack(side="left", padx=14)
        self.progress.start()
        self.status.configure(text="")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.write("Считаю…\n")
        threading.Thread(target=self._worker, args=(argv,), daemon=True).start()
        self.win.after(150, self._poll)

    def _worker(self, argv):
        from vcb.cli import main  # noqa: PLC0415

        buf = io.StringIO()
        code = 1
        try:
            with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
                code = main(argv)
        except SystemExit as e:
            code = int(e.code or 0) if isinstance(e.code, int) else 1
        except Exception as e:  # noqa: BLE001
            buf.write(f"\nОшибка: {e}\n")
        self._results.put((buf.getvalue(), code))  # в окно пишет только главный поток

    def _poll(self):
        try:
            text, code = self._results.get_nowait()
        except queue.Empty:
            self.win.after(150, self._poll)
            return
        self.progress.stop()
        self.progress.pack_forget()
        self.run_btn.configure(state="normal", text="Рассчитать и подготовить документы")
        self.write(text)
        ok = code == 0
        self.write("\nГотово. Строки с «!» — замечания: прочитайте их до использования документов.\n" if ok
                   else "\nЗавершено с ошибками — см. строки FAIL выше.\n")
        self.status.configure(text="Готово" if ok else "Есть ошибки", text_color=COLORS["ok" if ok else "warn"])
        self._save_last()

    def open_out(self):
        path = Path(self.vars["out"].get() or "out")
        path.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])

    # ---- последние настройки
    LAST_KEYS = ("out", "par.doc", "par.region", "claimant.name", "claimant.short_name", "claimant.ogrn",
                 "claimant.inn", "claimant.address", "claimant.phone", "claimant.sign_name", "rep.name",
                 "rep.sign_name", "rep.address", "rep.phone", "legal.executor", "legal.contract", "legal.costs")

    def _save_last(self):
        self.history.data["last"] = {k: self.vars[k].get() for k in self.LAST_KEYS if k in self.vars}
        self.history.data["last"]["mode"] = self.mode.get()
        try:
            self.history.path.parent.mkdir(parents=True, exist_ok=True)
            self.history.path.write_text(json.dumps(self.history.data, ensure_ascii=False, indent=1), encoding="utf-8")
        except OSError:
            pass

    def _load_last(self):
        last = self.history.data.get("last", {})
        for k, val in last.items():
            if k in self.vars and val:
                if k == "par.doc" and val not in DOC_CHOICES or k == "par.region" and val not in self.regions:
                    continue
                self.vars[k].set(val)
        if last.get("mode") in ("Один должник", "Пакет должников"):
            self.mode.set(last["mode"])
            self._switch_mode(last["mode"])

    def mainloop(self):
        self.win.mainloop()


def main() -> None:
    ensure_dependencies()
    App().mainloop()


if __name__ == "__main__":
    main()
