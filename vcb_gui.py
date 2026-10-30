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
    for key, flag in (("cap_payments", "--cap-payments"), ("cap_exclude_services", "--cap-exclude-services")):
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
    ("street", "Улица", "Кировоградская"),
    ("house", "Дом", "180"),
    ("flat", "Квартира", "24"),
    ("area", "Площадь, кв. м", "47,4"),
    ("birth_date", "Дата рождения", "дд.мм.гггг (для судебного приказа)"),
    ("birth_place", "Место рождения", "для судебного приказа"),
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
                    "birth_place", "period_from", "period_to"]


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
COLORS = {"bg": "#12151b", "card": "#1b1f27", "field": "#252a34", "line": "#313746", "text": "#e8ebf2",
          "muted": "#8d96a8", "accent": "#4c9aff", "accent_hover": "#3b86e8", "ok": "#4cc38a", "warn": "#f0b45c"}


class App:
    def __init__(self):
        import customtkinter as ctk

        self.ctk = ctk
        ctk.set_appearance_mode("dark")
        self.win = ctk.CTk(fg_color=COLORS["bg"])
        from vcb import __version__  # noqa: PLC0415

        self.win.title(f"vcb {__version__} — расчёт задолженности по ЖКУ и заявление в суд")
        self.win.geometry("1040x860")
        self.win.minsize(900, 680)
        from vcb.config import DATA, ROOT  # noqa: PLC0415

        self.root_dir, self.data_dir = ROOT, DATA
        self.history = History()
        self.vars: dict[str, tk.StringVar] = {}
        self.entries: dict[str, object] = {}
        self._results: queue.Queue = queue.Queue()
        self.mode = tk.StringVar(value="Один должник")
        self._build()
        self._load_last()

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
        ctk.CTkLabel(cell, text=label, text_color=COLORS["muted"], anchor="w", font=ctk.CTkFont(size=12)
                     ).pack(fill="x", padx=2)
        ent = ctk.CTkEntry(cell, textvariable=self.var(f"{prefix}.{key}"), placeholder_text=hint, height=38,
                           corner_radius=12, fg_color=COLORS["field"], border_color=COLORS["line"],
                           text_color=COLORS["text"], placeholder_text_color="#5d6678")
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
        ctk.CTkLabel(cell, text=label, text_color=COLORS["muted"], anchor="w", font=ctk.CTkFont(size=12)
                     ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=2)
        ent = ctk.CTkEntry(cell, textvariable=self.var(key), placeholder_text=hint, height=38, corner_radius=12,
                           fg_color=COLORS["field"], border_color=COLORS["line"], text_color=COLORS["text"],
                           placeholder_text_color="#5d6678")
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
        w = self.win
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
        _, body = self.card(s, "6. Параметры расчёта")
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
            ctk.CTkLabel(cell, text=label, text_color=COLORS["muted"], anchor="w", font=ctk.CTkFont(size=12)
                         ).pack(fill="x", padx=2)
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
        self.log = ctk.CTkTextbox(bottom, height=120, corner_radius=12, fg_color=COLORS["bg"],
                                  text_color=COLORS["text"], font=("Consolas", 12), border_width=1,
                                  border_color=COLORS["line"])
        self.log.pack(fill="x", padx=16, pady=(0, 16))
        self.log.insert("end", "Здесь появится ход расчёта и замечания по каждому должнику.\n")
        self.log.configure(state="disabled")
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
        self.pk[key] = {"single": single, "batch": batch, "pay": pay_holder, "hint": svc_hint, "grid": svc_grid,
                        "btns": svc_btns, "vars": {}, "names": {}, "job": None}
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
        svc_hint.pack(fill="x", padx=8, pady=(8, 2))
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
        if single:
            self.batch_extra.pack_forget()
            self.debtor_card.pack(fill="x", padx=6, pady=7, before=self.claimant_card)
        else:
            self.batch_extra.pack(fill="x")
            self.debtor_card.pack_forget()

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
            missing = [t for _, (_, pay, _, t) in packs.items() if not pay]
            if missing and not messagebox.askyesno(
                    "vcb", "Не выбран реестр оплат: " + ", ".join(missing) + ".\n\nОплаты по этому пакету будут взяты "
                           "из самой шахматки — без назначения платежа и без отбора платежей по исполнительному "
                           "производству. Продолжить?"):
                return
            gku, kr = packs.get("gku"), packs.get("kr")
            source = (gku or kr)[0]
            v = {"source": gku[0] if gku else "", "payments": gku[1] if gku else "",
                 "exclude_services": gku[2] if gku else "",
                 "cap_source": kr[0] if kr else "", "cap_payments": kr[1] if kr else "",
                 "cap_exclude_services": kr[2] if kr else "",
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
