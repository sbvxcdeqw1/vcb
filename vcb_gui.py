#!/usr/bin/env python3
"""Окно программы vcb: запускается двойным щелчком (или `python vcb_gui.py`), других файлов запускать не нужно.

Нужен только установленный Python 3.10+ (на Windows - с галочкой «tcl/tk and IDLE»). Недостающие библиотеки
(openpyxl, python-docx, pyyaml, xlrd, num2words) программа предложит установить сама при первом запуске.
"""
from __future__ import annotations

import contextlib
import io
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "src"))

REQUIRED = {"openpyxl": "openpyxl", "docx": "python-docx", "yaml": "pyyaml", "xlrd": "xlrd", "num2words": "num2words"}


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
    if not messagebox.askyesno("vcb", "Для работы нужны библиотеки:\n  " + ", ".join(missing) +
                               "\n\nУстановить сейчас? (нужен интернет, займёт около минуты)"):
        raise SystemExit(1)
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
    except Exception as e:  # noqa: BLE001
        messagebox.showerror("vcb", f"Не удалось установить библиотеки:\n{e}\n\nУстановите вручную:\n"
                                    f"pip install {' '.join(missing)}")
        raise SystemExit(1)
    root.destroy()


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
    if not v["source"]:
        raise ValueError("укажите выгрузку 1С (файл или папку)")
    argv = [v["source"], "-o", v.get("out") or "out", "--no-interactive", "--region", v.get("region") or "perm_krai"]
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
    return argv


DOC_CHOICES = {"автоматически (по сумме)": "auto", "судебный приказ": "order", "исковое заявление": "claim"}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("vcb — заявления в суд и карточки расчёта по ЖКУ")
        self.geometry("860x720")
        self.minsize(760, 620)
        from vcb.config import DATA, ROOT  # noqa: PLC0415

        self.root_dir, self.data_dir = ROOT, DATA
        self.vars = {k: tk.StringVar() for k in ("source", "payments", "registry", "profile", "overrides", "as_of",
                                                 "period_from", "period_to", "doc", "region", "out")}
        self.vars["as_of"].set(date.today().strftime("%d.%m.%Y"))
        self.vars["doc"].set(next(iter(DOC_CHOICES)))
        self.vars["out"].set(str(Path.home() / "vcb_результаты"))
        self._results: queue.Queue = queue.Queue()
        self._build()
        self._refresh_profiles()

    # ------------------------------------------------------------------ форма
    def _row(self, parent, r, label, key, browse=None, hint=""):
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", pady=3)
        ttk.Entry(parent, textvariable=self.vars[key]).grid(row=r, column=1, sticky="ew", padx=6)
        if browse:
            f = ttk.Frame(parent)
            f.grid(row=r, column=2, sticky="e")
            for text, cmd in browse:
                ttk.Button(f, text=text, command=cmd).pack(side="left", padx=2)
        if hint:
            ttk.Label(parent, text=hint, foreground="#666").grid(row=r + 1, column=1, columnspan=2, sticky="w", padx=6)

    def _build(self):
        pad = {"padx": 12, "pady": 6}
        src = ttk.LabelFrame(self, text="1. Исходные данные")
        src.pack(fill="x", **pad)
        src.columnconfigure(1, weight=1)
        self._row(src, 0, "Выгрузка 1С *", "source",
                  [("Файл…", lambda: self._pick("source", "file")), ("Папка…", lambda: self._pick("source", "dir"))])
        self._row(src, 2, "Реестр оплат", "payments", [("Файл…", lambda: self._pick("payments", "file"))],
                  "необязательно: реестр банковских документов — оплаты разносятся по назначению платежа")
        self._row(src, 4, "Реестр должников", "registry", [("Файл…", lambda: self._pick("registry", "file"))],
                  "необязательно: ФИО в род. падеже, адрес, площадь, период по каждому должнику (csv/xlsx)")

        who = ttk.LabelFrame(self, text="2. Взыскатель и представитель")
        who.pack(fill="x", **pad)
        who.columnconfigure(1, weight=1)
        ttk.Label(who, text="Профиль").grid(row=0, column=0, sticky="w", pady=3)
        self.profile_box = ttk.Combobox(who, textvariable=self.vars["profile"], state="readonly")
        self.profile_box.grid(row=0, column=1, sticky="ew", padx=6)
        f = ttk.Frame(who)
        f.grid(row=0, column=2)
        ttk.Button(f, text="Реквизиты…", command=self._edit_profile).pack(side="left", padx=2)
        ttk.Label(who, text="можно не выбирать — тогда в документах останутся пустые поля",
                  foreground="#666").grid(row=1, column=1, sticky="w", padx=6)

        par = ttk.LabelFrame(self, text="3. Параметры расчёта")
        par.pack(fill="x", **pad)
        par.columnconfigure(1, weight=1)
        par.columnconfigure(3, weight=1)
        ttk.Label(par, text="Дата расчёта").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(par, textvariable=self.vars["as_of"], width=14).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(par, text="Период с").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(par, textvariable=self.vars["period_from"], width=14).grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(par, text="по").grid(row=1, column=2, sticky="e")
        ttk.Entry(par, textvariable=self.vars["period_to"], width=14).grid(row=1, column=3, sticky="w", padx=6)
        ttk.Label(par, text="формат дд.мм.гггг; пусто — весь период выгрузки", foreground="#666").grid(
            row=2, column=1, columnspan=3, sticky="w", padx=6)
        ttk.Label(par, text="Документ").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Combobox(par, textvariable=self.vars["doc"], values=list(DOC_CHOICES), state="readonly",
                     width=26).grid(row=3, column=1, sticky="w", padx=6)
        ttk.Label(par, text="Регион").grid(row=3, column=2, sticky="e")
        regions = sorted(p.stem for p in (self.data_dir / "courts").glob("*.yaml")) or ["perm_krai"]
        self.vars["region"].set(regions[0])
        ttk.Combobox(par, textvariable=self.vars["region"], values=regions, state="readonly",
                     width=18).grid(row=3, column=3, sticky="w", padx=6)

        out = ttk.LabelFrame(self, text="4. Результат")
        out.pack(fill="x", **pad)
        out.columnconfigure(1, weight=1)
        self._row(out, 0, "Папка результата", "out", [("Выбрать…", lambda: self._pick("out", "dir"))])

        bar = ttk.Frame(self)
        bar.pack(fill="x", **pad)
        self.run_btn = ttk.Button(bar, text="Рассчитать", command=self._run)
        self.run_btn.pack(side="left")
        ttk.Button(bar, text="Открыть папку результата", command=self._open_out).pack(side="left", padx=8)
        self.log = scrolledtext.ScrolledText(self, height=12, state="disabled", font=("Consolas", 10))
        self.log.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    # ------------------------------------------------------------- действия
    def _pick(self, key, kind):
        if kind == "dir":
            path = filedialog.askdirectory(title="Выберите папку")
        else:
            path = filedialog.askopenfilename(title="Выберите файл", filetypes=[
                ("Таблицы и реестры", "*.xls *.xlsx *.csv"), ("Все файлы", "*.*")])
        if path:
            self.vars[key].set(path)

    def _profiles_dir(self) -> Path:
        return self.root_dir / "profiles"

    def _refresh_profiles(self, select: str | None = None):
        files = sorted(self._profiles_dir().glob("*.yaml"))
        self._profile_paths = {p.stem: str(p) for p in files}
        self.profile_box["values"] = ["(без профиля)"] + list(self._profile_paths)
        self.profile_box.set(select if select in self._profile_paths else "(без профиля)")

    def _edit_profile(self):
        from vcb.profiles import PARTY_FIELDS, REP_FIELDS, load_profile  # noqa: PLC0415
        import yaml  # noqa: PLC0415

        win = tk.Toplevel(self)
        win.title("Реквизиты взыскателя и представителя")
        win.grab_set()
        cur = self.vars["profile"].get()
        data = {}
        if cur in self._profile_paths:
            prof = load_profile(self._profile_paths[cur])
            data = {"claimant": prof["claimant"], "representative": prof["representative"]}
        entries: dict[tuple[str, str], tk.StringVar] = {}
        r = 0
        for block, title, fields in (("claimant", "Взыскатель", PARTY_FIELDS),
                                     ("representative", "Представитель", REP_FIELDS)):
            ttk.Label(win, text=title, font=("", 11, "bold")).grid(row=r, column=0, columnspan=2, sticky="w",
                                                                    padx=10, pady=(10, 2))
            r += 1
            for key, label in fields:
                ttk.Label(win, text=label).grid(row=r, column=0, sticky="w", padx=10, pady=2)
                var = tk.StringVar(value=getattr(data[block], key, "") if block in data else "")
                ttk.Entry(win, textvariable=var, width=60).grid(row=r, column=1, padx=10, pady=2)
                entries[(block, key)] = var
                r += 1
        ttk.Label(win, text="Сохранить как").grid(row=r, column=0, sticky="w", padx=10, pady=(12, 2))
        name = tk.StringVar(value=cur if cur in self._profile_paths else "мой")
        ttk.Entry(win, textvariable=name, width=30).grid(row=r, column=1, sticky="w", padx=10, pady=(12, 2))

        def save():
            fname = name.get().strip()
            if not fname:
                messagebox.showerror("vcb", "Укажите имя профиля", parent=win)
                return
            out = {b: {k: v.get().strip() for (bb, k), v in entries.items() if bb == b and v.get().strip()}
                   for b in ("claimant", "representative")}
            self._profiles_dir().mkdir(exist_ok=True)
            # *.local.yaml не попадает в git; имя для списка - без «.local»
            path = self._profiles_dir() / f"{fname}.local.yaml"
            path.write_text(yaml.safe_dump(out, allow_unicode=True, sort_keys=False), encoding="utf-8")
            self._refresh_profiles(select=path.name[:-5])
            win.destroy()

        ttk.Button(win, text="Сохранить", command=save).grid(row=r + 1, column=1, sticky="e", padx=10, pady=12)

    def _values(self) -> dict:
        v = {k: var.get().strip() for k, var in self.vars.items()}
        v["doc"] = DOC_CHOICES.get(v["doc"], "auto")
        v["profile"] = self._profile_paths.get(v["profile"], "")
        return v

    def _write(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _run(self):
        try:
            argv = build_argv(self._values())
        except ValueError as e:
            messagebox.showerror("vcb", str(e))
            return
        self.run_btn.configure(state="disabled")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self._write("Считаю…\n")
        threading.Thread(target=self._worker, args=(argv,), daemon=True).start()
        self.after(150, self._poll)

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
            self.after(150, self._poll)
            return
        self._done(text, code)

    def _done(self, text: str, code: int):
        self._write(text)
        self._write("\nГотово. Замечания (!) прочитайте до того, как использовать документы.\n" if code == 0
                    else "\nЗавершено с ошибками — см. строки FAIL выше.\n")
        self.run_btn.configure(state="normal")

    def _open_out(self):
        path = Path(self.vars["out"].get() or "out")
        path.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])


def main() -> None:
    ensure_dependencies()
    App().mainloop()


if __name__ == "__main__":
    main()
