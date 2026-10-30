"""Собирает справочник судов региона (yaml) из кэша страниц, скачанных tools/fetch_courts.py.

    python tools/build_courts.py КЭШ data/courts/perm_krai.yaml [manual.yaml]

manual.yaml - суды, которых нет в источнике (список courts в том же формате); они добавляются в конец.
"""
from __future__ import annotations

import glob
import html
import re
import sys
from pathlib import Path

import yaml

TYPES = r"(?:ул|пер|пр-кт|пр-т|просп|проезд|б-р|бул|ш|пл|наб|тупик|мкр|тракт|переулок)"
ITEM = re.compile(
    rf"(?P<type>{TYPES})\.\s+(?P<name>[А-ЯЁA-Z0-9][^,()]*?)\s+"
    r"(?P<houses>все дома|"
    r"\d[\d/а-яА-ЯёЁa-z\-]*(?:\s*\([^)]*\))?(?:,\s*\d[\d/а-яА-ЯёЁa-z\-]*(?:\s*\([^)]*\))?)*)"
    rf"(?=,\s*{TYPES}\.|\.\s*$|\s*$)")


def page_text(path) -> str:
    b = Path(path).read_bytes().decode("utf-8", "ignore")
    t = re.sub(r"<script.*?</script>|<style.*?</style>", "", b, flags=re.S)
    t = re.sub(r"<br\s*/?>|</p>|</div>|</tr>|</li>|</h\d>", "\n", t)
    t = html.unescape(re.sub(r"<[^>]+>", "", t))
    t = re.sub(r"[ \t\xa0]+", " ", t)
    return re.sub(r"\n\s*\n+", "\n", t)


def field(t: str, label: str) -> str:
    m = re.search(rf"^\s*{label}:\s*(.+)$", t, flags=re.M)
    return m.group(1).strip() if m else ""


def parse_streets(terr: str) -> list[dict]:
    """Список «ул. Название дома, ...»: берём серии из ≥2 подряд идущих элементов (так отсекается описание границ)."""
    items = list(ITEM.finditer(terr))
    runs, cur = [], []
    for m in items:
        if cur and m.start() - cur[-1].end() <= 3:
            cur.append(m)
        else:
            if cur:
                runs.append(cur)
            cur = [m]
    if cur:
        runs.append(cur)
    streets = []
    for run in runs:
        if len(run) < 2:
            continue
        for m in run:
            houses = "all" if m.group("houses") == "все дома" else m.group("houses").strip()
            streets.append({"street": m.group("name").strip(), "houses": houses})
    return streets


NAME_PAREN = re.compile(r"(?<![\w(])(?:(?:ул|пер|пр-кт|проезд|б-р|ш|пл|наб)\.\s*)?"
                        r"(?P<name>[А-ЯЁ][а-яё\-\.]+(?:\s(?:[А-ЯЁ][а-яё\-\.]+|\d+-[яйе]{1,2}))*)\s*\(")


def parse_streets_paren(terr: str) -> list[dict]:
    """Формат «Мира (45-63 (нечетные), 45а) Нефтяников (1-18) ...» (как в приложении к закону)."""
    ms = list(NAME_PAREN.finditer(terr))
    out = []
    for i, m in enumerate(ms):
        end = ms[i + 1].start() if i + 1 < len(ms) else len(terr)
        houses = terr[m.end():end].strip().rstrip(";,. ").rstrip(")").strip()
        if houses and re.match(r"\d", houses):
            out.append({"street": m.group("name").strip(), "houses": houses})
    return out if len(out) >= 2 else []


def parse_participant(path) -> dict | None:
    t = page_text(path)
    titles = [x for x in re.findall(r"^(Судебный участок № \d+ [^\n]+)$", t, flags=re.M) if "sudinforf" not in x]
    title = re.match(r"(.*)", titles[0]) if titles else None
    m = re.search(r"границах территории:\s*(.*?)\nКакие дела", t, flags=re.S)
    if not title or not m:
        return None
    terr = re.sub(r"\s+", " ", m.group(1)).strip()
    head = re.match(r"(?:г\.|пгт\.?|п\.|с\.|д\.)?\s*([^:,]{2,40}):", terr)
    city = ""
    if head and not head.group(1).startswith("Пермский край"):
        city = head.group(1).strip()
    num = re.search(r"№ (\d+)", title.group(1)).group(1)
    return {
        "name": title.group(1).strip(), "number": int(num),
        "address": field(t, "Адрес"), "judicial_district": field(t, "Судебный район"),
        "parent_name": field(t, "Вышестоящий суд"), "judge": field(t, "Мировой судья"),
        "city": city, "text": terr,
        "streets": (parse_streets(terr) or parse_streets_paren(terr)) if city else [],
    }


def parse_court(path) -> dict | None:
    t = page_text(path)
    m = re.search(r"Все суды » [^\n]+\n([^\n]+)\n", t)
    if not m or "Адрес:" not in t:
        return None
    return {"name": m.group(1).strip(), "address": field(t, "Адрес"), "judicial_district": field(t, "Судебный район"),
            "municipal": field(t, "Муниципальный район")}


def slug(s: str) -> str:
    tr = dict(zip("абвгдеёжзийклмнопрстуфхцчшщъыьэюя", "a b v g d e e zh z i j k l m n o p r s t u f h c ch sh sch _ y _ e yu ya".split()))
    return re.sub(r"[^a-z0-9]+", "-", "".join(tr.get(c, c) for c in s.lower())).strip("-")


def find_parent(p: dict, courts: list[dict]) -> dict | None:
    """Районный суд участка: по названию района («Кировского района г. Перми» -> Кировский) и городу."""
    m = re.search(r"№ \d+ (\S+) района", p["name"])
    if not m:
        return None
    stem = m.group(1).lower()[:-3]
    cands = [c for c in courts if c["level"] == "district" and c["judicial_district"].lower().startswith(stem)]
    if "Перми" in p["name"]:
        cands = [c for c in cands if "Перми" in c["name"]] or cands
    else:
        cands = [c for c in cands if "Перми" not in c["name"]] or cands
    return cands[0] if cands else None


def main(cache: str, out: str, manual: str | None = None):
    cache_p = Path(cache)
    courts, by_district = [], {}
    for f in sorted(cache_p.glob("*.html")):
        if "sudebnyj-uchastok" in f.name:
            continue
        c = parse_court(f)
        if c and ("суд" in c["name"].lower()) and f.name.count("_") == 2:
            cid = "d-" + slug(c["name"])
            courts.append({"id": cid, "level": "district", **{k: v for k, v in c.items() if k != "municipal"},
                           "territory": []})
            by_district.setdefault(c["judicial_district"], []).append(courts[-1])
    if manual:
        for m in (yaml.safe_load(Path(manual).read_text(encoding="utf-8")) or {}).get("courts", []):
            courts.append(m)
            if m.get("level") == "district":
                by_district.setdefault(m.get("judicial_district", ""), []).append(m)
    n_ms = n_struct = 0
    for f in sorted(cache_p.glob("*sudebnyj-uchastok*.html")):
        p = parse_participant(f)
        if not p:
            print("пропуск", f.name, file=sys.stderr)
            continue
        parent = find_parent(p, courts)
        entry = {"id": "ms-" + slug(p["name"]), "level": "magistrate", "name": p["name"], "address": p["address"],
                 "judge": p["judge"], "parent": parent["id"] if parent else None,
                 "judicial_district": p["judicial_district"],
                 "territory_text": p["text"],
                 "territory": [{"city": p["city"], "streets": p["streets"]}] if p["streets"] else []}
        courts.append(entry)
        n_ms += 1
        n_struct += bool(p["streets"])
    Path(out).write_text(
        "# Справочник судов и подсудности. Собран tools/build_courts.py из sudinforf.ru (агрегатор, не закон).\n"
        "# Перед подачей сверяйте границы участков с законом субъекта (для Пермского края - Закон 453-ПК и изменения).\n"
        + yaml.safe_dump({"region_name": "Пермский край", "default_city": "Пермь", "courts": courts}, allow_unicode=True, sort_keys=False, width=200),
        encoding="utf-8")
    print(f"судов: {len(courts) - n_ms}, участков: {n_ms}, из них со списком улиц: {n_struct}")


if __name__ == "__main__":
    main(*sys.argv[1:4])
