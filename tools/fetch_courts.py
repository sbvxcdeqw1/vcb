"""Скачивает страницы судов и судебных участков региона (sudinforf.ru) в локальный кэш.

    python tools/fetch_courts.py perm sudy-v-permskom-krae КЭШ_ПАПКА

Затем: python tools/build_courts.py КЭШ_ПАПКА data/courts/perm_krai.yaml
Источник - справочник-агрегатор, не официальный закон: перед подачей сверяйте с законом субъекта о границах участков.
"""
from __future__ import annotations

import re
import sys
import time
import urllib.request
from pathlib import Path

BASE = "https://sudinforf.ru"
UA = {"User-Agent": "Mozilla/5.0"}


def get(url: str, cache: Path) -> str:
    f = cache / (re.sub(r"[^\w\-]+", "_", url.replace(BASE, "")) + ".html")
    if f.exists():
        return f.read_text(encoding="utf-8")
    for attempt in range(4):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=40) as r:
                html = r.read().decode("utf-8", "ignore")
            f.write_text(html, encoding="utf-8")
            time.sleep(0.6)
            return html
        except Exception as e:  # noqa: BLE001
            print("retry", url, e, file=sys.stderr)
            time.sleep(2 ** attempt)
    raise RuntimeError(f"не удалось скачать {url}")


def links(html: str, prefix: str) -> list[str]:
    out = []
    for h in re.findall(r'href="(https://sudinforf\.ru' + re.escape(prefix) + r'[^"#?]*)"', html):
        if h not in out:
            out.append(h)
    return out


def main(region: str, index_slug: str, cache_dir: str):
    cache = Path(cache_dir)
    cache.mkdir(parents=True, exist_ok=True)
    index = get(f"{BASE}/{index_slug}", cache)
    courts = [u for u in links(index, f"/{region}/") if u.count("/") == 4]
    for c in courts:
        page = get(c, cache)
        for u in links(page, f"/{region}/"):
            if "sudebnyj-uchastok" in u:
                get(u, cache)
    print("страниц в кэше:", len(list(cache.glob("*.html"))))


if __name__ == "__main__":
    main(*sys.argv[1:4])
