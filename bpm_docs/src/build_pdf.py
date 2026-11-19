# -*- coding: utf-8 -*-
import re, math, html
from data import *

W, H = 1920, 1080
esc = html.escape
SRC_COL = {"Пр": "#2f6f4f", "З": "#6b3fa0", "Н": "#b8620a", "—": "#7a8594"}
SRC_BG  = {"Пр": "#e3f2ea", "З": "#efe6f8", "Н": "#fdeedb", "—": "#eef0f3"}

CSS = """
@page { size: 1920px 1080px; margin:0 }
*{box-sizing:border-box}
html,body{margin:0;padding:0;font-family:'Liberation Sans','DejaVu Sans',Arial,sans-serif;color:#16263d;background:#fff}
.slide{width:1920px;height:1080px;position:relative;overflow:hidden;page-break-after:always;background:#f6f8fc;padding:34px 40px 0 40px}
.slide h1{margin:0;font-size:38px;line-height:1.1;color:#14304f;font-weight:700}
.slide .sub{margin-top:8px;font-size:18px;color:#5b6b80}
.rule{height:4px;background:#d9e0ea;margin:12px 0 14px 0;border-radius:2px}
.foot{position:absolute;left:40px;right:40px;bottom:12px;font-size:13px;color:#7a8594;display:flex;justify-content:space-between}
.chip{display:inline-block;padding:1px 7px;border-radius:9px;font-size:12px;font-weight:700;white-space:nowrap}
.lane{position:absolute;left:0;right:0;border-top:1px solid #cfd8e4}
.lanelabel{position:absolute;left:0;width:176px;padding:6px 10px;font-size:15px;font-weight:700;color:#14304f;background:#e6ecf5;border-right:2px solid #b6c3d6;display:flex;align-items:center}
.box{position:absolute;border:2px solid #2c6aa0;border-radius:10px;background:#fff;padding:17px 4px 8px 4px;font-size:12.5px;line-height:1.15;text-align:center;display:flex;align-items:center;justify-content:center;overflow:hidden}
.box .n{position:absolute;left:4px;top:2px;font-size:11px;font-weight:700;color:#fff;background:#2c6aa0;border-radius:8px;padding:0 5px}
.box .dl{position:absolute;left:4px;right:4px;bottom:2px;font-size:11px;font-weight:700;border-radius:6px;padding:0 4px;line-height:15px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.box.sm{font-size:11.5px;padding:15px 3px 5px 3px;line-height:1.08}
.box.client{border-color:#7a4fa8;background:#f4edfb}.box.client .n{background:#7a4fa8}
.box.end{border-color:#2f7a4f;background:#e3f2ea;border-radius:22px;font-weight:700;color:#1f5a38}.box.end .n{background:#2f7a4f}
.box.dec{border:none;background:#d98a1f;clip-path:polygon(9% 0,91% 0,100% 50%,91% 100%,9% 100%,0 50%);padding:0}
.box.dec .in{position:absolute;inset:3px;background:#fff3de;clip-path:polygon(9% 0,91% 0,100% 50%,91% 100%,9% 100%,0 50%);display:flex;align-items:center;justify-content:center;padding:4px 16px;font-weight:700;color:#7a470a;font-size:12.5px;line-height:1.15}
.elabel{position:absolute;transform:translate(-50%,-50%);background:#f6f8fc;font-size:12px;font-weight:700;color:#8a4b00;padding:0 4px;border-radius:5px;white-space:nowrap}
table.t{border-collapse:collapse;width:100%;font-size:15px;background:#fff}
table.t th{background:#14304f;color:#fff;text-align:left;padding:7px 9px;font-size:14px}
table.t td{border:1px solid #cdd6e2;padding:6px 9px;vertical-align:top;line-height:1.25}
table.t tr:nth-child(even) td{background:#f4f7fb}
.tight td{padding:3px 8px;line-height:1.2}
.card{background:#fff;border:2px solid #2c6aa0;border-radius:12px;padding:10px 14px;font-size:15px;line-height:1.3}
.card h3{margin:0 0 5px 0;font-size:17px;color:#14304f}
.scalelabel{position:absolute;left:0;width:176px;padding:2px 10px;font-size:14px;font-weight:700;color:#44566d;display:flex;align-items:center;background:#eef2f8;border-right:2px solid #b6c3d6}
.scalebg{position:absolute;background:#eef2f8}
.span{position:absolute;border:2px solid;border-radius:6px;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding:0 4px}
.legend{position:absolute;left:40px;right:40px;bottom:36px;font-size:14px;color:#44566d;display:flex;gap:22px;align-items:center}
.sw{display:inline-block;width:14px;height:14px;border-radius:4px;vertical-align:-2px;margin-right:5px}
"""

def lane_name(k): return LANES[k]

def short_dl(dl, src):
    if not dl or dl == "—": return ""
    if dl.lower().startswith("вне"): return "вне 5 р.д."
    m = re.search(r"(\d+)\s*р\.д\.", dl)
    if m and not dl.lower().startswith(("срок задаёт","срок указан")):
        t = f"{m.group(1)} р.д."
    elif dl.lower().startswith("постоянно"): t = "постоянно"
    elif dl.lower().startswith("ежеквартально"): t = "ежеквартально"
    elif src == "З": t = "по закону"
    elif dl.lower().startswith(("до ","не позднее","в назнач","по срок","по граф","по учёт","срок")): t = "по сроку"
    else: t = ""
    return t

# ------------------------------------------------------------------ шкала сроков
def _skip_dl(dl):
    return (not dl) or dl.strip() in ("—", "-") or dl.startswith("Не определён")

def build_spans(p, col):
    """Скобки сроков: явные p['spans'] = [(метка, шаг_с, шаг_по, источник)] + подряд идущие шаги с одинаковым сроком."""
    steps = p["steps"]; order = {s["id"]: i for i, s in enumerate(steps)}
    out = []; covered = set()
    for t in p.get("spans", []):
        lab, a, bb, src = t[:4]; o = t[4] if len(t) > 4 else {}
        ids = [s["id"] for s in steps[order[a]:order[bb] + 1] if s["id"] not in o.get("exclude", [])]
        if o.get("cover", True): covered.update(ids)
        out.append(dict(label=lab, src=src, c1=min(col[i] for i in ids), c2=max(col[i] for i in ids)))
    group = None
    def flush():
        nonlocal group
        if group:
            out.append(dict(label=group["label"], src=group["src"], c1=min(group["cols"]), c2=max(group["cols"])))
        group = None
    for s in steps:
        if s["kind"] == "decision" or s["id"] in covered or _skip_dl(s["dl"]):
            flush(); continue
        key = (s["dl"], s["src"])
        if group and group["key"] == key: group["cols"].append(col[s["id"]])
        else:
            flush(); group = dict(key=key, label=s["dl"], src=s["src"], cols=[col[s["id"]]])
    flush()
    seen = set(); uniq = []
    for d in out:
        k = (d["c1"], d["c2"], d["src"], short_dl(d["label"], d["src"]) or d["label"])
        if k in seen: continue
        seen.add(k); uniq.append(d)
    return sorted(uniq, key=lambda d: (d["c1"], -d["c2"]))

def assign_levels(spans):
    ends = []; levels = []
    for sp in spans:
        for i, e in enumerate(ends):
            if e < sp["c1"]:
                ends[i] = sp["c2"]; levels.append(i); break
        else:
            ends.append(sp["c2"]); levels.append(len(ends) - 1)
    return levels

# ------------------------------------------------------------------ swimlane
def diagram(p):
    lanes = p["lanes"]; steps = p["steps"]
    nl = len(lanes)
    # columns
    col = {}; last = 0
    for s in steps:
        c = s["c"] if s["c"] else last + 1
        col[s["id"]] = c; last = c
    ncols = max(col.values())
    spans = build_spans(p, col)
    levels = assign_levels(spans)
    SH = 26
    so = (max(levels) + 1) * SH + 8 if spans else 0
    X0, LW = 0, 176
    area_w = W - 80
    top = 0
    avail_h = 800 - so
    lh = min(150, avail_h // nl)
    left = LW + 8
    colw = (area_w - left) / ncols
    gap = 18
    bh = min(lh - 22, 118)
    byid = {}
    for s in steps:
        li = lanes.index(s["lane"]); ci = col[s["id"]]
        x = left + (ci - 1) * colw + gap / 2; w = colw - gap
        y = so + li * lh + (lh - bh) / 2
        byid[s["id"]] = dict(x=x, y=y, w=w, h=bh, s=s, ci=ci, li=li)
    out = [f'<div style="position:relative;width:{area_w}px;height:{so+nl*lh+70}px">']
    for i, l in enumerate(lanes):
        out.append(f'<div class="lane" style="top:{so+i*lh}px;height:{lh}px;background:{"#ffffff" if i%2==0 else "#f1f5fa"}"></div>')
        out.append(f'<div class="lanelabel" style="top:{so+i*lh}px;height:{lh}px">{esc(lane_name(l))}</div>')
    if spans:
        out.append(f'<div class="scalelabel" style="top:0;height:{so-6}px">Шкала сроков</div>')
        out.append(f'<div class="scalebg" style="top:0;height:{so-6}px;left:{left-8}px;right:0"></div>')
        for sp, lv in zip(spans, levels):
            x1 = left + (sp["c1"] - 1) * colw + gap / 2; x2 = left + sp["c2"] * colw - gap / 2
            sw = x2 - x1
            lab = sp["label"] if len(sp["label"]) * 7 <= sw - 10 else (short_dl(sp["label"], sp["src"]) or sp["label"])
            out.append(f'<div class="span" title="{esc(sp["label"])}" style="left:{x1:.0f}px;top:{lv*SH+2}px;width:{sw:.0f}px;height:{SH-5}px;'
                       f'background:{SRC_BG[sp["src"]]};border-color:{SRC_COL[sp["src"]]};color:{SRC_COL[sp["src"]]}">{esc(lab)}</div>')
    # edges
    edges = []
    seen = set()
    skip = set(p.get("seq_skip", []))
    for e in p.get("edges", []):
        a, b = e[0], e[1]; lab = e[2] if len(e) > 2 else None; sty = e[3] if len(e) > 3 else None
        if lab == "dash": lab, sty = None, "dash"
        edges.append((a, b, lab, sty)); seen.add((a, b))
    if p.get("auto", True):
        for i in range(len(steps) - 1):
            a, b = steps[i]["id"], steps[i + 1]["id"]
            if (a, b) in seen or (a, b) in skip: continue
            if steps[i]["kind"] == "end": continue
            edges.append((a, b, None, None)); seen.add((a, b))
    svg = [f'<svg width="{area_w}" height="{so+nl*lh+70}" style="position:absolute;left:0;top:0" xmlns="http://www.w3.org/2000/svg">',
           '<defs><marker id="ar" markerUnits="userSpaceOnUse" markerWidth="10" markerHeight="10" refX="9" refY="5" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="#52647c"/></marker>'
           '<marker id="ar2" markerUnits="userSpaceOnUse" markerWidth="10" markerHeight="10" refX="9" refY="5" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="#b8620a"/></marker></defs>']
    labels = []
    loopn = 0
    bottom = so + nl * lh
    for a, b, lab, sty in edges:
        A, B = byid[a], byid[b]
        ax, ay = A["x"] + A["w"], A["y"] + A["h"] / 2
        bx, by = B["x"], B["y"] + B["h"] / 2
        acx, bcx = A["x"] + A["w"] / 2, B["x"] + B["w"] / 2
        color, mk = "#52647c", "ar"
        dash = ""
        if sty in ("loop",) or B["ci"] < A["ci"]:
            cy = bottom + 16 + 15 * loopn; loopn += 1
            pts = [(acx, A["y"] + A["h"]), (acx, cy), (bcx, cy), (bcx, B["y"] + B["h"])]
            color, mk, dash = "#b8620a", "ar2", ' stroke-dasharray="6 4"'
            lp = ((acx + bcx) / 2, cy)
        elif sty == "ul":
            sy = A["y"] if B["li"] < A["li"] else A["y"] + A["h"]
            pts = [(acx, sy), (acx, by), (B["x"] + B["w"], by)]
            color, mk, dash = "#b8620a", "ar2", ' stroke-dasharray="6 4"'
            lp = ((acx + B["x"] + B["w"]) / 2, by - 11)
        elif sty == "dr":
            pts = [(acx, A["y"] + A["h"]), (acx, by), (bx, by)]
            lp = (acx, (A["y"] + A["h"] + by) / 2)
            lp = (acx, A["y"] + A["h"] + 12)
        elif A["ci"] == B["ci"]:
            if B["li"] > A["li"]: pts = [(acx, A["y"] + A["h"]), (acx, B["y"])]
            else: pts = [(acx, A["y"]), (acx, B["y"] + B["h"])]
            lp = (acx + 0, (pts[0][1] + pts[1][1]) / 2)
        else:
            if abs(ay - by) < 1: pts = [(ax, ay), (bx, by)]; lp = ((ax + bx) / 2, ay - 11)
            else:
                vx = bx - gap / 2
                pts = [(ax, ay), (vx, ay), (vx, by), (bx, by)]
                if vx - ax > 60: lp = ((ax + vx) / 2, ay - 11)
                else:
                    y1, y2 = sorted((ay, by))
                    bounds = [so + k * lh for k in range(nl + 1) if y1 < so + k * lh < y2]
                    mid = (y1 + y2) / 2
                    ly = min(bounds, key=lambda v: abs(v - mid)) if bounds else mid
                    lp = (vx, ly)
        if sty == "dash": dash = ' stroke-dasharray="4 4"'; color = "#7a8594"
        d = "M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in pts)
        svg.append(f'<path d="{d}" fill="none" stroke="{color}" stroke-width="2"{dash} marker-end="url(#{mk})"/>')
        if lab: labels.append(f'<div class="elabel" style="left:{lp[0]:.0f}px;top:{lp[1]:.0f}px">{esc(lab)}</div>')
    svg.append("</svg>")
    out.append("".join(svg))
    for sid, r in byid.items():
        s = r["s"]
        st = f'left:{r["x"]:.0f}px;top:{r["y"]:.0f}px;width:{r["w"]:.0f}px;height:{r["h"]:.0f}px'
        if s["kind"] == "decision":
            out.append(f'<div class="box dec" style="{st}"><div class="in">{esc(s["title"])}</div></div>')
        elif s["kind"] == "end":
            out.append(f'<div class="box end" style="{st};padding:4px"><span class="n">{esc(s["id"])}</span>{esc(s["title"].replace("ФИНИШ: ","ФИНИШ: "))}</div>')
        else:
            cls = ("box client" if s["lane"] == "client" else "box") + (" sm" if bh < 90 else "")
            sd = short_dl(s["dl"], s["src"])
            chip = ""
            out.append(f'<div class="{cls}" style="{st}"><span class="n">{esc(s["id"])}</span><span>{esc(s["title"])}</span>{chip}</div>')
    out.extend(labels)
    out.append("</div>")
    return "".join(out)

def legend():
    return ('<div class="legend">'
            '<span><span class="sw" style="background:#fff;border:2px solid #2c6aa0"></span>шаг</span>'
            '<span><span class="sw" style="background:#fff3de;border:2px solid #d98a1f"></span>решение</span>'
            '<span><span class="sw" style="background:#e3f2ea;border:2px solid #2f7a4f"></span>финиш</span>'
            '<span><span class="sw" style="background:#f4edfb;border:2px solid #7a4fa8"></span>действие клиента</span>'
            '<span style="color:#b8620a">- - ► возврат / повторный запрос</span>'
            '<span>Шкала сроков над схемой: скобка охватывает один или несколько шагов; цвет — источник срока: <span class="chip" style="background:#e3f2ea;color:#2f6f4f">Пр</span> из презентации &nbsp;'
            '<span class="chip" style="background:#efe6f8;color:#6b3fa0">З</span> по закону РФ &nbsp;'
            '<span class="chip" style="background:#fdeedb;color:#b8620a">Н</span> предложено, требует утверждения</span>'
            '</div>')

# ------------------------------------------------------------------ helpers
pageno = [0]
def slide(title, sub, body, extra=""):
    pageno[0] += 1
    return (f'<section class="slide"><h1>{title}</h1><div class="sub">{sub}</div><div class="rule"></div>{body}{extra}'
            f'<div class="foot"><span>Рабочая версия 17.11.2026 · формы документов не утверждены · сроки «Н» — предложения до утверждения</span><span>{pageno[0]}</span></div></section>')

def chip(src): return f'<span class="chip" style="background:{SRC_BG[src]};color:{SRC_COL[src]}">{src}</span>'

def ul(items, sep="; "):
    return esc(sep.join(items)) if items else "—"

def est_lines(txt, width_chars): return max(1, math.ceil(len(txt) / width_chars))

def note(s):
    return f'<div style="color:#5b6b80;font-size:13px;margin-top:3px">{esc(s["act"])}</div>' if s['act'] != s['title'] else ''

def proc_table_pages(p):
    rows = p["steps"]
    pages = []; cur = []; h = 0
    limit_first = 790; limit = 900
    def rh(s):
        n = max(est_lines(s["title"], 52) + (est_lines(s["act"], 60) if s["act"] != s["title"] else 0), est_lines("; ".join(s["inp"]) or "—", 30), est_lines("; ".join(s["out"]) or "—", 30), est_lines(s["dl"] or "—", 24), 1)
        return n * 19 + 16
    for s in rows:
        r = rh(s)
        lim = limit_first if not pages else limit
        if h + r > lim and cur:
            pages.append(cur); cur = []; h = 0
        cur.append(s); h += r
    if cur: pages.append(cur)
    return pages

def proc_inputs_outputs(p):
    produced = set(); ins = []; outs = []
    for s in p["steps"]:
        for d in s["out"]:
            if d not in outs: outs.append(d)
    pr = set(outs)
    for s in p["steps"]:
        for d in s["inp"]:
            if d not in pr and d not in ins: ins.append(d)
    return ins, outs

def process_slides(p):
    res = []
    dg = diagram(p)
    sub = f'Старт: {esc(p["start"])} → Финиш: {esc(p["finish"])}'
    if p.get("draft"):
        sub += ' &nbsp;<span class="chip" style="background:#fde2e2;color:#9a2a2a;font-size:14px">ЗАГОТОВКА — функции не согласованы, показаны только согласованные точки входа/выхода</span>'
    res.append(slide(f'{p["id"]}. {esc(p["name"])}', sub, dg, legend()))
    pages = proc_table_pages(p)
    ins, outs = proc_inputs_outputs(p)
    for i, pg in enumerate(pages):
        strip = ""
        if i == 0:
            strip = ('<div style="display:flex;gap:16px;margin-bottom:12px">'
                     f'<div class="card" style="flex:1;border-color:#7a4fa8"><h3>Для начала процесса требуется</h3>{esc("; ".join(ins)) if ins else "Достаточно факта старта (см. выше); входных документов за пределами процесса нет"}</div>'
                     f'<div class="card" style="flex:1.3;border-color:#2f7a4f"><h3>Продукты процесса (документы и результаты)</h3>{esc("; ".join(outs)) if outs else "—"}</div></div>')
        trs = ""
        for s in pg:
            trs += (f'<tr><td><b>{esc(s["id"])}</b></td><td>{esc(lane_name(s["lane"]))}</td><td>{esc(s["title"])}{note(s)}</td>'
                    f'<td>{ul(s["inp"])}</td><td>{ul(s["out"])}</td><td>{esc(s["dl"] or "—")}</td><td>{chip(s["src"])}</td></tr>')
        tb = (f'{strip}<table class="t"><tr><th style="width:56px">Шаг</th><th style="width:190px">Роль</th><th>Действие</th>'
              f'<th style="width:300px">Вход (что нужно)</th><th style="width:300px">Выход (продукт)</th><th style="width:230px">Срок</th><th style="width:56px">Ист.</th></tr>{trs}</table>')
        suffix = f' (стр. {i+1}/{len(pages)})' if len(pages) > 1 else ''
        res.append(slide(f'{p["id"]}. Входы, действия, продукты и сроки{suffix}', esc(p["name"]), tb))
    return res

# ------------------------------------------------------------------ static slides
def pid_name(i):
    p = next(p for p in P if p["id"] == i)
    return f'<b>{esc(p["id"])}</b> {esc(p["name"])}' + (' <i style="color:#9a2a2a">(заготовка)</i>' if p.get("draft") else '')

def cover():
    letters = "".join(f'<div style="margin:3px 0"><b style="font-size:20px;color:#14304f">{a}</b> — {esc(b)}: {esc(c)}</div>' for a,b,c in CLS_LETTERS)
    body = ('<div style="font-size:22px;margin:14px 0 22px 0;color:#2b3f58">Оргструктура · классификация процессов · основные и поддерживающие бизнес-процессы · регламентные сроки · документы</div>'
            '<div style="display:flex;gap:16px;align-items:stretch">'
            f'<div class="card" style="flex:1.1"><h3>Классификация и идентификаторы процессов</h3>{letters}<div style="margin-top:8px;color:#44566d">Коды можно сочетать: ОС1 — основной сквозной, ПЛ6 — поддерживающий локальный. Номер в коде — порядковый номер процесса; дочерний процесс — через точку (ОЛ2.1).</div></div>'
            '<div class="card" style="flex:1"><h3>Процессы в текущем наборе</h3>' + "".join(f'<div style="margin:3px 0">{pid_name(p["id"])}</div>' for p in P) + '</div>'
            '</div>'
            '<div style="display:flex;gap:16px;margin-top:18px">'
            '<div class="card" style="flex:1"><h3>Состав пакета</h3>1. Классификация, реестр процессов и карта<br>2. Оргструктура и роли<br>3. Процессы: схема (swimlane) + таблица «вход → действие → продукт → срок»<br>4. Регламентные сроки: из презентации, по закону РФ, предложенные<br>5. Реестр документов (входы и продукты)<br>6. Допущения, открытые вопросы, процессы в резерве<br>Дополнительно: Excel-матрица с теми же данными</div>'
            '<div class="card" style="flex:1"><h3>Условные обозначения</h3>Роли в схемах: Руководитель направления, Начальник отдела, Специалист (ведущий / 1 / 2 категории).<br>'
            f'Источник срока: {chip("Пр")} презентация (согласовано) · {chip("З")} законодательство РФ · {chip("Н")} предложено, требует утверждения.<br>р.д. — рабочие дни по производственному календарю РФ.</div>'
            '<div class="card" style="flex:1;border-color:#b8620a"><h3>Важно</h3>«Юридическое сопровождение» и «Привлечение клиента» — заготовки: показаны только согласованные точки входа/выхода, внутренние функции не домысливались. Формы документов не утверждены. Должностные инструкции и чек-листы — следующий этап.</div></div>')
    return slide("Документация бизнес-процессов бухгалтерской организации", "Рабочая версия, редакция 2 (с классификацией процессов) · 17.11.2026", body)

def classif():
    letters = "".join(f'<tr><td style="font-size:22px;font-weight:700;text-align:center;width:50px">{a}</td><td><b>{esc(b)}</b></td><td>{esc(c)}</td></tr>' for a,b,c in CLS_LETTERS)
    t1 = f'<table class="t" style="font-size:15px"><tr><th>Код</th><th>Тип</th><th>Смысл</th></tr>{letters}</table>'
    rows = ""
    for p in P:
        st = esc(p["status"])
        col = "#9a2a2a" if p.get("draft") else "#2f6f4f"
        rows += (f'<tr><td><b>{esc(p["id"])}</b></td><td>{esc(p["name"])}</td><td>{esc(p["cls"])}</td><td>{esc(p["owner"])}</td>'
                 f'<td style="color:{col};font-weight:700">{st}</td></tr>')
    t2 = ('<table class="t" style="font-size:14px"><tr><th style="width:80px">ID</th><th style="width:330px">Процесс</th><th style="width:60px">Класс</th>'
          '<th style="width:430px">Владелец / участвуют</th><th style="width:260px">Статус</th></tr>' + rows + '</table>')
    body = f'<div style="display:flex;gap:18px;align-items:flex-start"><div style="flex:0 0 560px">{t1}<div class="card" style="margin-top:14px;font-size:14px">Код составной: первая буква — роль процесса в бизнесе (О, П, У, И), вторая — охват (С, Л). Пример: <b>ОС1</b> — основной сквозной. Дочерние процессы нумеруются через точку: <b>ОЛ2.1</b>.</div></div><div style="flex:1">{t2}</div></div>'
    return slide("Классификация процессов и реестр", "Идентификаторы присваиваются отдельно или вместе: ОС, ОЛ, ПС, ПЛ, УС, УЛ, ИС, ИЛ", body)

def procmap():
    def box(t, items, color="#2c6aa0", flex=1, dashed=False):
        li = "".join(f'<div style="margin:3px 0">{i}</div>' for i in items)
        return f'<div class="card" style="flex:{flex};border-color:{color};{"border-style:dashed;" if dashed else ""}"><h3 style="color:{color}">{t}</h3>{li}</div>'
    stages = [("0. Привлечение","ОЛ4"),("1. Приём и запуск","ОС1"),("2. Обслуживание","ОС2 → ОЛ2.1–ОЛ2.5"),("3. Изменение состава услуг","ПС5"),("4. Прекращение","ОС3")]
    st = "".join(f'<div class="card" style="flex:1;text-align:center;border-color:#7a4fa8"><h3 style="color:#7a4fa8">{a}</h3>{b}</div>' + ('<div style="align-self:center;font-size:30px;color:#52647c;padding:0 6px">►</div>' if i < 4 else '') for i,(a,b) in enumerate(stages))
    body = ('<div style="font-size:18px;font-weight:700;color:#14304f;margin-bottom:6px">ЖИЗНЕННЫЙ ЦИКЛ КЛИЕНТА</div>'
            f'<div style="display:flex;gap:6px;margin-bottom:18px">{st}</div>'
            '<div style="display:flex;gap:16px;margin-bottom:16px">'
            + box("Основные сквозные (ОС)", [pid_name("ОС1"), pid_name("ОС2"), pid_name("ОС3")], "#7a4fa8")
            + box("Основные локальные (ОЛ)", [pid_name("ОЛ2.1"), pid_name("ОЛ2.2"), pid_name("ОЛ2.3"), pid_name("ОЛ2.4"), pid_name("ОЛ2.5"), pid_name("ОЛ4")], "#2f7a4f")
            + box("Поддерживающие (ПС, ПЛ)", [pid_name("ПС5"), pid_name("ПЛ6"), pid_name("ПЛ7")], "#2c6aa0") + '</div>'
            '<div style="display:flex;gap:16px;margin-bottom:16px">'
            + box("Управленческие (У) и инновационные (И)", ["Пока не определены — будут добавлены позже"], "#7a8594", 1, True)
            + box("Процессы в резерве (разработаны ранее, ждут классификации)", ["Внутренний запрос между специалистами · Разногласия и эскалация", "Ошибки · Претензии · Контроль качества · Трудовая дисциплина", "<i>Предлагаемый класс — У (С/Л по охвату); решение за вами</i>"], "#b8620a", 1.6, True) + '</div>'
            '<div class="card" style="border-color:#7a4fa8;font-size:16px"><b>Координатор клиента</b> — связующее звено в общении с клиентом: <b>данные клиенту и от клиента передаёт и получает только он</b> (исключения: отдел судебно-претензионной работы общается с клиентом напрямую; Главный бухгалтер направляет клиенту счета и акты, ПЛ7). Назначается для каждого клиента: при одной услуге — специалист этого направления, при нескольких — специалист направления с наивысшим приоритетом (бухгалтерия → РЦ → ГИС ЖКХ → юристы), назначает руководитель этого направления.</div>')
    return slide("Карта процессов", "Как процессы связаны с жизненным циклом клиента", body)

def orgchart():
    def role(t, s="", dashed=False, col="#2c6aa0", bg="#fff"):
        return f'<div class="card" style="border-color:{col};background:{bg};{"border-style:dashed;" if dashed else ""}padding:8px 10px;font-size:14px;text-align:center"><b style="color:{col};font-size:15px">{t}</b><br>{s}</div>'
    def dept(title, note=""):
        return ('<div style="border:2px solid #2c6aa0;border-radius:10px;background:#fff;padding:6px;margin-top:8px">'
                f'<div style="text-align:center;font-weight:700;color:#14304f;font-size:14px">{title}</div>'
                '<div class="card" style="margin-top:5px;padding:5px;font-size:13px;text-align:center;border-color:#9bb3cf">Начальник отдела<br>«…»</div>'
                '<div class="card" style="margin-top:5px;padding:5px;font-size:13px;text-align:center;border-color:#2f7a4f;background:#e3f2ea">Специалисты:<br>ведущий / 1 категории / 2 категории</div>'
                f'{note}</div>')
    dirs = [
      ("Бухгалтерия","Руководитель бухгалтерии<br>(Главный бухгалтер)","#2c6aa0",False),
      ("Расчётный центр","Руководитель РЦ","#2c6aa0",False),
      ("ГИС ЖКХ","Руководитель направления ГИС ЖКХ","#2c6aa0",False),
      ("Юридическое","Руководитель юридического направления","#5b6b80",False),
      ("Продажи","Руководитель / ответственное лицо продаж","#7a4fa8",True),
      ("Техническое","Руководитель / ответственное лицо техподдержки","#5b6b80",True)]
    cells = ""
    for i,(n,r,c,d) in enumerate(dirs):
        if n == "Юридическое":
            inner = role("Юридическое сопровождение", r, False, c, "#e9edf2") + '<div style="display:flex;gap:6px">' + '<div style="flex:1">'+dept("Отдел сопровождения взыскания задолженностей", '<div style="font-size:12px;color:#7a4fa8;text-align:center;margin-top:4px;font-weight:700">взыскание ЖКУ, в т. ч. через суд · «единое окно»: с клиентом — через координатора</div>')+'</div><div style="flex:1">'+dept("Отдел судебно-претензионной работы", '<div style="font-size:12px;color:#b8620a;text-align:center;margin-top:4px;font-weight:700">с клиентом — напрямую · иски на контроль после подачи в суд</div>')+'</div></div>'
            cells += f'<div style="flex:2.1;border:2px solid #5b6b80;border-radius:14px;padding:8px;background:#f3f5f8">{inner}</div>'
        elif n == "Бухгалтерия":
            inner = role(n, r, False, c, "#dbe9f7") + '<div style="display:flex;gap:6px">' + '<div style="flex:1">'+dept("Отдел ЖКХ")+'</div><div style="flex:1">'+dept("Отдел коммерческих клиентов")+'</div></div>'
            cells += f'<div style="flex:2.1;border:2px solid #2c6aa0;border-radius:14px;padding:8px;background:#eef4fb">{inner}</div>'
        else:
            if n in ("Расчётный центр","ГИС ЖКХ"):
                body = dept("Отделы", '<div style="font-size:12px;color:#5b6b80;text-align:center;margin-top:4px">число отделов определяется после оценки нагрузки</div>')
            else:
                body = ('<div style="border:2px dashed #9aa6b6;border-radius:10px;padding:8px;margin-top:8px;text-align:center;font-size:13px;color:#5b6b80;background:#fff">Структура уточняется<br>Начальник отдела «…» / Специалисты</div>')
            cells += f'<div style="flex:1;border:2px {"dashed" if d else "solid"} {c};border-radius:14px;padding:8px;background:{"#f3f5f8" if d else "#eef4fb"}">{role(n, r, False, c, "#ffffff")}{body}</div>'
    body = ('<div style="display:flex;flex-direction:column;align-items:center;gap:6px">'
            '<div style="background:#14304f;color:#fff;border-radius:40px;padding:8px 60px;font-weight:700;font-size:18px">СОБСТВЕННИКИ <span style="font-weight:400;font-size:14px">— стратегическое управление</span></div>'
            '<div style="width:3px;height:14px;background:#52647c"></div>'
            '<div style="background:#14304f;color:#fff;border-radius:10px;padding:8px 40px;text-align:center"><b style="font-size:19px">ДИРЕКТОР</b><br><span style="font-size:14px">общие правила · KPI-принципы · ресурсы · кадровые решения · проверки</span></div>'
            '<div style="width:3px;height:14px;background:#52647c"></div></div>'
            '<div style="border-top:3px solid #52647c;margin:0 4% 8px 4%"></div>'
            f'<div style="display:flex;gap:12px;align-items:flex-start">{cells}</div>'
            '<div style="display:flex;gap:14px;margin-top:14px">'
            '<div class="card" style="flex:1.2;border-color:#5b6b80"><h3>Именование должностей в схемах и документах</h3>Руководитель &lt;наименование направления&gt; · Начальник отдела &lt;наименование отдела&gt; · Специалист &lt;ведущий / 1 категории / 2 категории&gt;. Численность отделов определяется после оценки нагрузки и здесь не фиксируется. Пунктир — структура не определена.</div>'
            '<div class="card" style="flex:1;border-color:#7a4fa8"><h3>Координатор клиента</h3>Роль, а не подразделение. Единственный канал данных с клиентом (кроме отдела судебно-претензионной работы и направления счетов/актов Главным бухгалтером). Один координатор на клиента; при нескольких услугах — по приоритету направлений.</div>'
            '<div class="card" style="flex:1;border-color:#b8620a"><h3>Сквозные контуры</h3>Общие процессы взаимодействия · единый контур качества · управление персоналом (дисциплина, нагрузка, обучение) — действуют во всех направлениях.</div></div>')
    return slide("Оргструктура и управленческие роли", "Собственники → Директор → руководители направлений → начальники отделов → специалисты", body)

def roles():
    rows = [
     ("Директор","Общие принципы и правила, KPI-принципы, ресурсы и кадровые решения; вправе проверить любое направление без графика; окончательные решения по серьёзным дисциплинарным мерам и увольнению; подключается к критическим претензиям и неурегулированным разногласиям","Решение по разногласию; решение по претензии; квартальная сводка; решения по увольнению; решения по KPI","2 р.д. (с руководителями); совещание раз в 2 недели"),
     ("Руководитель направления","Методология, общий контроль, ресурсы, эскалации; назначает координатора и ответственных за ошибки; готовит ответы по претензиям; выборочные проверки; предлагает директору взыскание/увольнение","Подтверждение готовности направления; назначение ответственного за ошибку; план корректирующих мер; результаты проверок","2 р.д. (между руководителями)"),
     ("Начальник отдела","Текущая организация и контроль работы, назначает специалистов и задачи, классифицирует ошибки и претензии, первичный контроль дисциплины, подтверждает готовность отдела","Назначение специалиста и задач; категория ошибки; подтверждение готовности отдела; меры по дисциплине","1 р.д. внутри отдела"),
     ("Специалист (ведущий / 1 / 2 категории)","Исполнение услуг по закреплённым клиентам; запросы и ответы по своей зоне; с клиентом напрямую не общается — только через координатора (исключения: специалисты отдела судебно-претензионной работы; Главный бухгалтер при направлении счетов и актов); регистрация и исправление ошибок","Результаты услуг; запросы и ответы; записи об ошибках","1 р.д. / 2 р.д. / 3 р.д. по адресату"),
     ("Координатор клиента","Связующее звено в общении с клиентом: все запросы, документы и результаты клиенту и от клиента проходят только через него (исключения — отдел судебно-претензионной работы: напрямую с клиентом; Главный бухгалтер — направление счетов и актов, ПЛ7); распределяет материалы по направлениям, согласует дату запуска, контролирует общий договорный срок","Сводный запрос; пакет материалов; согласованная дата запуска; записи о претензиях и отзывах","1 р.д. подтверждение; 3 р.д. с клиентом"),
     ("Клиент","Предоставляет исходные документы и принимает хозяйственные решения; всё передаёт и получает только через координатора клиента","Первичные документы; решения клиента","Срок указывается в запросе"),
    ]
    trs = "".join(f'<tr><td><b>{esc(a)}</b></td><td>{esc(b)}</td><td>{esc(c)}</td><td>{esc(d)}</td></tr>' for a,b,c,d in rows)
    tb = f'<table class="t" style="font-size:16px"><tr><th style="width:250px">Роль</th><th>Зона ответственности и полномочия</th><th style="width:420px">Типовые документы-решения</th><th style="width:260px">Ключевые сроки</th></tr>{trs}</table>'
    return slide("Роли и полномочия (основа для должностных инструкций)", "Только согласованные полномочия; детализация — на этапе должностных инструкций", tb)

def backoffice():
    cards = [
     ("Продажи","#7a4fa8",["Первичная передача сведений о клиенте и составе услуг (ОС1)","Общая координация начального приёма","Оформление изменения состава услуг (ПС5)","Оформление прекращения услуги/договора (ОС3)","Участие в оценке/организации дополнительных услуг"],["Договор; перечень услуг; карточка клиента; изменение договора; соглашение о прекращении"],["Внутренний цикл продажи и заключения договора; формы и сроки"]),
     ("Технический отдел","#5b6b80",["Диагностика и исправление проблем ПО компании","Передача подрядчику при необходимости","Резервное копирование и восстановление","Поддержка доступов, интеграций и загрузок для ГИС","Реализация технических требований направлений"],["Функциональные требования направлений; обращения РЦ и ГИС"],["Состав технических заявок и учёта инцидентов; формы и сроки"]),
     ("Юридическое направление","#5b6b80",["Два отдела: сопровождения взыскания задолженностей за ЖКУ (включая работу с судом; «единое окно», через координатора) и судебно-претензионной работы (напрямую с клиентом; принимает исковые заявления на контроль сразу после подачи в суд)","Участвует в очерёдности назначения координатора","Получает данные о задолженности от РЦ"],["Данные о задолженности от РЦ; договор"],["Функциональные операции, входы/выходы, документы и сроки отделов — не согласованы"]),
    ]
    h = ""
    for t, c, known, io, gaps in cards:
        h += (f'<div class="card" style="flex:1;border-color:{c}"><h3 style="color:{c}">{t}</h3><b>Что известно</b><ul style="margin:4px 0 8px 18px;padding:0">' + "".join(f"<li>{esc(k)}</li>" for k in known) + f'</ul><b>Известные входы/выходы</b><br>{esc(io[0])}<div style="margin-top:8px;color:#9a2a2a"><b>Не проработано</b><br>{esc(gaps[0])}</div></div>')
    body = f'<div style="display:flex;gap:16px">{h}</div><div class="card" style="margin-top:16px;border-color:#b8620a">Эти блоки изображены как «чёрные ящики»: в сквозных процессах показаны только согласованные точки входа и выхода. Предположений о функциях не добавлено. Следующий этап — детальная проработка функций, входов/выходов, документов и сроков этих направлений.</div>'
    return slide("Обеспечивающие направления: продажи, технический отдел, юристы", "Показано только то, что согласовано в презентации", body)

def sla_slides():
    out = []
    t1 = "".join(f'<tr><td>{esc(a)}</td><td><b>{esc(b)}</b></td><td>{esc(c)}</td></tr>' for a,b,c in PDF_SLAS)
    out.append(slide("Регламентные сроки: согласованные нормативы", "Источник: презентация (р.д. — рабочие дни по производственному календарю РФ). Это сроки готовности ответа/действия, а не срок первичной реакции",
        f'<table class="t"><tr><th>Вид взаимодействия</th><th style="width:560px">Срок</th><th style="width:220px">Процесс</th></tr>{t1}</table>'))
    # новые
    news = [(p, s) for p in P for s in p["steps"] if s["src"] == "Н" and s["dl"] and s["kind"] != "decision"]
    pages = [news[i:i+17] for i in range(0, len(news), 17)]
    for i, pg in enumerate(pages):
        trs = "".join(f'<tr><td><b>{esc(s["id"])}</b></td><td>{esc(p["name"])}</td><td>{esc(lane_name(s["lane"]))}</td><td>{esc(s["title"])}</td><td><b>{esc(s["dl"])}</b></td></tr>' for p, s in pg)
        out.append(slide(f'Предложенные сроки: не прописаны в документации ({i+1}/{len(pages)})', "Требуют утверждения; внутренних нормативов нет, будут разработаны позже. Сроки для клиента — рекомендации, не обязательства",
            f'<table class="t" style="font-size:14px"><tr><th style="width:60px">Шаг</th><th style="width:380px">Процесс</th><th style="width:210px">Роль</th><th>Действие</th><th style="width:300px">Предлагаемый срок</th></tr>{trs}</table>'))
    # закон
    half = [LEGAL[:10], LEGAL[10:]]
    for i, pg in enumerate(half):
        trs = "".join(f'<tr><td><b>{esc(a)}</b></td><td>{esc(b)}</td><td>{esc(c)}</td></tr>' for a,b,c in pg)
        out.append(slide(f'Сроки по законодательству РФ ({i+1}/2)', "Ориентиры для календаря задач; перед утверждением сверить с актуальной редакцией норм на дату применения (законодательство меняется)",
            f'<table class="t" style="font-size:15px"><tr><th style="width:430px">Что</th><th>Срок</th><th style="width:300px">Норма</th></tr>{trs}</table>'))
    return out

STORE = {"client":("Папка клиента направления (электронный архив)","Срок договора + 3 года (ст. 196 ГК РФ)"),
         "acc":("Папка клиента бухгалтерии; архив первички и ЭДО","Не менее 5 лет после отчётного года; годовая отчётность — по Перечню Росархива №236"),
         "internal":("Рабочая система задач / переписка (на усмотрение руководителя)","1 год после закрытия задачи"),
         "quality":("Реестр качества направления","3 года"),
         "rc":("Архив РЦ по клиенту и периоду","5 лет после закрытия периода"),
         "gis":("Архив ГИС по клиенту","5 лет"),
         "hr":("Личное дело / служебные записки у начальника отдела","По решению руководства (уточнить по Перечню №236)")}
CAT = {"ОС1":"client","ОС2":"acc","ОЛ2.1":"acc","ОЛ2.2":"gis","ОЛ2.3":"rc","ОЛ2.4":"client","ОЛ2.5":"client","ОС3":"client","ОЛ4":"client","ПС5":"client","ПЛ6":"internal","ПЛ7":"acc"}

def registry():
    reg = {}
    for p in P:
        for s in p["steps"]:
            for d in s["out"]:
                r = reg.setdefault(d, dict(name=d, created=None, owner=None, cat=CAT[p["id"]], used=[]))
                if r["created"] is None: r["created"] = s["id"]; r["owner"] = lane_name(s["lane"]); r["cat"] = CAT[p["id"]]
            for d in s["inp"]:
                r = reg.setdefault(d, dict(name=d, created=None, owner="Внешний вход", cat=CAT[p["id"]], used=[]))
                r["used"].append(s["id"])
    return list(reg.values())

def registry_slides():
    rg = registry()
    out = []
    pages = [rg[i:i+19] for i in range(0, len(rg), 19)]
    for i, pg in enumerate(pages):
        trs = ""
        for r in pg:
            st, ret = STORE[r["cat"]]
            trs += (f'<tr><td><b>{esc(r["name"])}</b></td><td>{esc(r["owner"])}</td><td>{esc(r["created"] or "вне процессов")}</td>'
                    f'<td>{esc(", ".join(dict.fromkeys(r["used"])) or "—")}</td><td>{esc(st)}</td><td>{esc(ret)}</td></tr>')
        out.append(slide(f'Реестр документов: входы и продукты ({i+1}/{len(pages)})', "Хранение и сроки — значения по умолчанию, требуют утверждения; формы документов не утверждены — разработать",
            f'<table class="t" style="font-size:13.5px"><tr><th style="width:360px">Документ / результат</th><th style="width:220px">Создаёт (роль)</th><th style="width:80px">Шаг</th><th style="width:230px">Нужен на шагах</th><th style="width:340px">Хранение (по умолчанию)</th><th>Срок хранения (по умолчанию)</th></tr>{trs}</table>'))
    return out

ASSUME = [
 ("Допущение","Сквозные процессы показаны ролями, без численности и реальных наименований отделов. Должности: Руководитель <направления>, Начальник отдела <отдела>, Специалист <ведущий/1/2 категории>."),
 ("Допущение","Координатор клиента назначается для каждого клиента (в т. ч. при одной услуге) и является единственным каналом данных с клиентом. Это роль специалиста направления с наивысшим приоритетом услуг, а не отдельная должность."),
 ("Допущение","Исключение: до заключения и при изменении договора с клиентом общаются продажи (ОЛ4, ПС5); все вопросы сопровождения — только координатор, кроме отдела судебно-претензионной работы (ОЛ2.5), который общается с клиентом напрямую, и Главного бухгалтера, направляющего клиенту счета и акты (ПЛ7). Отдел сопровождения взыскания задолженностей (ОЛ2.4) работает по «единому окну». Жители (получатели услуг клиента-УК) клиентом не являются: обращения жителей решает РЦ."),
 ("Допущение","Рабочие дни — по производственному календарю РФ. Сроки из презентации — это готовность ответа/действия."),
 ("Допущение","Взыскание задолженности за ЖКУ (ОЛ2.4) включает работу с судом и выполняется условно автоматизированно. Исковое заявление сразу после оформления и отправки в суд, независимо от результата, переходит на контроль в отдел судебно-претензионной работы (ОЛ2.5). Судебный приказ ведёт отдел сопровождения взыскания: отмена должником либо передача выданного приказа приставам — окончание работы по приказу, после чего клиента информирует координатор."),
 ("Допущение","При передаче дела из ОЛ2.4 в ОЛ2.5 клиента информирует координатор. ОЛ2.5 сообщает координатору только об окончании работ; прочая коммуникация — с клиентом напрямую."),
 ("Открыто","ОЛ2.4: сроки большинства шагов не определены."),
 ("Допущение","Судебно-претензионная работа (ОЛ2.5) и продажи (ОЛ4) показаны заготовками: только согласованные точки входа/выхода, внутренние функции и сроки не домысливались."),
 ("Допущение","ПЛ7 «Выставление счёта и оплата услуг»: владелец — Главный бухгалтер; счёт и акт клиенту направляет он сам (исключение из правила координатора); ответ клиента (оплата, подписанный акт) — тоже ему, это моё допущение для подтверждения. Добавлены мной: ветка «данные не соответствуют договору → уточнить данные у координатора», шаг подтверждения завершения Главным бухгалтером, предложенные сроки (Н). Координатор может быть из другого направления — это интерфейс процесса, класс «локальный» сохранён."),
 ("Допущение","Классы присвоены так: ОС — приём, обслуживание, прекращение; ОЛ — бухгалтерия, ГИС, РЦ, юристы, привлечение (владелец — одно направление); ПС — изменение состава услуг; ПЛ — передача дел (внутри отдела). Привлечение клиента — ОЛ подтверждено (юристы не участвуют). Передача дел — ПЛ подтверждено (внутри отдела)."),
 ("Допущение","Процессы, убранные из набора (внутренний запрос, эскалация, ошибки, претензии, контроль качества, дисциплина), сохранены в резерве и в истории репозитория; предлагаемый класс — У. Нормативы из презентации по ним остаются в таблице сроков."),
 ("Допущение","Сроки «Н» (предложены) введены только там, где в документации срок не определён. Эскалация по времени (просрочка → уровень) не вводилась."),
 ("Допущение","Хранение документов и сроки — значения по умолчанию из норм РФ и практики; формы и шаблоны документов не утверждены."),
 ("Допущение","Система задач не определена; фиксация задач, запросов и записей — на усмотрение руководителя."),
 ("Проверить","Законные сроки приведены как ориентиры на 17.11.2026: перед утверждением регламента сверить с актуальной редакцией НК РФ, ТК РФ, ЖК РФ, 209-ФЗ и приказов."),
 ("Открыто","Функционал юридического направления, внутренние процессы продаж и техотдела, модель заявок и инцидентов."),
 ("Открыто","Общая система кадрового обеспечения и расчёта численности; числовые KPI; номенклатура и сроки архива."),
 ("Открыто","Сроки по клиентским срокам (передача первички до 5 числа, согласование отчётности за 3 р.д. до сдачи) — предложены, нужны решение руководства."),
 ("Следующий этап","Чек-листы и должностные инструкции на основе утверждённых процессов, ролей и сроков."),
]
def assumptions():
    trs = "".join(f'<tr><td><b>{esc(a)}</b></td><td>{esc(b)}</td></tr>' for a,b in ASSUME)
    return slide("Допущения, что проверить и открытые вопросы", "Что принято по умолчанию и что остаётся за руководством", f'<table class="t" style="font-size:17px"><tr><th style="width:200px">Тип</th><th>Содержание</th></tr>{trs}</table>')

def main():
    slides = [cover(), classif(), procmap(), orgchart(), roles(), backoffice()]
    for p in P: slides.extend(process_slides(p))
    slides.extend(sla_slides()); slides.extend(registry_slides()); slides.append(assumptions())
    doc = f'<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>Документация бизнес-процессов</title><style>{CSS}</style></head><body>{"".join(slides)}</body></html>'
    open("../deck.html", "w", encoding="utf-8").write(doc)
    print("slides", len(slides))

if __name__ == "__main__": main()
