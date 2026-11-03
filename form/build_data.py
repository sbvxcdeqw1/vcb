#!/usr/bin/env python3
"""Готовит справочники для формы показаний из выгрузок 1С.

  python3 build_data.py --accounts список_помещений.xlsx --meters ввод_показаний.xlsx [ещё.xlsx ...] -o папка

accounts.csv — Лицевой счёт;Помещение;Объект учета;Площадь
    счёт — до первой запятой в «Лицевой счет» (12–13 цифр); помещение — после третьей запятой в «Жилое помещение»:
    «кв. 5» -> «5», без номера квартиры (нежилое, «кв. Иванов И.И.», «кв. 0») -> «Помещение ИИИ» (инициалы владельца).
    Полные имена в файл не попадают.
meters.csv   — Порядок;Объект учета;Прибор учета;Заводской номер;Тип;Место;Пред. показание;Дата пред.
    из документов «Ввод показаний индивидуальных приборов учета»; последнее показание каждого прибора
    становится «предыдущим» для следующего приёма. Можно передать несколько документов — берётся самое свежее.
"""
import argparse, csv, datetime as dt, os, re, sys
import openpyxl


def initials(fio):
    fio = re.sub(r'\(.*?\)', ' ', fio or '')
    return ''.join(t[0] for t in re.split(r'[\s.]+', fio) if t).upper()


def read_accounts(path):
    ws = openpyxl.load_workbook(path, read_only=True).active
    out, skipped = {}, []
    for r in ws.iter_rows(values_only=True):
        if len(r) < 4 or not (isinstance(r[0], str) and ',' in r[0] and r[3]):
            continue
        acc, _, holder = r[0].partition(',')
        acc, obj = acc.strip(), str(r[3]).strip()
        parts = obj.split(',')
        place = ','.join(parts[3:]).strip() if len(parts) > 3 else ''
        m = re.fullmatch(r'кв\.\s*(\d+)', place)
        place = m.group(1) if m and int(m.group(1)) > 0 else 'Помещение ' + initials(holder)
        area = r[8] if len(r) > 8 and isinstance(r[8], (int, float)) else ''
        if not re.fullmatch(r'\d{12,13}', acc) or place == 'Помещение ':
            skipped.append(f'{acc}: {obj}'); continue
        if acc in out:
            skipped.append(f'{acc}: дубликат'); continue
        out[acc] = (place, obj, area)
    return out, skipped


def meter_kind(name):
    """'Х Вван ' -> ('ХВ', 'ванная'); 'ГВ' -> ('ГВ', ''); 'ХВ 2' -> ('ХВ', '№2')."""
    s = re.sub(r'\s+', '', name or '').upper()
    kind = s[:2] if s[:2] in ('ХВ', 'ГВ') else ''
    rest = s[2:] if kind else s
    if rest.startswith('ВАН'): place = 'ванная'
    elif rest.startswith('КУХ'): place = 'кухня'
    elif rest.isdigit(): place = '№' + rest
    else: place = rest.lower()
    return kind, place


def read_meters(path):
    ws = openpyxl.load_workbook(path, read_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    out = []
    for i, r in enumerate(rows):
        if len(r) < 8 or not (r[1] is not None and str(r[1]).strip().isdigit() and r[2] and r[3]):
            continue
        nxt = rows[i + 1] if i + 1 < len(rows) else ()
        date = nxt[6] if len(nxt) > 6 else None                      # дата текущего показания
        if isinstance(date, dt.datetime): date = date.strftime('%d.%m.%Y')
        out.append(dict(n=int(str(r[1]).strip()), obj=str(r[2]).strip(), name=str(r[3]), serial=str(r[4] or '').strip(),
                        value=r[6], date=str(date or '')))
    return out


def ddmmyyyy_key(s):
    m = re.fullmatch(r'(\d\d)\.(\d\d)\.(\d{4})', s or '')
    return (m.group(3), m.group(2), m.group(1)) if m else ('',)


def num(x):
    return f'{x:.3f}'.replace('.', ',') if isinstance(x, (int, float)) else ''


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--accounts', required=True)
    ap.add_argument('--meters', nargs='*', default=[])
    ap.add_argument('-o', '--out', default='.')
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)

    accounts, skipped = read_accounts(a.accounts)
    with open(os.path.join(a.out, 'accounts.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(['Лицевой счёт', 'Помещение', 'Объект учета', 'Площадь'])
        for acc, (place, obj, area) in accounts.items():
            w.writerow([acc, place, obj, str(area).replace('.', ',')])
    print(f'accounts.csv: {len(accounts)} счетов, пропущено {len(skipped)}')
    for s in skipped: print('  пропущено:', s)

    if not a.meters:
        return
    meters = {}                                       # (объект, прибор, номер) -> запись; порядок — как в первом документе
    for path in a.meters:
        for m in read_meters(path):
            key = (m['obj'], m['name'].strip(), m['serial'])
            old = meters.get(key)
            if old is None:
                m['order'] = len(meters) + 1; meters[key] = m
            elif ddmmyyyy_key(m['date']) >= ddmmyyyy_key(old['date']):
                m['order'] = old['order']; meters[key] = m
    with open(os.path.join(a.out, 'meters.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(['Порядок', 'Объект учета', 'Прибор учета', 'Заводской номер', 'Тип', 'Место', 'Пред. показание', 'Дата пред.'])
        for m in sorted(meters.values(), key=lambda m: m['order']):
            kind, place = meter_kind(m['name'])
            w.writerow([m['order'], m['obj'], m['name'], m['serial'], kind, place, num(m['value']), m['date']])
    objs = {m['obj'] for m in meters.values()}
    known = {obj for _, obj, _ in accounts.values()}
    print(f'meters.csv: {len(meters)} приборов в {len(objs)} помещениях')
    print(f'  счетов без приборов в документах: {sum(1 for _, o, _ in accounts.values() if o not in objs)} (для них в форме — ручной ввод)')
    for o in sorted(objs - known): print('  приборы есть, лицевого счёта нет:', o)
    unk = sorted({m['name'] for m in meters.values() if not meter_kind(m['name'])[0]})
    for n in unk: print('  не распознан тип прибора:', repr(n))


if __name__ == '__main__':
    sys.exit(main())
