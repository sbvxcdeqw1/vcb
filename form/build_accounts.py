#!/usr/bin/env python3
"""Из выгрузки 1С «Список жилых помещений» (xlsx) делает accounts.csv: Лицевой счёт;Помещение.

Лицевой счёт — часть ячейки до первой запятой (12 или 13 цифр), помещение — всё после третьей запятой
в столбце «Жилое помещение»: «кв. 5» -> «5». Если номера квартиры нет (нежилое помещение, «кв. Иванов И.И.», «кв. 0»),
выводится «Помещение ИИИ» — первые буквы ФИО владельца счёта. Полные имена в файл не попадают. Использование: python3 build_accounts.py выгрузка.xlsx accounts.csv
"""
import csv, re, sys
import openpyxl

def initials(fio):
    fio = re.sub(r'\(.*?\)', ' ', fio)                      # «(Аптека)» и т.п. не берём
    return ''.join(t[0] for t in re.split(r'[\s.]+', fio) if t).upper()


src, dst = sys.argv[1], sys.argv[2]
ws = openpyxl.load_workbook(src, read_only=True).active
out, skipped = {}, []
for r in ws.iter_rows(values_only=True):
    if not (r[0] and r[3]) or not isinstance(r[0], str) or ',' not in r[0]:
        continue
    acc, _, holder = r[0].partition(',')
    acc = acc.strip()
    parts = str(r[3]).split(',')
    place = ','.join(parts[3:]).strip() if len(parts) > 3 else ''
    m = re.fullmatch(r'кв\.\s*(\d+)', place)
    if m and int(m.group(1)) > 0:
        place = m.group(1)
    else:
        place = 'Помещение ' + initials(holder)
    if not re.fullmatch(r'\d{12,13}', acc) or place == 'Помещение ':
        skipped.append((acc, r[3])); continue
    if acc in out:
        skipped.append((acc, 'дубликат счёта')); continue
    out[acc] = place
with open(dst, 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f, delimiter=';'); w.writerow(['Лицевой счёт', 'Помещение']); w.writerows(out.items())
print(f'записано {len(out)}, пропущено {len(skipped)}')
for s in skipped: print('  пропущено:', *s)
