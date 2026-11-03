#!/usr/bin/env python3
"""Из выгрузки 1С «Список жилых помещений» (xlsx) делает accounts.csv: Лицевой счёт;Помещение.

Лицевой счёт — часть ячейки до первой запятой (ровно 12 цифр), помещение — всё после третьей запятой
в столбце «Жилое помещение» («кв. 5» -> «5»; иное, например «н/п Иванов», остаётся текстом).
Имена жильцов в файл не попадают. Использование: python3 build_accounts.py выгрузка.xlsx accounts.csv
"""
import csv, re, sys
import openpyxl

src, dst = sys.argv[1], sys.argv[2]
ws = openpyxl.load_workbook(src, read_only=True).active
out, skipped = {}, []
for r in ws.iter_rows(values_only=True):
    if not (r[0] and r[3]) or not isinstance(r[0], str) or ',' not in r[0]:
        continue
    acc = r[0].split(',')[0].strip()
    parts = str(r[3]).split(',')
    place = ','.join(parts[3:]).strip() if len(parts) > 3 else ''
    m = re.fullmatch(r'кв\.\s*(\d+)', place)
    place = m.group(1) if m else re.sub(r'^кв\.\s*', 'кв. ', place)
    if not re.fullmatch(r'\d{12}', acc) or not place:
        skipped.append((acc, r[3])); continue
    if acc in out:
        skipped.append((acc, 'дубликат счёта')); continue
    out[acc] = place
with open(dst, 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f, delimiter=';'); w.writerow(['Лицевой счёт', 'Помещение']); w.writerows(out.items())
print(f'записано {len(out)}, пропущено {len(skipped)}')
for s in skipped: print('  пропущено:', *s)
