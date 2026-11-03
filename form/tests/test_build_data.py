"""pytest form/tests/test_build_data.py"""
import csv, subprocess, sys
from pathlib import Path
import openpyxl

TOOL = Path(__file__).resolve().parents[1] / 'build_data.py'


def _accounts(path):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(['Лицевой счет', None, None, 'Жилое помещение', None, None, 'Ответственный потребитель', None, None])
    ws.append(['г.Тестовск, Тестовая, д.1', None, None, None, None, None, None, 10, 100])
    ws.append(['990000000122, Примеров Иван Иванович', None, None, 'г.Тестовск, Тестовая, д.1, кв. 1', None, None, 'x', 1, 76.5])
    ws.append(['9900000009022, Образцов П.П.', None, None, 'г.Тестовск, Тестовая, д.1, кв. 90', None, None, 'x', 1, 49.9])
    ws.append(['990000026222, Тестова А.С. (Магазин)', None, None, 'Тестовск г, Тестовая, д.1, кв. Тестова А.С. (Магазин)', None, None, 'x', None, 41.7])
    ws.append(['990000026822, Пробин А.Н.', None, None, 'г.Тестовск, Тестовая, д.1, кв. 0', None, None, 'x', None, 95.3])
    ws.append(['99000000012, Короткий', None, None, 'г.Тестовск, Тестовая, д.1, кв. 5', None, None, 'x', None, 1])
    wb.save(path)


def _meters(path, value, date):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append([None, 'Ввод показаний индивидуальных приборов учета'])
    ws.append([None, 'N', 'Объект учета', 'Прибор учета', 'Заводской номер', 'Предыдущее показание', 'Текущее показание', 'Объем'])
    ws.append([None, None, None, 'Тарифная зона', 'Коэффициент', 'Дата пред.', 'Дата показания'])
    ws.append([None, '1', 'г.Тестовск, Тестовая, д.1, кв. 1', 'Х Вван ', '0100000001', 1.0, value, 1.0])
    ws.append([None, None, None, None, None, '01.01.2026', date])
    ws.append([None, '2', 'г.Тестовск, Тестовая, д.1, кв. 1', 'ГВ', '', 1.0, 2.5, 1.5])
    ws.append([None, None, None, None, None, '01.01.2026', date])
    ws.append([None, None, None, 'ХВ', None, None, None, 9.9])          # итоговая строка — не прибор
    wb.save(path)


def _read(p):
    with open(p, encoding='utf-8-sig') as f:
        return list(csv.reader(f, delimiter=';'))


def test_build(tmp_path):
    _accounts(tmp_path / 'a.xlsx')
    _meters(tmp_path / 'm1.xlsx', 5.0, '23.09.2026')
    _meters(tmp_path / 'm2.xlsx', 7.25, '23.10.2026')
    r = subprocess.run([sys.executable, str(TOOL), '--accounts', str(tmp_path / 'a.xlsx'),
                        '--meters', str(tmp_path / 'm2.xlsx'), str(tmp_path / 'm1.xlsx'), '-o', str(tmp_path)],
                       capture_output=True, text=True, check=True)
    acc = _read(tmp_path / 'accounts.csv')
    assert acc[0] == ['Лицевой счёт', 'Помещение', 'Объект учета', 'Площадь']
    assert acc[1] == ['990000000122', '1', 'г.Тестовск, Тестовая, д.1, кв. 1', '76,5']
    assert acc[2][:2] == ['9900000009022', '90']
    assert acc[3][1] == 'Помещение ТАС'
    assert acc[4][1] == 'Помещение ПАН'                     # «кв. 0» — без номера
    assert len(acc) == 5 and 'пропущено 1' in r.stdout     # 11 цифр — пропущен
    assert 'Примеров' not in (tmp_path / 'accounts.csv').read_text(encoding='utf-8-sig')
    met = _read(tmp_path / 'meters.csv')
    assert met[1] == ['1', 'г.Тестовск, Тестовая, д.1, кв. 1', 'Х Вван ', '0100000001', 'ХВ', 'ванная', '7,250', '23.10.2026']
    assert met[2][4:7] == ['ГВ', '', '2,500']
    assert len(met) == 3                                    # итоговые строки не попали, дубли слиты по свежей дате
