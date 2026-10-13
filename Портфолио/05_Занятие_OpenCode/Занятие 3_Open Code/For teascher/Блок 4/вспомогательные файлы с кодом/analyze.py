# -*- coding: utf-8 -*-
import xlrd, sys
from collections import OrderedDict

sys.stdout.reconfigure(encoding='utf-8')

files = [
    (r'C:\Users\user\Downloads\for_Cloude Code\prepord-card\Занятие 3_Open Code\Блок 4\численность по площадкам-типам-классам 2-05-24.xls', 2024),
    (r'C:\Users\user\Downloads\for_Cloude Code\prepord-card\Занятие 3_Open Code\Блок 4\площадка-тип_групп-класс_06.05.2025.xls', 2025),
]

def parse_file(filepath):
    wb = xlrd.open_workbook(filepath)
    ws = wb.sheet_by_index(0)
    header_row = None
    for r in range(ws.nrows):
        if ws.cell_value(r, 0) == 'Площадки':
            header_row = r
            break

    data = []
    current_platform = None
    current_type = None

    for r in range(header_row + 1, ws.nrows):
        col0 = str(ws.cell_value(r, 0)).strip()
        col1 = str(ws.cell_value(r, 1)).strip()
        col2 = str(ws.cell_value(r, 2)).strip()
        col3 = ws.cell_value(r, 3)

        if col3 == '' or col3 == 0:
            continue

        if 'Итого' in col0 or 'Итого' in col2:
            continue

        val = int(col3) if isinstance(col3, float) else col3

        if col0 != '' and col1 == '' and col2 == '':
            current_platform = col0
            current_type = None
            data.append({
                'platform': current_platform,
                'type': 'ВСЕГО',
                'class': '',
                'value': int(val)
            })
        elif col0 == '' and col1 != '' and col2 == '':
            current_type = col1
            data.append({
                'platform': current_platform,
                'type': current_type,
                'class': 'ВСЕГО',
                'value': int(val)
            })
        elif col0 == '' and col1 == '' and col2 != '':
            data.append({
                'platform': current_platform,
                'type': current_type,
                'class': str(int(col2)) if isinstance(col2, float) else str(col2),
                'value': int(val)
            })

    return data

data_2024 = parse_file(files[0][0])
data_2025 = parse_file(files[1][0])

detail_2024 = [d for d in data_2024 if d['class'] not in ('', 'ВСЕГО')]
detail_2025 = [d for d in data_2025 if d['class'] not in ('', 'ВСЕГО')]
type_total_2024 = [d for d in data_2024 if d['class'] == 'ВСЕГО' and d['type'] != 'ВСЕГО']
type_total_2025 = [d for d in data_2025 if d['class'] == 'ВСЕГО' and d['type'] != 'ВСЕГО']
platform_total_2024 = [d for d in data_2024 if d['type'] == 'ВСЕГО']
platform_total_2025 = [d for d in data_2025 if d['type'] == 'ВСЕГО']

print('=' * 70)
print('1. ПРОВЕРКА СХОДИМОСТИ ДАННЫХ')
print('=' * 70)

sum_detail_2024 = sum(d['value'] for d in detail_2024)
sum_detail_2025 = sum(d['value'] for d in detail_2025)

print()
print(f'2024: сумма детальных строк (классы) = {sum_detail_2024}, заявленный итог = 363')
print(f'2025: сумма детальных строк (классы) = {sum_detail_2025}, заявленный итог = 304')

print()
print('Проверка по площадкам 2024:')
for pt in platform_total_2024:
    plat = pt['platform']
    classes_sum = sum(d['value'] for d in detail_2024 if d['platform'] == plat)
    if classes_sum == pt['value']:
        match = 'OK'
    else:
        match = f'РАСХОЖДЕНИЕ: итог={pt["value"]}, сумма классов={classes_sum}'
    print(f'  {plat}: {match}')

print()
print('Проверка по площадкам 2025:')
for pt in platform_total_2025:
    plat = pt['platform']
    classes_sum = sum(d['value'] for d in detail_2025 if d['platform'] == plat)
    if classes_sum == pt['value']:
        match = 'OK'
    else:
        match = f'РАСХОЖДЕНИЕ: итог={pt["value"]}, сумма классов={classes_sum}'
    print(f'  {plat}: {match}')

print()
print('Проверка по типам групп 2024:')
for tt in type_total_2024:
    typ = tt['type']
    classes_sum = sum(d['value'] for d in detail_2024 if d['type'] == typ)
    if classes_sum == tt['value']:
        match = 'OK'
    else:
        match = f'РАСХОЖДЕНИЕ: итог={tt["value"]}, сумма классов={classes_sum}'
    print(f'  {typ}: {match}')

print()
print('Проверка по типам групп 2025:')
for tt in type_total_2025:
    typ = tt['type']
    classes_sum = sum(d['value'] for d in detail_2025 if d['type'] == typ)
    if classes_sum == tt['value']:
        match = 'OK'
    else:
        match = f'РАСХОЖДЕНИЕ: итог={tt["value"]}, сумма классов={classes_sum}'
    print(f'  {typ}: {match}')

print()
print('=' * 70)
print('2. СВОДНАЯ ТАБЛИЦА ПО ПЛОЩАДКАМ')
print('=' * 70)

platforms_all = sorted(set(d['platform'] for d in data_2024 + data_2025))
header = f'{"Площадка":<20} {"02.05.2024":>10} {"06.05.2025":>10} {"Динамика":>10} {"Δ, %":>10}'
print(header)
print('-' * 60)
for p in platforms_all:
    v24 = sum(d['value'] for d in detail_2024 if d['platform'] == p)
    v25 = sum(d['value'] for d in detail_2025 if d['platform'] == p)
    if v24 == 0:
        dyn = f'+{v25}'
        pct = 'новая'
    elif v25 == 0:
        dyn = f'-{v24}'
        pct = '-100.00 %'
    else:
        diff = v25 - v24
        if diff > 0:
            dyn = f'+{diff}'
        else:
            dyn = f'{diff}'
        pct = f'{diff / v24 * 100:.2f} %'
    print(f'{p:<20} {v24:>10} {v25:>10} {dyn:>10} {pct:>10}')

total_24 = sum(d['value'] for d in detail_2024)
total_25 = sum(d['value'] for d in detail_2025)
diff_total = total_25 - total_24
pct_total = diff_total / total_24 * 100
print('-' * 60)
print(f'{"ИТОГО":<20} {total_24:>10} {total_25:>10} {diff_total:>10} {pct_total:.2f} %')

print()
print('=' * 70)
print('3. СВОДНАЯ ТАБЛИЦА ПО ТИПАМ ГРУПП')
print('=' * 70)

types_all = sorted(set(d['type'] for d in detail_2024 + detail_2025))
header = f'{"Тип группы":<10} {"02.05.2024":>10} {"06.05.2025":>10} {"Динамика":>10} {"Δ, %":>10}'
print(header)
print('-' * 50)
for t in types_all:
    v24 = sum(d['value'] for d in detail_2024 if d['type'] == t)
    v25 = sum(d['value'] for d in detail_2025 if d['type'] == t)
    if v24 == 0:
        dyn = f'+{v25}'
        pct = 'новый'
    elif v25 == 0:
        dyn = f'-{v24}'
        pct = '-100.00 %'
    else:
        diff = v25 - v24
        if diff > 0:
            dyn = f'+{diff}'
        else:
            dyn = f'{diff}'
        pct = f'{diff / v24 * 100:.2f} %'
    print(f'{t:<10} {v24:>10} {v25:>10} {dyn:>10} {pct:>10}')
print('-' * 50)
print(f'{"ИТОГО":<10} {total_24:>10} {total_25:>10} {diff_total:>10} {pct_total:.2f} %')

print()
print('=' * 70)
print('4. АНАЛИЗ ИЗМЕНЕНИЙ В НАБОРАХ ГРУПП')
print('=' * 70)

types_2024 = set(d['type'] for d in detail_2024)
types_2025 = set(d['type'] for d in detail_2025)
gone = types_2024 - types_2025
new = types_2025 - types_2024
common = types_2024 & types_2025

print()
if gone:
    print(f'Типы групп, которые БЫЛИ в 2024 и ИСЧЕЗЛИ в 2025: {sorted(gone)}')
    for t in sorted(gone):
        v = sum(d['value'] for d in detail_2024 if d['type'] == t)
        print(f'  {t}: {v} чел.')
else:
    print('Типы групп, которые БЫЛИ в 2024 и ИСЧЕЗЛИ в 2025: —')

print()
if new:
    print(f'Типы групп, которые ПОЯВИЛИСЬ в 2025: {sorted(new)}')
    for t in sorted(new):
        v = sum(d['value'] for d in detail_2025 if d['type'] == t)
        print(f'  {t}: {v} чел.')
else:
    print('Типы групп, которые ПОЯВИЛИСЬ в 2025: —')

print()
print('Типы групп, которые остались (изменение численности):')
print(f'{"Тип":<6} {"2024":>6} {"2025":>6} {"Δ":>6} {"Δ%":>8}')
print('-' * 34)
for t in sorted(common):
    v24 = sum(d['value'] for d in detail_2024 if d['type'] == t)
    v25 = sum(d['value'] for d in detail_2025 if d['type'] == t)
    diff = v25 - v24
    pct = diff / v24 * 100
    if diff > 0:
        dyn = f'+{diff}'
    else:
        dyn = f'{diff}'
    print(f'{t:<6} {v24:>6} {v25:>6} {dyn:>6} {pct:+7.2f} %')

print()
print('=' * 70)
print('5. РАСПРЕДЕЛЕНИЕ ПО КЛАССАМ (ВОЗРАСТНОЙ СОСТАВ)')
print('=' * 70)

classes_2024 = {}
classes_2025 = {}
for d in detail_2024:
    c = d['class']
    classes_2024[c] = classes_2024.get(c, 0) + d['value']
for d in detail_2025:
    c = d['class']
    classes_2025[c] = classes_2025.get(c, 0) + d['value']

all_classes = sorted(set(list(classes_2024.keys()) + list(classes_2025.keys())), key=lambda x: int(x))
print(f'{"Класс":<8} {"02.05.2024":>10} {"% 2024":>8} {"06.05.2025":>10} {"% 2025":>8} {"Δ чел":>8}')
print('-' * 52)
for c in all_classes:
    v24 = classes_2024.get(c, 0)
    v25 = classes_2025.get(c, 0)
    p24 = v24 / total_24 * 100
    p25 = v25 / total_25 * 100
    diff = v25 - v24
    if diff > 0:
        dyn = f'+{diff}'
    else:
        dyn = f'{diff}'
    print(f'{c:<8} {v24:>10} {p24:>7.2f} % {v25:>10} {p25:>7.2f} % {dyn:>8}')
print('-' * 52)
print(f'{"ИТОГО":<8} {total_24:>10} {100:>7.2f} % {total_25:>10} {100:>7.2f} % {diff_total:>+8}')

def age_group(c):
    c_int = int(c)
    if c_int <= 6:
        return '1'
    elif c_int <= 9:
        return '2'
    else:
        return '3'

age_names = {'1': '1-6 (младшие)', '2': '7-9 (средние)', '3': '10-12 (старшие)'}
print()
print('Группировка по возрастам:')
print(f'{"Категория":<20} {"2024":>8} {"%2024":>8} {"2025":>8} {"%2025":>8}')
print('-' * 52)
for grp_key in ['1', '2', '3']:
    v24 = sum(classes_2024.get(c, 0) for c in all_classes if age_group(c) == grp_key)
    v25 = sum(classes_2025.get(c, 0) for c in all_classes if age_group(c) == grp_key)
    p24 = v24 / total_24 * 100
    p25 = v25 / total_25 * 100
    print(f'{age_names[grp_key]:<20} {v24:>8} {p24:>7.2f} % {v25:>8} {p25:>7.2f} %')

print()
print('=' * 70)
print('6. ВЫВОДЫ')
print('=' * 70)
print()
print('**Вывод 1. Общее снижение численности**')
print(f'  Факт: за год число слушателей сократилось на {abs(diff_total)} чел.')
print(f'  Цифра: {total_24} → {total_25} ({pct_total:.2f} %)')
print(f'  Детали: снижение затронуло 5 из 6 старых площадок; новая площадка Наличная (+6 чел.) не компенсирует падение.')
print()
print('**Вывод 2. Сизова — ключевая площадка с наибольшим падением**')
v_sizova_24 = sum(d['value'] for d in detail_2024 if d['platform'] == 'Сизова')
v_sizova_25 = sum(d['value'] for d in detail_2025 if d['platform'] == 'Сизова')
print(f'  Факт: Сизова потеряла больше всех — {v_sizova_24 - v_sizova_25} чел.')
print(f'  Цифра: {v_sizova_24} → {v_sizova_25} ({(v_sizova_25 - v_sizova_24) / v_sizova_24 * 100:.2f} %)')
print(f'  Детали: это {abs(v_sizova_25 - v_sizova_24) / abs(diff_total) * 100:.0f} % от общего сокращения.')
print()
print('**Вывод 3. Структура типов групп изменилась**')
for t in sorted(gone):
    print(f'  Исчез: {t} ({sum(d["value"] for d in detail_2024 if d["type"] == t)} чел.)')
for t in sorted(new):
    print(f'  Появился: {t} ({sum(d["value"] for d in detail_2025 if d["type"] == t)} чел.)')
print(f'  Детали: возможно, связано с переименованием или реорганизацией программ.')
print()
print('**Вывод 4. Смещение возрастного состава в средние классы**')
print(f'  Факт: доля средних классов (7-9) выросла, младших (1-6) — снизилась')
for grp_key in ['1', '2', '3']:
    v24 = sum(classes_2024.get(c, 0) for c in all_classes if age_group(c) == grp_key)
    v25 = sum(classes_2025.get(c, 0) for c in all_classes if age_group(c) == grp_key)
    p24 = v24 / total_24 * 100
    p25 = v25 / total_25 * 100
    print(f'  {age_names[grp_key]}: {p24:.1f}% → {p25:.1f}%')
print()
print('**Вывод 5. Появилась новая площадка Наличная**')
print(f'  Факт: в 2025 открылась площадка Наличная')
print(f'  Цифра: 6 чел. (2 % от общего числа)')
print(f'  Детали: пока небольшой вклад, но это расширение географии.')
