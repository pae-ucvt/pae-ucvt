# -*- coding: utf-8 -*-
import xlrd, sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.chart import BarChart, LineChart, RadarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

# ─── Парсинг ───
files = [
    (r'C:\Users\user\Downloads\for_Cloude Code\prepord-card\Занятие 3_Open Code\Блок 4\численность по площадкам-типам-классам 2-05-24.xls', 2024),
    (r'C:\Users\user\Downloads\for_Cloude Code\prepord-card\Занятие 3_Open Code\Блок 4\площадка-тип_групп-класс_06.05.2025.xls', 2025),
]

def parse_file(filepath):
    wb = xlrd.open_workbook(filepath)
    ws = wb.sheet_by_index(0)
    hr = None
    for r in range(ws.nrows):
        if ws.cell_value(r, 0) == 'Площадки':
            hr = r; break
    detail = []
    cp = None; ct = None
    for r in range(hr + 1, ws.nrows):
        c0 = str(ws.cell_value(r, 0)).strip()
        c1 = str(ws.cell_value(r, 1)).strip()
        c2 = str(ws.cell_value(r, 2)).strip()
        c3 = ws.cell_value(r, 3)
        if c3 == '' or c3 == 0: continue
        if 'Итого' in c0 or 'Итого' in c2: continue
        v = int(c3) if isinstance(c3, float) else c3
        if c0 != '' and c1 == '' and c2 == '':
            cp = c0; ct = None
        elif c0 == '' and c1 != '' and c2 == '':
            ct = c1
        elif c0 == '' and c1 == '' and c2 != '':
            cl = str(int(c2)) if isinstance(c2, float) else str(c2)
            detail.append({'platform': cp, 'type': ct, 'class': cl, 'value': int(v)})
    return detail

d24 = parse_file(files[0][0])
d25 = parse_file(files[1][0])

def agg(seq, key):
    d = {}
    for item in seq:
        k = item[key]
        d[k] = d.get(k, 0) + item['value']
    return d

t24 = agg(d24, 'type')
t25 = agg(d25, 'type')
all_types = sorted(set(list(t24.keys()) + list(t25.keys())))

# Разделяем типы
changed = [t for t in all_types if t24.get(t, 0) > 0 and t25.get(t, 0) > 0]
drastic = [t for t in all_types if t24.get(t, 0) == 0 or t25.get(t, 0) == 0]

total_24 = sum(d['value'] for d in d24)
total_25 = sum(d['value'] for d in d25)

# Стили
thin = Side(style='thin')
b_all = Border(left=thin, right=thin, top=thin, bottom=thin)
h_font = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
h_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
d_font = Font(name='Calibri', size=11)
db_font = Font(name='Calibri', size=11, bold=True)
t_font = Font(name='Calibri', bold=True, size=14, color='2F5496')

def hdr(ws, row, cols):
    for c in range(1, cols + 1):
        cl = ws.cell(row=row, column=c)
        cl.font = h_font; cl.fill = h_fill; cl.border = b_all
        cl.alignment = Alignment(horizontal='center', vertical='center')

def write_data(ws, row, data_list):
    for i, (k, v24, v25) in enumerate(data_list, row):
        ws.cell(row=i, column=1, value=k).font = d_font
        ws.cell(row=i, column=2, value=v24).font = d_font
        ws.cell(row=i, column=3, value=v25).font = d_font
        for c in range(1, 4):
            ws.cell(row=i, column=c).border = b_all
            if c > 1:
                ws.cell(row=i, column=c).alignment = Alignment(horizontal='center')

def add_bar_chart(ws, title, anchor, data_ref, cat_ref, w=16, h=10):
    ch = BarChart()
    ch.type = 'col'; ch.style = 10; ch.width = w; ch.height = h
    ch.title = title; ch.legend.position = 'b'
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cat_ref)
    ch.series[0].graphicalProperties.solidFill = '4472C4'
    ch.series[1].graphicalProperties.solidFill = 'ED7D31'
    for s in ch.series:
        s.dLbls = DataLabelList(); s.dLbls.showVal = True; s.dLbls.numFmt = '0'
        s.dLbls.font = Font(size=9)
    ws.add_chart(ch, anchor)

# ═══════════════════════════════════════════════
# ФАЙЛ B: Линейчатая + Лепестковая диаграммы
# ═══════════════════════════════════════════════
wb = Workbook()

# --- Данные ---
wsd = wb.active
wsd.title = 'Данные'
wsd.cell(row=1, column=1, value='Типы групп — все').font = t_font
for c, h in enumerate(['Тип группы', '2024', '2025'], 1):
    wsd.cell(row=2, column=c, value=h)
hdr(wsd, 2, 3)
row_data = 3
for t in all_types:
    v24 = t24.get(t, 0); v25 = t25.get(t, 0)
    wsd.cell(row=row_data, column=1, value=t).font = d_font
    wsd.cell(row=row_data, column=2, value=v24).font = d_font
    wsd.cell(row=row_data, column=3, value=v25).font = d_font
    for c in range(1, 4):
        wsd.cell(row=row_data, column=c).border = b_all
        if c > 1: wsd.cell(row=row_data, column=c).alignment = Alignment(horizontal='center')
    row_data += 1
# Итого
wsd.cell(row=row_data, column=1, value='ИТОГО').font = db_font
wsd.cell(row=row_data, column=2, value=total_24).font = db_font
wsd.cell(row=row_data, column=3, value=total_25).font = db_font
for c in range(1, 4):
    wsd.cell(row=row_data, column=c).border = b_all
    wsd.cell(row=row_data, column=c).fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
    if c > 1: wsd.cell(row=row_data, column=c).alignment = Alignment(horizontal='center')

last_data = row_data - 1

# --- Лист: Линейчатая диаграмма ---
ws_line = wb.create_sheet('Линейчатая')
ws_line.cell(row=1, column=1, value='Линейчатая диаграмма: численность по типам групп').font = t_font
ch_line = LineChart()
ch_line.title = 'Численность по типам групп (2024 vs 2025)'
ch_line.style = 10; ch_line.width = 22; ch_line.height = 14
ch_line.y_axis.title = 'Человек'
ch_line.x_axis.tickLblPos = 'low'
cats = Reference(wsd, min_col=1, min_row=3, max_row=last_data)
d = Reference(wsd, min_col=2, max_col=3, min_row=2, max_row=last_data)
ch_line.add_data(d, titles_from_data=True)
ch_line.set_categories(cats)
ch_line.series[0].graphicalProperties.line.solidFill = '4472C4'
ch_line.series[1].graphicalProperties.line.solidFill = 'ED7D31'
ch_line.series[0].graphicalProperties.line.width = 25000  # EMUs ~2pt
ch_line.series[1].graphicalProperties.line.width = 25000
for s in ch_line.series:
    s.dLbls = DataLabelList(); s.dLbls.showVal = True; s.dLbls.numFmt = '0'
    s.dLbls.font = Font(size=7)
ws_line.add_chart(ch_line, 'A3')

# --- Лист: Лепестковая диаграмма ---
ws_radar = wb.create_sheet('Лепестковая')
ws_radar.cell(row=1, column=1, value='Лепестковая диаграмма (радар): численность по типам групп').font = t_font
ws_radar.cell(row=2, column=1, value='Внимание: 16 категорий — диаграмма плотная, смотри в полноэкранном режиме').font = Font(name='Calibri', size=10, color='888888')

ch_radar = RadarChart()
ch_radar.type = 'filled'  # 'standard' or 'filled' or 'marker'
ch_radar.style = 10; ch_radar.width = 22; ch_radar.height = 16
ch_radar.title = 'Численность по типам групп (радар)'

cats_r = Reference(wsd, min_col=1, min_row=3, max_row=last_data)
d_r = Reference(wsd, min_col=2, max_col=3, min_row=2, max_row=last_data)
ch_radar.add_data(d_r, titles_from_data=True)
ch_radar.set_categories(cats_r)
ch_radar.series[0].graphicalProperties.solidFill = '4472C4'
ch_radar.series[0].graphicalProperties.line.solidFill = '2F5496'
ch_radar.series[1].graphicalProperties.solidFill = 'ED7D31'
ch_radar.series[1].graphicalProperties.line.solidFill = 'C55A11'
# Make semi-transparent
from openpyxl.chart.series import DataPoint
from copy import copy

for s in ch_radar.series:
    s.graphicalProperties.line.width = 15000

ws_radar.add_chart(ch_radar, 'A4')

# Сохраняем файл B
out_b = r'C:\Users\user\Downloads\for_Cloude Code\prepord-card\Занятие 3_Open Code\Блок 4\Диаграммы_типы_групп.xlsx'
wb.save(out_b)
print(f'Файл B: {out_b}')

# ═══════════════════════════════════════════════
# ФАЙЛ C: Разбивка на 2 графика
# ═══════════════════════════════════════════════
wb2 = Workbook()

# --- Данные для drastic (исчезли/появились) ---
wsd1 = wb2.active
wsd1.title = 'Исчезли_появились'
wsd1.cell(row=1, column=1, value='Типы групп: исчезли / появились / без изменения').font = t_font
wsd1.cell(row=2, column=1, value='3D, ВМ, ПМ2 — исчезли; ДИ, КМ — появились; АП, ВПР — без изменений').font = Font(name='Calibri', size=10, color='888888')
for c, h in enumerate(['Тип группы', '2024', '2025'], 1):
    wsd1.cell(row=3, column=c, value=h)
hdr(wsd1, 3, 3)
for i, t in enumerate(drastic, 4):
    v24 = t24.get(t, 0); v25 = t25.get(t, 0)
    wsd1.cell(row=i, column=1, value=t).font = d_font
    wsd1.cell(row=i, column=2, value=v24).font = d_font
    wsd1.cell(row=i, column=3, value=v25).font = d_font
    for c in range(1, 4):
        wsd1.cell(row=i, column=c).border = b_all
        if c > 1: wsd1.cell(row=i, column=c).alignment = Alignment(horizontal='center')
    # colour
    if v24 == 0:  # new
        for c in range(1, 4): wsd1.cell(row=i, column=c).fill = PatternFill(start_color='E8F5E9', end_color='E8F5E9', fill_type='solid')
    elif v25 == 0:  # gone
        for c in range(1, 4): wsd1.cell(row=i, column=c).fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    else:  # no change
        for c in range(1, 4): wsd1.cell(row=i, column=c).fill = PatternFill(start_color='FFF9C4', end_color='FFF9C4', fill_type='solid')

last_drastic = 3 + len(drastic)

# График исчезли/появились
ch_drastic = BarChart()
ch_drastic.type = 'col'; ch_drastic.style = 10; ch_drastic.width = 20; ch_drastic.height = 12
ch_drastic.title = 'Исчезли, появились, без изменений'
ch_drastic.legend.position = 'b'
cats_d = Reference(wsd1, min_col=1, min_row=4, max_row=last_drastic)
d_d = Reference(wsd1, min_col=2, max_col=3, min_row=3, max_row=last_drastic)
ch_drastic.add_data(d_d, titles_from_data=True)
ch_drastic.set_categories(cats_d)
ch_drastic.series[0].graphicalProperties.solidFill = '4472C4'
ch_drastic.series[1].graphicalProperties.solidFill = 'ED7D31'
for s in ch_drastic.series:
    s.dLbls = DataLabelList(); s.dLbls.showVal = True; s.dLbls.numFmt = '0'
    s.dLbls.font = Font(size=10)
wsd1.add_chart(ch_drastic, 'A' + str(last_drastic + 2))

# --- Данные для changed (остались с изменением) ---
wsd2 = wb2.create_sheet('Остались')
wsd2.cell(row=1, column=1, value='Типы групп: остались с изменением численности').font = t_font
for c, h in enumerate(['Тип группы', '2024', '2025', 'Δ, чел.', 'Δ, %'], 1):
    wsd2.cell(row=2, column=c, value=h)
hdr(wsd2, 2, 5)
for i, t in enumerate(changed, 3):
    v24 = t24.get(t, 0); v25 = t25.get(t, 0)
    d = v25 - v24; p = d / v24 * 100
    wsd2.cell(row=i, column=1, value=t).font = d_font; wsd2.cell(row=i, column=1).border = b_all
    wsd2.cell(row=i, column=2, value=v24).font = d_font; wsd2.cell(row=i, column=2).border = b_all; wsd2.cell(row=i, column=2).alignment = Alignment(horizontal='center')
    wsd2.cell(row=i, column=3, value=v25).font = d_font; wsd2.cell(row=i, column=3).border = b_all; wsd2.cell(row=i, column=3).alignment = Alignment(horizontal='center')
    wsd2.cell(row=i, column=4, value=d).font = d_font; wsd2.cell(row=i, column=4).border = b_all; wsd2.cell(row=i, column=4).alignment = Alignment(horizontal='center')
    wsd2.cell(row=i, column=5, value=f'{p:+.2f} %').font = d_font; wsd2.cell(row=i, column=5).border = b_all; wsd2.cell(row=i, column=5).alignment = Alignment(horizontal='center')
    if d < 0:
        wsd2.cell(row=i, column=4).font = Font(name='Calibri', bold=True, size=11, color='C00000')
        wsd2.cell(row=i, column=5).font = Font(name='Calibri', bold=True, size=11, color='C00000')

last_changed = 2 + len(changed)

# График оставшихся
ch_changed = BarChart()
ch_changed.type = 'col'; ch_changed.style = 10; ch_changed.width = 20; ch_changed.height = 12
ch_changed.title = 'Оставшиеся типы групп: 2024 vs 2025'
ch_changed.legend.position = 'b'
cats_c = Reference(wsd2, min_col=1, min_row=3, max_row=last_changed)
d_c = Reference(wsd2, min_col=2, max_col=3, min_row=2, max_row=last_changed)
ch_changed.add_data(d_c, titles_from_data=True)
ch_changed.set_categories(cats_c)
ch_changed.series[0].graphicalProperties.solidFill = '4472C4'
ch_changed.series[1].graphicalProperties.solidFill = 'ED7D31'
for s in ch_changed.series:
    s.dLbls = DataLabelList(); s.dLbls.showVal = True; s.dLbls.numFmt = '0'
    s.dLbls.font = Font(size=9)
wsd2.add_chart(ch_changed, 'A' + str(last_changed + 2))

# Сохраняем файл C
out_c = r'C:\Users\user\Downloads\for_Cloude Code\prepord-card\Занятие 3_Open Code\Блок 4\Типы_групп_разбивка.xlsx'
wb2.save(out_c)
print(f'Файл C: {out_c}')
