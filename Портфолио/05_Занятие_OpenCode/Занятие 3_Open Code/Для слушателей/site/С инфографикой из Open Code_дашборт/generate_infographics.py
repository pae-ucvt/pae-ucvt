from PIL import Image, ImageDraw, ImageFont, ImageFilter
import os

W, H = 800, 500
OUT_DIR = os.path.join(os.path.dirname(__file__) or '.', 'infographics')
os.makedirs(OUT_DIR, exist_ok=True)

# === Fonts ===
def get_font(size, bold=False):
    name = 'segoeuib.ttf' if bold else 'segoeui.ttf'
    try:
        return ImageFont.truetype(f'C:/Windows/Fonts/{name}', size)
    except:
        try:
            return ImageFont.truetype('C:/Windows/Fonts/arial.ttf', size)
        except:
            return ImageFont.load_default()

F_BIG  = get_font(52, True)
F_MED  = get_font(24, True)
F_SM   = get_font(18)
F_XS   = get_font(14)
F_KPI  = get_font(36, True)
F_BADGE = get_font(16, True)

def glow_circle(draw, cx, cy, r, color, blur=60):
    g = Image.new('RGBA', (W, H), (0,0,0,0))
    gd = ImageDraw.Draw(g)
    gd.ellipse([cx-r, cy-r, cx+r, cy+r], fill=(*color, 18))
    g = g.filter(ImageFilter.GaussianBlur(blur))
    base.paste(g, (0,0), g)

def round_rect(draw, xy, r, fill, outline=None):
    draw.rounded_rectangle(xy, radius=r, fill=fill, outline=outline)

def card(draw, x, y, w, h, title, items, accent_color, icon_text="", extra=None):
    round_rect(draw, [x, y, x+w, y+h], 14, (22, 28, 56))
    draw.line([x+6, y+10, x+6, y+h-10], fill=accent_color, width=4)
    if icon_text:
        draw.text((x+20, y+16), icon_text, fill=accent_color, font=F_MED)
        tx = x+20
    else:
        tx = x+20
    draw.text((tx, y+18), title, fill=(255,255,255), font=F_MED)
    yy = y+52
    for item in items:
        draw.text((x+24, yy), f"• {item}", fill=(180,190,210), font=F_SM)
        yy += 30
    return yy

def badge(draw, x, y, text, bg):
    tw = draw.textlength(text, font=F_BADGE) + 20
    round_rect(draw, [x, y, x+tw, y+30], 15, bg)
    draw.text((x+10, y+4), text, fill=(255,255,255), font=F_BADGE)

# ============================================================
# INFOGRAPHIC 1 — БЛОК 1: Знакомство с OpenCode
# ============================================================
base = Image.new('RGB', (W, H), (13, 17, 38))
draw = ImageDraw.Draw(base)
glow_circle(draw, 700, 100, 300, (0, 150, 255))
glow_circle(draw, 150, 450, 250, (130, 80, 220))

badge(draw, 30, 20, "БЛОК 1", (0, 150, 255))
draw.text((30, 60), "Знакомство с OpenCode Desktop", fill=(255,255,255), font=F_BIG)
draw.text((30, 120), "Скачать → Установить → Первый запрос → Работа с папкой", fill=(100, 150, 200), font=F_SM)

# Иконка-подложка
round_rect(draw, [30, 160, 370, 460], 14, (22, 28, 56))
draw.text((50, 180), "🧩", fill=(255,255,255), font=get_font(32))
draw.text((90, 184), "Интерфейс", fill=(255,255,255), font=F_MED)
items_iface = ["Слева — файлы проекта", "Окно чата — пишем запросы", "Build / Plan — режимы", "Подтверждение действий", "AGENTS.md — правила проекта"]
yy = 224
for item in items_iface:
    draw.text((50, yy), f"• {item}", fill=(180,190,210), font=F_SM)
    yy += 30

# Правая колонка — первые шаги
round_rect(draw, [400, 160, 770, 280], 14, (22, 28, 56))
draw.text((418, 176), "🚀", fill=(255,255,255), font=get_font(32))
draw.text((458, 180), "Первые запросы", fill=(255,255,255), font=F_MED)
items_q = ['"Привет! Что ты умеешь?"', '"Покажи интерфейс"', '"Сколько времени?"', '"Что ты можешь сделать?"']
yq = 224
for item in items_q:
    draw.text((418, yq), f"• {item}", fill=(180,190,210), font=F_SM)
    yq += 28

round_rect(draw, [400, 300, 770, 460], 14, (22, 28, 56))
draw.text((418, 316), "📁", fill=(255,255,255), font=get_font(32))
draw.text((458, 320), "Открытие папки", fill=(255,255,255), font=F_MED)
items_dir = ['Узнаём текущую папку', 'Открываем рабочую папку', 'Считаем файлы', 'Всё через запросы в чат']
y_dir = 366
for item in items_dir:
    draw.text((418, y_dir), f"• {item}", fill=(180,190,210), font=F_SM)
    y_dir += 28

draw.text((30, 470), "Результат: готовый AGENTS.md + понимание интерфейса OpenCode", fill=(80, 150, 200), font=F_XS)
base.save(os.path.join(OUT_DIR, 'block1_opencode.png'))
print("OK block1_opencode.png")

# ============================================================
# INFOGRAPHIC 2 — БЛОК 2: Порядок в папках
# ============================================================
base = Image.new('RGB', (W, H), (13, 17, 38))
draw = ImageDraw.Draw(base)
glow_circle(draw, 200, 120, 280, (0, 210, 200))
glow_circle(draw, 650, 400, 250, (0, 200, 150))

badge(draw, 30, 20, "БЛОК 2", (0, 210, 200))
draw.text((30, 60), "Порядок в папках", fill=(255,255,255), font=F_BIG)
draw.text((30, 120), "Анализ → Сортировка → Аудит → HTML-дашборд", fill=(100, 200, 190), font=F_SM)

# Левая колонка
card(draw, 30, 160, 230, 290, "Анализ", [
    "Структура папок",
    "Группировка по типам",
    "Размер файлов",
    "Топ-5 больших",
    "Дубликаты"
], (0, 210, 200), "🔍")

card(draw, 285, 160, 230, 290, "Сортировка", [
    "Документы → Документы",
    "Изображения → Изо",
    "Архивы → Архивы",
    "Остальные → Прочее",
    "Удаление дубликатов"
], (0, 200, 150), "📂")

card(draw, 540, 160, 230, 290, "Дашборд", [
    "HTML-дашборд audit.html",
    "Диаграммы по типам",
    "Таблица топ-файлов",
    "Тёмная тема",
    "Готов к демонстрации"
], (0, 150, 255), "📊")

draw.text((30, 470), "Итог: чистая папка + дашборд для презентации руководству", fill=(80, 200, 180), font=F_XS)
base.save(os.path.join(OUT_DIR, 'block2_files.png'))
print("OK block2_files.png")

# ============================================================
# INFOGRAPHIC 3 — БЛОК 3: Анализ текстовых документов
# ============================================================
base = Image.new('RGB', (W, H), (13, 17, 38))
draw = ImageDraw.Draw(base)
glow_circle(draw, 700, 400, 300, (200, 100, 180))
glow_circle(draw, 100, 100, 250, (130, 80, 220))

badge(draw, 30, 20, "БЛОК 3", (200, 100, 180))
draw.text((30, 60), "Анализ текстовых документов", fill=(255,255,255), font=F_BIG)
draw.text((30, 120), "Поиск → Чтение → Сравнение → Сводная таблица", fill=(180, 130, 180), font=F_SM)

# 3 уровня
round_rect(draw, [30, 160, 770, 270], 14, (22, 28, 56))
draw.text((50, 176), "Три уровня задач", fill=(255,255,255), font=F_MED)

lvls = [
    ("Ур.1 — Текстовый ответ", "«Что такое договор поставки?»", (0, 150, 255)),
    ("Ур.2 — Действие с файлом", "«Прочитай договор, выдели стороны и сумму»", (0, 210, 200)),
    ("Ур.3 — Многошаговая задача", "«Найди все договоры → сравни → сделай сводку»", (200, 100, 180)),
]
for i, (title, example, color) in enumerate(lvls):
    y = 212 + i * 38
    draw.rounded_rectangle([50, y, 56, y+26], 4, fill=color)
    draw.text((68, y+1), title, fill=color, font=F_SM)

# Практика
card(draw, 30, 290, 350, 170, "Практика", [
    "Поиск: «договор» или «contract»",
    "Извлечение: стороны, сумма, срок",
    "Сравнение двух документов",
    "Поиск: НДС, пени, штрафы"
], (130, 80, 220), "📝")

card(draw, 410, 290, 350, 170, "Результат", [
    "Сводная таблица по договорам",
    "Статус: действует / истёк",
    "Markdown или Excel-отчёт",
    "Готовый шаблон запроса"
], (0, 150, 255), "✅")

draw.text((30, 470), "Итог: вместо ручного чтения — одна команда и таблица за секунды", fill=(160, 100, 180), font=F_XS)
base.save(os.path.join(OUT_DIR, 'block3_documents.png'))
print("OK block3_documents.png")

# ============================================================
# INFOGRAPHIC 4 — БЛОК 4: Анализ таблиц Excel
# ============================================================
base = Image.new('RGB', (W, H), (13, 17, 38))
draw = ImageDraw.Draw(base)
glow_circle(draw, 700, 100, 300, (255, 170, 50))
glow_circle(draw, 100, 400, 250, (255, 70, 130))

badge(draw, 30, 20, "БЛОК 4", (255, 170, 50))
draw.text((30, 60), "Анализ таблиц Excel", fill=(255,255,255), font=F_BIG)
draw.text((30, 120), "CSV → Чтение → Фильтрация → Итоги → Дашборд", fill=(200, 170, 100), font=F_SM)

card(draw, 30, 160, 350, 190, "Обработка данных", [
    "Конвертация xlsx → CSV",
    "Первые 10 строк — знакомство",
    "Пустые ячейки, типы столбцов",
    "Объединение двух таблиц"
], (255, 170, 50), "📊")

card(draw, 410, 160, 350, 190, "Анализ", [
    "Сумма, среднее по столбцу",
    "Фильтр: > 10 000",
    "Дубликаты по номеру заказа",
    "Группировка по категориям"
], (0, 210, 200), "📈")

card(draw, 30, 370, 730, 90, "Реальный кейс: численность 2024 vs 2025", [
    "Excel-файлы за два года → проверка баланса → сравнение площадок → дашборд с диаграммами и выводами"
], (255, 70, 130), "🏆")

draw.text((30, 470), "Итог: готовый дашборд (Excel + HTML) с анализом и рекомендациями", fill=(200, 150, 80), font=F_XS)
base.save(os.path.join(OUT_DIR, 'block4_excel.png'))
print("OK block4_excel.png")

print("All 4 infographics saved to:", OUT_DIR)
