import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

base = r'C:\Users\user\Downloads\for_Cloude Code\prepord-card\Программа ПК - ИИ для учителей'

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

WHITE = RGBColor(255, 255, 255)
DARK_BLUE = RGBColor(0x1A, 0x3C, 0x6E)
LIGHT_BLUE = RGBColor(0x3A, 0x7C, 0xBD)
GRAY = RGBColor(0x66, 0x66, 0x66)
LIGHT_GRAY = RGBColor(0xF0, 0xF0, 0xF0)

def add_bg(slide, color=DARK_BLUE):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_text_box(slide, left, top, width, height, text, font_size=18, bold=False, color=WHITE, align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = align
    return txBox

def add_bullet_slide(slide, left, top, width, height, items, font_size=16, color=DARK_BLUE):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = item
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.space_after = Pt(6)

# --- SLIDE 1: Title ---
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
add_bg(slide, DARK_BLUE)

add_text_box(slide, 1, 1.5, 11, 1.5, 'Программа повышения квалификации', 36, True, WHITE, PP_ALIGN.CENTER)
add_text_box(slide, 1, 3, 11, 1.5, '«Введение в искусственный интеллект:\nметодика преподавания для учителей-предметников»', 24, False, WHITE, PP_ALIGN.CENTER)
add_text_box(slide, 1, 5.5, 11, 0.5, '36 часов • Очно-заочная с ДОТ • Все предметники 5–9 классов', 14, False, RGBColor(0xAA, 0xCC, 0xEE), PP_ALIGN.CENTER)

# --- SLIDE 2: Overview ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_text_box(slide, 0.5, 0.3, 12, 0.8, 'Структура программы (36 ч)', 28, True, DARK_BLUE)

# Module cards
modules = [
    ('Модуль 1', 'Основы ИИ\nи контекст', '6 ч', '2Л + 3П + 1С'),
    ('Модуль 2', 'УМК «Введение\nв ИИ»', '6 ч', '2Л + 3П + 1С'),
    ('Модуль 3', 'Инструменты ИИ\nв работе учителя', '8 ч', '2Л + 4П + 2С'),
    ('Модуль 4', 'Методика для\nразных предметов', '8 ч', '2Л + 4П + 2С'),
    ('Модуль 5', 'Проектирование\nи AI-аудит', '6 ч', '1Л + 4П + 1С'),
    ('Аттестация', 'Тест +\nзащита', '2 ч', '2П'),
]

for i, (title, subt, hours, form) in enumerate(modules):
    col = i % 3
    row = i // 3
    x = 0.5 + col * 4.3
    y = 1.5 + row * 2.8
    
    # Background rect
    shape = slide.shapes.add_shape(
        1, Inches(x), Inches(y), Inches(3.8), Inches(2.3)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = LIGHT_GRAY if i < 5 else RGBColor(0xE8, 0xF0, 0xFE)
    shape.line.color.rgb = LIGHT_BLUE
    shape.line.width = Pt(1)
    
    add_text_box(slide, x + 0.2, y + 0.1, 3.4, 0.4, title, 14, True, DARK_BLUE)
    add_text_box(slide, x + 0.2, y + 0.5, 3.4, 0.8, subt, 11, False, GRAY)
    add_text_box(slide, x + 0.2, y + 1.4, 3.4, 0.3, hours, 14, True, LIGHT_BLUE)
    add_text_box(slide, x + 0.2, y + 1.8, 3.4, 0.3, form, 10, False, GRAY)

# --- SLIDE 3: AI-audit thread ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, DARK_BLUE)
add_text_box(slide, 0.5, 0.3, 12, 0.8, 'Сквозная линия: AI-аудит', 28, True, WHITE)

threads = [
    ('Модуль 1', '5П-аудит понимания ИИ\n+ проверка с 3 ролей', 'Знакомство с моделью'),
    ('Модуль 2', 'Аудит КТП\nна интеграцию ИИ', 'Применение к планам'),
    ('Модуль 3', 'Совет экспертов (5 ролей)\n+ факты/выводы', 'Методический инструмент'),
    ('Модуль 4', 'Аудит уроков\nпо предметам', 'Предметная адаптация'),
    ('Модуль 5', 'Полный цикл:\n5 шагов аудита → доработка', 'Кульминация'),
]

for i, (mod, what, stage) in enumerate(threads):
    x = 0.5 + i * 2.6
    # Arrow connector
    if i > 0:
        arr = slide.shapes.add_shape(
            5, Inches(x - 0.1), Inches(4.2), Inches(0.3), Inches(0.3)
        )
        arr.fill.solid()
        arr.fill.fore_color.rgb = LIGHT_BLUE
        arr.line.fill.background()
    
    # Step box
    shape = slide.shapes.add_shape(
        1, Inches(x), Inches(1.5), Inches(2.2), Inches(3.5)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0x24, 0x4D, 0x8A)
    shape.line.color.rgb = LIGHT_BLUE
    
    add_text_box(slide, x + 0.1, 1.6, 2.0, 0.4, mod, 14, True, LIGHT_BLUE)
    add_text_box(slide, x + 0.1, 2.0, 2.0, 1.8, what, 11, False, WHITE)
    add_text_box(slide, x + 0.1, 4.0, 2.0, 0.5, stage, 9, False, RGBColor(0xAA, 0xCC, 0xEE))

# --- SLIDE 4: Key numbers ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_text_box(slide, 0.5, 0.3, 12, 0.8, 'Ключевые цифры программы', 28, True, DARK_BLUE)

stats = [
    ('36', 'часов'),
    ('20', 'часов практики'),
    ('5', 'модулей'),
    ('10+', 'промптов для AI-аудита'),
    ('5', 'предметных треков'),
    ('6', 'ролей для проверки'),
]

for i, (num, label) in enumerate(stats):
    col = i % 3
    row = i // 3
    x = 1 + col * 4
    y = 1.8 + row * 2.5
    
    shape = slide.shapes.add_shape(
        1, Inches(x), Inches(y), Inches(3), Inches(1.8)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = LIGHT_GRAY
    shape.line.color.rgb = LIGHT_BLUE
    
    add_text_box(slide, x + 0.3, y + 0.2, 2.4, 0.7, num, 36, True, DARK_BLUE, PP_ALIGN.CENTER)
    add_text_box(slide, x + 0.3, y + 1.0, 2.4, 0.5, label, 14, False, GRAY, PP_ALIGN.CENTER)

# --- SLIDE 5: Target audience ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, DARK_BLUE)
add_text_box(slide, 0.5, 0.3, 12, 0.8, 'Для кого эта программа?', 28, True, WHITE)

groups = [
    ('Информатика\nи технология', 'Прямая интеграция\nмодулей ИИ в КТП'),
    ('Естественно-научный\nцикл', 'ИИ как инструмент\nисследования и анализа'),
    ('Гуманитарный\nцикл', 'Этика ИИ, генерация\nтекста, авторство'),
    ('Математика', 'Данные, машинное\nобучение, статистика'),
    ('Начальная\nшкола', 'Базовые представления\nоб ИИ и технологиях'),
    ('Завуч /\nметодист', 'Проектирование\nКПК и контроль'),
]

for i, (title, desc) in enumerate(groups):
    col = i % 3
    row = i // 3
    x = 0.5 + col * 4.3
    y = 1.5 + row * 2.8
    
    shape = slide.shapes.add_shape(
        1, Inches(x), Inches(y), Inches(3.8), Inches(2.2)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0x24, 0x4D, 0x8A)
    shape.line.color.rgb = LIGHT_BLUE
    
    add_text_box(slide, x + 0.2, y + 0.2, 3.4, 0.8, title, 16, True, WHITE)
    add_text_box(slide, x + 0.2, y + 1.2, 3.4, 0.8, desc, 11, False, RGBColor(0xAA, 0xCC, 0xEE))

# --- SLIDE 6: AI-audit model ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_text_box(slide, 0.5, 0.3, 12, 0.8, 'Модель AI-аудита: 5П', 28, True, DARK_BLUE)

steps = [
    ('Понятность', 'Что неясно?'),
    ('Полнота', 'Чего не хватает?'),
    ('Противоречия', 'Где логика\nломается?'),
    ('Проблемы', 'Что может пойти\nне так?'),
    ('Предложения', 'Как исправить?'),
]

for i, (title, desc) in enumerate(steps):
    x = 0.5 + i * 2.6
    y = 2.0
    
    shape = slide.shapes.add_shape(
        1, Inches(x), Inches(y), Inches(2.2), Inches(2.2)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = LIGHT_GRAY if i % 2 == 0 else RGBColor(0xE8, 0xF0, 0xFE)
    shape.line.color.rgb = LIGHT_BLUE
    shape.line.width = Pt(1)
    
    add_text_box(slide, x + 0.1, y + 0.3, 2.0, 0.5, title, 20, True, DARK_BLUE, PP_ALIGN.CENTER)
    add_text_box(slide, x + 0.1, y + 1.0, 2.0, 1.0, desc, 12, False, GRAY, PP_ALIGN.CENTER)
    
    conn = f'Шаг {i+1}'
    add_text_box(slide, x + 0.6, y + 0.1, 1.0, 0.3, conn, 9, False, LIGHT_BLUE, PP_ALIGN.CENTER)

add_text_box(slide, 0.5, 5.0, 12, 0.5, 'Формула промпта: Роль → Материал → Критерии (5П) → Формат → Ограничение', 14, True, GRAY, PP_ALIGN.CENTER)

# --- SLIDE 7: 5 roles ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, DARK_BLUE)
add_text_box(slide, 0.5, 0.3, 12, 0.8, '«Совет экспертов» — 5 ролей для проверки', 28, True, WHITE)

roles = [
    ('👤 Ученик', 'Что непонятно?\nГде скучно?'),
    ('📋 Завуч', 'Результаты?\nРесурсы?'),
    ('❓ Новичок', 'Какие шаги\nпропущены?'),
    ('⚠ Скептик', 'Почему не\nсработает?'),
    ('✏ Редактор', 'Где вода\nи штампы?'),
]

for i, (role, question) in enumerate(roles):
    x = 0.5 + i * 2.6
    y = 2.0
    
    shape = slide.shapes.add_shape(
        1, Inches(x), Inches(y), Inches(2.2), Inches(2.5)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0x24, 0x4D, 0x8A)
    shape.line.color.rgb = LIGHT_BLUE
    
    add_text_box(slide, x + 0.1, y + 0.2, 2.0, 0.6, role, 16, True, WHITE, PP_ALIGN.CENTER)
    add_text_box(slide, x + 0.1, y + 1.0, 2.0, 1.2, question, 11, False, RGBColor(0xAA, 0xCC, 0xEE), PP_ALIGN.CENTER)

add_text_box(slide, 0.5, 5.5, 12, 0.5, 'Один материал — 5 точек зрения. Сначала диагностика, потом улучшение.', 14, False, RGBColor(0xAA, 0xCC, 0xEE), PP_ALIGN.CENTER)

# --- SLIDE 8: Final slide ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, DARK_BLUE)

add_text_box(slide, 1, 2, 11, 1, 'Итог', 36, True, WHITE, PP_ALIGN.CENTER)
add_text_box(slide, 1, 3.3, 11, 1.5, 'Программа КПК содержит готовые промпты\nдля каждого модуля и формирует навык AI-аудита\nкак регулярную практику учителя', 18, False, RGBColor(0xAA, 0xCC, 0xEE), PP_ALIGN.CENTER)
add_text_box(slide, 1, 5.5, 11, 0.5, 'Сначала диагностика → Потом улучшение', 14, False, WHITE, PP_ALIGN.CENTER)

# Save
pptx_path = os.path.join(base, 'Презентация_для_тьюторов_КПК.pptx')
prs.save(pptx_path)
print(f'PPTX created: {pptx_path}')
print(f'Size: {os.path.getsize(pptx_path)} B, Slides: {len(prs.slides)}')
