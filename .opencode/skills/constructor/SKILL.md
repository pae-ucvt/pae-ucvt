---
name: constructor
description: >
  Use ONLY when the user wants to create a new educational program (ДПП ПК) using the Constructor:
  generating programs from JSON, filling templates, converting to .docx, or working with program data.
  Do NOT use for creating lesson plans, tests, or other documents outside the Constructor.
---

# Конструктор учебных программ

## Назначение

Инструмент для быстрого создания программ повышения квалификации педагогов в формате .docx.
Поддерживает объёмы 36 часов (9 занятий) и 72 часа (18 занятий) по 4 часа в день.

## Быстрый старт

### Способ 1: Через JSON (рекомендуется)

1. Скопируйте шаблон:
   ```
   cp "Выпускной проект/Нейросети в работе преподавателя/neural_networks.json" my_program.json
   ```

2. Отредактируйте JSON — замените данные на свои:
   - `name` — название программы
   - `goal` — цель
   - `audience` — категория слушателей
   - `hours` — 36 или 72
   - `lessons` — список занятий с темами и вопросами

3. Запустите генерацию:
   ```
   python scripts/build_program.py my_program.json
   ```

4. Готовый .docx появится рядом с JSON-файлом

### Способ 2: Через Markdown-шаблон

1. Откройте `templates/constructor_template.md`
2. Замените все `{{PLACEHOLDER}}` на реальные значения
3. Скопируйте содержимое в отдельный файл
4. Сгенерируйте .docx командой:
   ```
   python scripts/convert_to_docx.py "my_program.md"
   ```

## Структура JSON-файла

```json
{
  "name": "Название программы",
  "city": "Город",
  "year": 2026,
  "goal": "Цель программы",
  "audience": "Категория слушателей",
  "hours": 36,
  "schedule_format": "4 академических часа в день",
  "duration": "9 недель (1 занятие в неделю)",
  "lessons": [
    {
      "num": 1,
      "title": "Название занятия",
      "hours": 4,
      "lectures": 2,
      "practice": 2,
      "control": "Устный опрос",
      "topics": [
        {
          "num": "1.1",
          "title": "Название темы",
          "hours": 2,
          "lectures": 1,
          "practice": 1,
          "content": "Описание содержания темы"
        }
      ],
      "questions": [
        "Вопрос 1?",
        "Вопрос 2?",
        "Вопрос 3?"
      ]
    }
  ],
  "purpose": "Краткое описание цели",
  "tasks": ["Задача 1", "Задача 2"],
  "relevance": "Актуальность программы",
  "uniqueness": ["Особенность 1", "Особенность 2"],
  "professional_standard": {
    "description": "Описание связи с профстандартом",
    "table": [
      ["Трудовая функция", "Трудовые действия", "Номера занятий"]
    ]
  },
  "practical_value": "Практическая значимость",
  "structure": "Описание структуры программы",
  "conditions": ["Условие 1", "Условие 2"],
  "must_know": ["Знание 1", "Знание 2"],
  "must_be_able": ["Умение 1", "Умение 2"],
  "current_control": "Текущий контроль",
  "final_assessment": "Итоговая аттестация",
  "portfolio": ["Элемент 1", "Элемент 2"],
  "criteria": [
    {"text": "Критерий отличного результата", "grade": "отлично"},
    {"text": "Критерий хорошего результата", "grade": "хорошо"}
  ],
  "regulatory_docs": ["Документ 1", "Документ 2"],
  "internet_resources": ["Ресурс 1", "Ресурс 2"],
  "output_dir": "."
}
```

## Как заполнить поля

| Поле | Тип | Обязательно | Описание |
|------|-----|-------------|----------|
| name | строка | Да | Название программы в кавычках |
| city | строка | Да | Город |
| year | число | Да | Год (2026) |
| goal | строка | Да | Одно предложение: для чего программа |
| audience | строка | Да | Кто обучается |
| hours | число | Да | 36 или 72 |
| lessons | массив | Да | Список занятий (9 шт. для 36ч, 18 шт. для 72ч) |
| lessons[num] | число | Да | Порядковый номер занятия |
| lessons[title] | строка | Да | Название занятия |
| lessons[hours] | число | Да | Часов на занятие (всегда 4) |
| lessons[topics] | массив | Да | Темы внутри занятия (2 темы по 2 ч) |
| lessons[questions] | массив | Да | Контрольные вопросы (мин. 3) |
| professional_standard | объект | Нет | Связь с профстандартом (если есть) |

## Формулы для автоматического расчёта

При заполнении JSON используйте:
- `lessons[*].lectures + lessons[*].practice = lessons[*].hours` (должно сходиться)
- `sum(lessons[*].hours) = data.hours` (общий объём)
- `lessons[*].topics[*].hours = 2` (стандартное распределение: 2 темы по 2 ч)

## Частые ошибки

1. **Не сходятся часы** — проверьте, что сумма по занятиям равна общему объёму
2. **Нет контрольных вопросов** — после каждого занятия минимум 3 вопроса
3. **Пустые обязательные поля** — purpose, tasks, must_know, must_be_able обязательны
4. **Неправильный формат чисел** — hours: 36 (без кавычек), не "36"

## Структура файлов

```
methodologist/
├── .opencode/
│   ├── agents/methodologist.md          # Конфигурация агента
│   └── skills/constructor/SKILL.md      # Этот файл
│
├── scripts/
│   ├── build_program.py                 # Генерация .docx из JSON
│   ├── convert_to_docx.py               # Конвертация .md → .docx
│   ├── gen_presentation.py              # Генерация презентаций
│   └── gen_summary_docx.py              # Генерация сводок
│
├── templates/
│   └── constructor_template.md          # Шаблон с инструкцией
│
├── programs/
│   ├── Промпт-инжиниринг для педагога/  # Готовые программы
│   └── Цифровые инструменты оценивания в школе/
│
└── Выпускной проект/
    └── Нейросети в работе преподавателя/ # Пример JSON (36ч)
```

## Зависимости

```bash
pip install python-docx openpyxl
```

## Решение проблем

| Проблема | Решение |
|----------|---------|
| `ModuleNotFoundError: No module named 'docx'` | `pip install python-docx` |
| `FileNotFoundError` | Проверьте путь к JSON-файлу |
| Пустой .docx | Проверьте, что JSON заполнен корректно |
