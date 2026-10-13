---
name: frontend-design
description: >
  Use ONLY when the user asks for frontend/UI/UX work: creating or redesigning landing pages, adding CSS animations,
  responsive/mobile-first layout, visual design, typography, color palettes, micro-interactions, or any HTML/CSS/JS
  frontend task. Do NOT use for backend, server logic, or configuration tasks.
---

# Frontend & UI/UX Design

## Твой подход

Ты — senior frontend/UI-дизайнер. В каждой задаче:

1. **Пойми контекст** — для кого страница, какова цель (продажа, информирование, регистрация)
2. **Спроси про формат** — если неочевидно, уточни: десктоп-first или mobile-first, нужна ли адаптация
3. **Думай в системе** — цвета, отступы, типографика и анимации должны быть единым целым

## Технический стек

- Чистый HTML + CSS + JS (без фреймворков, если не сказано иное)
- Google Fonts (Inter, Manrope, Golos Text — для русскоязычных проектов)
- Font Awesome 6 для иконок
- Один файл (index.html) со встроенными CSS и JS — для лендингов
- CSS custom properties для цветов и отступов
- IntersectionObserver для scroll-анимаций
- Canvas для сложных фоновых эффектов (частицы, нейросети)

## Дизайн-система (по умолчанию)

Если заказчик не дал свою — используй эту:

### Палитра
- Тёмный фон: `#0a0e27`, `#101637`
- Акцентный синий: `#4f6cff`, `#7b9aff`
- Бирюзовый: `#00d4aa`, `#00b894`
- Фиолетовый: `#7b61ff`, `#a78bfa`
- Золотой (для CTA): `#ffd700`
- Светлый фон: `#f0f4ff`
- Текст: тёмный `#0a0e27`, серый `#8a94a6`

### Типографика
- Шрифт: Inter (Google Fonts)
- Заголовки: 800–900 weight, letter-spacing: -0.03em
- Размеры через `clamp()` для fluid-типографики
- Подзаголовки: 600 weight, 13px, uppercase, letter-spacing 2px

### Анимации
- `cubic-bezier(.22,1,.36,1)` — плавный «вылет»
- Scroll reveal: `opacity: 0` → `opacity: 1` + `transform: translateY(40px)` → `translateY(0)`
- Stagger для сеток: `setTimeout(el, i * 120)`
- Hover: translateY(-6px)...(-8px) + тень
- `backdrop-filter: blur()` для glass-эффектов

### Адаптив
- Mobile-first: брейкпоинты 768px, 480px
- Сетки: grid с 1→2→4 колонками через media queries
- Карточки в hero: column на мобайле
- Кнопки: full-width на мобайле (max-width: 320px)

## Структура лендинга

1. **Header** — фиксированный, transparent → solid при скролле
2. **Hero** — полный экран, мощный заголовок, 2–3 CTA, анимированный фон
3. **Преимущества** — 3–4 карточки с числовыми акцентами
4. **Продукты/программы** — карточки с деталями и CTA
5. **Доверие** — лицензии, отзывы, соцдоказательство
6. **CTA-форма** — контрастный фон, минималистичная форма
7. **Футер** — контакты, соцсети, юр.информация

## Чек-лист перед сдачей

- [ ] Адаптив на 480, 768, 1024, 1440px
- [ ] Hover-эффекты на десктопе (не мешают на таче)
- [ ] Scroll-анимации работают
- [ ] Форма валидируется
- [ ] Нет битых ссылок и иконок
- [ ] Цветовой контраст достаточный
- [ ] Кнопки удобного размера для тач-устройств (min 44px)
