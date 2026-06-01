---
name: cloud-ru-slides
description: Use this skill when the user wants to format, create, or audit a PowerPoint presentation according to Cloud.ru 2.0 brand standards. Activates for tasks involving .pptx files for Cloud.ru, slide layout design, brand guidelines compliance checks, or when the user mentions "свёрстай презентацию", "Cloud.ru шаблон", "брендбук Cloud", or attaches a Cloud.ru .pptx draft.
---

# Cloud.ru Slides — AI Layout Agent

Skill для автоматической вёрстки PowerPoint-презентаций по бренду **Cloud.ru 2.0**. Поддерживает три сценария:

1. **Драфт → Вёрстка:** пользователь даёт сырой `.pptx` (текст + черновая структура) → возвращаем свёрстанный по шаблону
2. **Бриф → Презентация:** пользователь даёт текст/markdown → создаём `.pptx` с нуля по шаблону
3. **Аудит:** пользователь даёт готовый `.pptx` → возвращаем отчёт о соответствии бренду + предложения

## Обязательное чтение перед работой

Открывай эти файлы при активации:

1. `brand/brand-rules.md` — компактные правила бренда (цвета, шрифты, сетка, иерархия, запреты)
2. `brand/template-analysis.md` — семантический каталог 102 layouts шаблона
3. `brand/template-layouts-dump.json` — машинный дамп с координатами placeholder'ов

## Что нужно от пользователя

- **Обязательно:** черновик `.pptx` или markdown-бриф
- **Шаблон `Cloud.ru_Template_2026.pptx`** (29MB, не входит в скилл) — подтягивается **автоматически** через `scripts/template_path.py`. Подгружать в чат больше не нужно, если файл есть на диске. Путь настраивается в `brand/template-version.json → template_path` (или env `CLOUD_RU_TEMPLATE`). Загрузка в чат / в текущую папку остаётся fallback'ом, если файл не найден.
- **Опционально:** референсные слайды (PNG), брендбук (если нужны нестандартные правила)

> Получить путь к шаблону в любой команде: `TPL=$(python3 scripts/template_path.py)` — резолвер сам найдёт файл (env → config → известные пути) или подскажет, что делать.

> ⚠️ **Если шаблон не найден** (`template_path.py` вернул ошибку / непустой stderr) — **НЕ продолжать вёрстку**. Остановиться и попросить пользователя: «Не вижу шаблон `Cloud.ru_Template_2026.pptx` на диске — загрузи его в чат или положи в рабочую папку, и я продолжу.» Без шаблона финальную сборку делать нельзя (это происходит, например, в облаке claude.ai, где диск десктопа недоступен).

## Pipeline v0.17 — Per-slide LLM Design Loop

**Главный принцип:** Скилл = LLM-дизайнер, не batch скрипт. Каждый слайд проходит **независимый цикл** проектирования с визуальной обратной связью.

### Workflow

```
PHASE 1 — GLOBAL ANALYSIS (один раз):
  [G1] parse_pptx / parse_md / parse_docx / extract_images
  [G2] Brief Reader — общий бриф презентации (агент 01)
  [G3] Slide Classifier — категория для каждого слайда (агент 02)

PHASE 2 — PER-SLIDE DESIGN LOOP (для каждого слайда независимо):
  ┌── for slide in classified.slides:
  │      [S1] Layout Designer (агент 04) — выбирает donor с design thinking
  │           — image как контент или декор?
  │           — нужны remove_shapes для PNG-stripping?
  │           — image_zone для main-content картинок
  │
  │      [S2] Content Distributor (агент 03) — распределяет text по slots
  │
  │      [S3] Copy Editor (агент 07) — kill_widows.py + анти-эмодзи
  │
  │      [S4] Local validate — проверить slot fits, overflow strategies
  │
  │      [S5] Build SINGLE slide через build_v8.py (с pre-cleanup)
  │
  │      [S6] Render single slide → PNG
  │
  │      [S7] LLM Visual Verifier (агент 10) — Я (Claude) смотрю PNG:
  │           ✓ Текст не закрыт декорами?
  │           ✓ Картинка как задумано (контент vs декор)?
  │           ✓ Иерархия читаема?
  │           ✓ Композиция сбалансирована?
  │           ✓ Сравнить с эталоном template/png/ той же категории
  │
  │      [S8] If verdict ≠ PERFECT:
  │           → Fix one issue (другой donor / pre-cleanup / size override)
  │           → Goto [S5]
  │      Max 3 iterations per slide
  └── end for

PHASE 3 — GLOBAL VERIFICATION (один раз в конце):
  [V1] Brand Guardian — палитра/шрифты/композиция (агент 08)
  [V2] Visual Validator (PIL) — aspect ratio, empty, edges
  [V3] Process Verifier (агент 09) — собрать все verdicts → READY/NEEDS_REWORK
```

### Ключевые отличия от v0.9 batch pipeline

| | v0.9 batch | v0.17 per-slide |
|---|---|---|
| Структура | Все слайды разом | Слайд за слайдом |
| Visual feedback | В конце | После КАЖДОГО слайда |
| Итерации | Re-run всё | Re-do один слайд |
| LLM роль | Декоратор | **Дизайнер** |
| Image | Прикрепить файл | **Контент vs декор решение** |
| Table | Add new | **Fill existing brand-styled** |

## Альтернативный путь — batch pipeline (v0.9 legacy)

Для не-критичных кейсов можно batch:

```
[0] python3 scripts/parse_pptx.py <draft.pptx> output/parsed_draft.json
    + scripts/extract_images.py — извлекает PNG/JPG в /draft_images_<name>/

[1] Brief Reader (agents/01-brief-reader.md)
    Читает parsed_draft.json → output/brief.json
    {topic, audience, key_messages, tone, slide_count, has_numbers, has_quotes}

[2] Slide Classifier (agents/02-slide-classifier.md)
    Для каждого слайда → category + subcategory
    Output: output/classified.json

[3] Layout Designer (agents/04-layout-designer.md) — v0.9 ALGORITHM
    Загрузить brand/donor-slot-map.yaml
    Для каждого слайда:
      - Кандидаты donors из tone_groups
      - Выбрать первого donor где len(text) ≤ slot.safe_max_chars для всех слотов
      - Если overflow — apply STRATEGY 1/2/3/4 (см. brand-rules.md §14b)
      - ОБЯЗАТЕЛЬНО для KPI 43/44: задать size_pt=199 для num_* слотов
    Output: output/plan_draft.json (с slot_styles_override)

[4] Content Distributor (agents/03-content-distributor.md)
    Распределить тексты по слотам выбранного donor
    Output: output/plan.json (полный plan для build)

[5] Copy Editor (agents/07-copy-editor.md)
    kill_widows.py — типографика (nbsp, тире, ёлочки)
    Удалить эмодзи (заменить на иконки/маркеры)

[6] python3 scripts/validate_plan.py output/plan.json brand/donor-slot-map.yaml output/plan_validated.json
    Auto-добавляет canonical_color/canonical_size_pt/canonical_bold
    Auto-применяет STRATEGY 3 (size reduction 20-25%) при overflow ≤30%
    FAIL если overflow >30% — нужен другой donor / split (вернуться к [3])

[7] TPL=$(python3 scripts/template_path.py)   # авто-резолв шаблона
    python3 scripts/build_v6.py output/plan_validated.json "$TPL" output/result.pptx

[8] python3 scripts/render_slides.py output/result.pptx output/render/

[9] python3 scripts/brand_guardian.py output/result.pptx output/brand_report.json
    Проверяет: цвета (palette + tolerance 50), шрифты (SB Sans*), размеры ≥10pt,
    overflow heuristic, эмодзи, composition (title pressed, subtitle green, KPI consistency)
    Exit 0=PASS, 1=WARN, 2=FAIL

[10] python3 scripts/visual_validator.py output/render/ output/visual_report.json
     Пиксельный анализ: aspect ratio, empty slide, edge overflow (bg-aware), palette dominance

[11] LLM Visual Verifier (agents/10-llm-visual-verifier.md)  ← ОБЯЗАТЕЛЬНЫЙ финальный шаг
     LLM ЧИТАЕТ каждый PNG-рендер и проверяет:
       - Текст заменился (не placeholder донора)
       - Цифры/KPI не обрезаны
       - Composition / overflow внутри слайда
       - Контраст, читаемость, бренд-tone
     ЭТО критично — все автоматические валидаторы могут давать PASS,
     но LLM увидит то, что они физически не могут (логика, семантика).

[12] Process Verifier (agents/09-verifier.md)
     Финальный 10-step checklist + объединение всех verdict-ов:
       validate_plan + brand_guardian + visual_validator + llm_visual_verifier
     READY или NEEDS_REWORK с конкретными next_actions
```

Для **аудита** работают только: parse → classify → guardian → verifier (без build).

## Скрипты v0.20

| Скрипт | Назначение |
|---|---|
| `parse_pptx.py` | .pptx → JSON структуры (slides + text + images) |
| `extract_images.py` | Извлекает PNG/JPG из draft.pptx в файлы |
| **`build_v9.py`** | **JSON plan + шаблон → .pptx, поддержка native slide_type (kpi_native / image_native / chart_pptx_native / chart_native)** |
| `validate_plan.py` | Gate перед build: auto-add canonical, auto-apply overflow strategy. **v0.20: поддержка native slide_types** |
| **`kpi_renderer.py`** | **Рисует KPI shapes на blank canvas (3 числа равного размера; цифры всегда `#222222`, accent = зелёная плашка-подчёркивание, не зелёный текст — Problem #2)** |
| **`image_renderer.py`** | **Image-as-content auto-fit (mode: fit/fill, опц. caption в серой плашке)** |
| **`chart_native_pptx.py`** ⭐ | **DEFAULT для charts (v1.4+). Редактируемая PowerPoint chart через `pptx.chart.add_chart()` — пользователь меняет данные через Edit Data → Excel. Поддержка: area_stacked/area_100/bar/bar_stacked/line/pie** |
| `chart_engine.py` | Legacy: Matplotlib chart redraw в PNG (canonical pastel palette). Использовать ТОЛЬКО когда нужны custom annotations или прозрачные overlapping areas, которые native chart не умеет |
| **`flow_renderer.py`** ⭐ | **v1.6+. Редактируемые схемы / блок-диаграммы / process maps через примитивы PowerPoint (blocks, arrows, dashed groups). slide_type: flow_diagram_native** |
| **`table_renderer.py`** ⭐ | **v1.8+. Редактируемые native PowerPoint таблицы (slide.shapes.add_table) в zebra-стиле slide 56 шаблона. slide_type: table_native** |
| `render_slides.py` | .pptx → PNG (LibreOffice) для визуальной проверки |
| `brand_guardian.py` | Валидация цветов/шрифтов/композиции готового .pptx |
| **`visual_validator_v2.py`** | **Pixel-level анализ rendered PNG (PIL): unfilled placeholders, off-palette > 15%, dominant bg** |
| `kill_widows.py` | Типографика (nbsp, тире, ёлочки) |

## Native slide types

Кроме стандартного `clone_from_slide` в plan можно указать `slide_type`:

```json
{"slide_type": "kpi_native", "dark": false, "kpi": {
  "title": "...",
  "numbers": [{"value": "12", "desc": "...", "pct": true, "accent": false}, ...]
}}

{"slide_type": "image_native", "image": {
  "title": "...", "image_path": "path/to/img.png", "caption": "..."
}}

// ⭐ DEFAULT для всех графиков и диаграмм (v1.4+) — РЕДАКТИРУЕМАЯ PowerPoint chart.
// Пользователь жмёт ПКМ по диаграмме → «Изменить данные» → открывается Excel.
{"slide_type": "chart_pptx_native", "dark": false, "chart": {
  "type": "area_100|area_stacked|bar|bar_stacked|line|pie",
  "title": "...", "caption": "...",
  "x": [...], "series": [{"name": "...", "data": [...]}, ...],
  "accent_idx": N  // индекс серии под GREEN hero
}}

// Legacy: matplotlib → PNG. Только для случаев когда native chart не справляется
// (custom annotations, прозрачные overlapping). НЕ редактируется в PowerPoint.
{"slide_type": "chart_native", "chart": {
  "type": "area_stacked|bar|line|pie",
  "slide_title": "...", "x": [...], "series": [...], "accent_idx": N
}}

// ⭐ v1.8+ Редактируемая native PowerPoint таблица в zebra-стиле (slide 56).
// Пользователь может в PowerPoint двигать колонки, добавлять строки кнопкой,
// редактировать данные через таблицу. Только для регулярных таблиц
// (≥3 cols × ≥3 rows, БЕЗ merged cells). Если есть merged cells / irregular
// grid → anti-distortion stop+ask (см. feedback_anti_distortion_safety.md).
{"slide_type": "table_native", "dark": false, "table": {
  "header": "Сравнение тарифов",
  "subtitle": "опц.",
  "style": "zebra",
  "headers": ["Тариф", "Старт", "Бизнес", "Энтерпрайз"],
  "data": [
    ["Цена/мес", "10 000 ₽", "50 000 ₽", "По запросу"],
    ["Лимит API", "10K", "100K", "Безлимит"],
    ["SLA", "99.5%", "99.9%", "99.99%"]
  ],
  "first_col_wider": true,
  "borders": {                           // опц., гибкое управление границами
    "vertical": true,                     // внутренние вертикали (default true)
    "horizontal": false,                  // внутренние горизонтали (default false)
    "outer_top": false,                   // внешняя верхняя (default false)
    "outer_bottom": false,                // внешняя нижняя (default false)
    "outer_left": false,                  // внешняя левая (default false)
    "outer_right": false,                 // внешняя правая (default false)
    "color": "#434343",                   // default #434343
    "width_pt": 1.0                       // default 1.0
  }
}}

// ⭐ v1.6+ Редактируемые схемы / блок-диаграммы / process maps.
// Блоки + стрелки + опц. пунктирные группы + декор. Все элементы редактируемые
// в PowerPoint (можно подвинуть блок, переписать текст, переподключить стрелку).
{"slide_type": "flow_diagram_native", "dark": false, "flow": {
  "header": "Заголовок схемы",
  "subtitle": "опц.", "subtitle_url": "опц.",
  "blocks": [
    {"id": "b1", "x": 175, "y": 180, "w": 235, "h": 50,
     "lines": ["Title"], "font_sizes": [13], "bolds": [true]}
  ],
  "arrows": [
    {"from": "b1", "to": "b2", "side": "right"},
    {"x1": 100, "y1": 200, "x2": 300, "y2": 200,
     "with_head": true, "dashed": false}
  ],
  "groups": [{"label": "Phase 1", "x": 167, "y": 154, "w": 251, "h": 86}],
  "labels": [{"x": 35, "y": 122, "w": 600, "h": 22,
              "text": "подпись", "font_size": 11, "align": "left"}],
  "decor": {"enabled": true, "x_start": 950, "y_start": 625,
            "count": 4, "size": 38, "gap": 12}
}}
```

**Декор-стрелки (Problem #4, 2026-05-29):** `decor` рисует ряд фирменных
зелёных **стрелок ↗** (древко вверх-вправо + уголок-наконечник, геометрия
`brand/icons/brand_arrow.svg`). Каждая стрелка — **нативная редактируемая фигура
PowerPoint** (группа линий `p:grpSp`): сразу двигается/перекрашивается/
реформируется, без шага «Преобразовать в фигуру». Толщина линии **1pt**.
(Утилита `add_svg_picture` для вставки SVG-вектора в слайд тоже доступна —
например для логотипов/иконок — но декор-стрелки делаются нативными фигурами.)

### Canonical правило v2.3 — пресеты дизайн-архетипов (2026-06-01)

Готовые композиции из эталонных дек (`brand/design-principles-from-decks.md`),
чтобы драфт сразу собирался на уровне финального дизайна. Задаются одним полем
`preset` — рендерер сам считает раскладку, кегли и фирменный зелёный мотив.

**Сверено по референсам 2026-06-01.** Единая 3-уровневая иерархия во всех
пресетах: **заголовок** — SemiBold графит `#222222`; **тело** — regular серым
`TEXT_GRAY #5C5C5C` (на тёмной плашке — светло-серое `#CFCFCF`); номера — графит
**regular** (НЕ bold, НЕ белый). Отступы между фреймами `PRESET_GAP=4` (≤10px).
**Отступ под подзаголовок:** при наличии `subtitle`/`subtitle_url` тело пресета
автоматически стартует ниже (`content_top` 172 / 190 px), чтобы не прижиматься к
подзаголовку; без подзаголовка — с `SAFE_TOP=140`. Можно задать `content_top` явно.

**flow_diagram_native** (`flow.preset`):

```json
// 5B — нумерованные строки. Чистый стиль (default): номер графитом + заголовок
//      SemiBold в одну строку, тело серым ниже. Без лент и чипов. cols=2 для 6–8 пунктов.
{"slide_type":"flow_diagram_native","flow":{"header":"Почему Cloud.ru",
  "preset":"numbered_rows","cols":2,
  "rows":[{"title":"Time-to-Market","text":"описание"}, ...]}}
//      Вариант "band": серые ленты во всю ширину + зелёный чип-номер слева —
//      добавь "style":"band" (для коротких чеклистов).

// 4 — сетка карточек. Заголовок SemiBold + тело серым. Чип ОПЦИОНАЛЕН:
//      "check":true → зелёный чип с иконкой (кружок+галочка графитом),
//      "number":"01" → чип с графитовой цифрой. Чип+заголовок в ОДНУ строку,
//      тело ниже. Без чипа — заголовок сверху, тело под ним.
{"flow":{"header":"...","preset":"card_grid","cols":3,
  "cards":[{"title":"...","text":"...","check":true}, ...]}}

// 5A — открытые нумерованные колонки. Без черты сверху (default), тело серым,
//      внизу крупный (~56pt) REGULAR номер (зелёный). "rule":true вернёт черту.
{"flow":{"header":"...","preset":"numbered_columns",
  "columns":[{"number":"01","title":"...","text":"..."}, ...]}}

// 6 — hero-утверждение (крупный CAPS-текст на зелёной плашке + контурные рамки-декор)
{"flow":{"header":"...","preset":"hero_statement",
  "statement":"короткий мощный тезис","support":"опц. поддержка снизу-справа"}}
```

**table_native** (`table.preset`):

```json
// 8 — before/after. Три колонки: метрика | «Было» | «Стало». Заголовки-табы
//      ЦВЕТНЫЕ (серый / зелёный), текст по ЛЕВОМУ краю во всю ширину колонки.
//      Лёгкая серая подложка под «Было», тонкие серые разделители строк.
//      Метрика двухстрочная: "Название\nдеталь" → название SemiBold + деталь серым.
{"slide_type":"table_native","table":{"header":"On-prem vs Cloud","preset":"before_after",
  "metric_label":"Критерий","before_label":"On-prem","after_label":"Cloud.ru",
  "rows":[{"metric":"Время","before":"Месяцы на закупку","after":"Минуты"}, ...]}}
```

**Зелёный чип-мотив** (`add_green_chip`, правило A-10): зелёный квадрат ~44–50px с
ГРАФИТОВОЙ цифрой `01`–`05` (regular) ИЛИ иконкой-галочкой (**контурный кружок +
галочка графитом**) — повторяющийся акцент слева. Используется в `card_grid`
(check/number) и `numbered_rows` (style=band). **Контент на зелёном — графит, НЕ
белый** (white-on-green запрещён).

Ещё не реализованы (рисовать вручную по принципам): архетип 1 (матрица сравнения
N×M), 2 (карточка-персона с фото), 3 (flow с кодировкой узлов по заливке),
7 (код-панель).

### Canonical правило v2.1 — грид-композиция схем + branching + панели (2026-05-29)

Композиция схем (раньше падала в overflow/разнобой, т.к. блоки хардкодились):

- **Грид-режим** (`flow.grid=true`): блоки задаются логически (`row`/`col`/`lines`/
  `fill`) — `compose_grid` сам считает раскладку: колонки равной ширины, **высота
  строки = под текст** (frame-to-text), **единый кегль** (`font_size`, дефолт 16),
  единые зазоры. Никакого сжатия шрифта (autofit отвергнут — даёт разнобой кеглей).
- **Стрелки по ячейкам** `{"from":[r,c],"to":[r,c]}`: направление авто (строка→
  горизонталь, колонка→вертикаль). **Branching** (пайплайн → N выходов): диагональ →
  **ортогональный Z-маршрут** (`add_orthogonal_arrow`, вправо→верт→вправо) с резервом
  под наконечник (`ARROW_ENTRY_RESERVE`); для веток зазор колонок больше (`GRID_GAP_BRANCHING`).
- **Залитые панели** для нагруженных схем (`groups:[{"style":"panel",...}]`,
  `add_filled_panel`): серая секция-фон + заголовок, блоки внутри `fill:"white"`.
  Меньше пунктира = чище. Пунктир (`add_dashed_rect`) — только для лёгких схем.

### Canonical правило v1.4 — editable charts

**Все диаграммы и графики должны быть редактируемыми (`chart_pptx_native`).**
Это позволяет пользователю менять данные прямо в PowerPoint без перевыпуска
файла. Использование `chart_native` (PNG-картинка) — только когда native
PowerPoint chart не может реализовать необходимый эффект.

**Header chart-слайда** — в стиле content slides шаблона: 20pt SemiBold CAPS,
top-left (35, 60). НЕ 32pt и НЕ без CAPS (должно совпадать визуально с
заголовками donor 21/29).

**Палитра:** `accent_idx` получает GREEN, остальные серии — pastel из
`NON_ACCENT_COLORS` (graphite, pastel blue, pastel yellow, pastel mint, gray).
GREEN никогда не дублируется на не-accent серии.

**Автоматическая детекция:** агент 02 (Slide Classifier) определяет `slide_type` по контенту:
- KPI цифры → `kpi_native`
- chart-like data (series, axis, time-data) → `chart_pptx_native` (DEFAULT)
- main image (фото/скриншот) → `image_native`
- схема / процесс / архитектура (блоки+стрелки) → `flow_diagram_native` ⭐ (v1.6+)
- регулярная таблица ≥3×3 с шапкой, БЕЗ merged cells → `table_native` ⭐ (v1.8+)
- **anti-distortion триггер** (merged cells, RACI, color-coded, brand-объекты) → **STOP+ASK** (v1.8+)

### Canonical правило v2.0 — заголовок в TITLE-placeholder (Problem #6)

Заголовок слайда у ВСЕХ native-типов (flow/table/kpi/image/chart) вписывается в
**штатный TITLE-placeholder шаблона** — единая позиция/размер **(35, 38) / 963×54
/ 20pt SemiBold CAPS** (графит, на тёмном — белый). Реализовано в общей
`kpi_renderer.set_slide_title()`; `clean_slide_to_blank()` сохраняет
title-placeholder (очищая донорский текст). Если на доноре нет placeholder —
fallback на канонический textbox той же геометрии. Раньше каждый рендерер рисовал
свой header (35,38 vs 35,60; 20 vs 32pt; CAPS/не-CAPS) — теперь единообразно.

### Canonical правило v2.1 — enforce_canonical для CLONE-слайдов

**Важно:** общие презентации из чата собираются преимущественно через
`clone_from_slide` (клон шаблонного донора + текст в слоты), а НЕ через native-
рендереры — поэтому native-фиксы v2.0 до них не доходят. Финальный пост-пасс
`enforce_canonical.py` (вшит в конец `build_v9`, прогон по ВСЕМ слайдам) делает
ТОЛЬКО безопасное и проверенное:
- **цвет** текста: зелёный/белый → `#222222` (белый остаётся только на тёмном
  фоне/плашке) — Problem #2;
- **вес**: bold-флаг → начертание `SB Sans Display Semibold` — Problem #3;
- **размер**: `< 12pt → 12pt`;
- **заголовок** контент-слайда → штатный **TITLE-placeholder** (35,38)/20pt
  SemiBold CAPS, **семантически** (`normalize_header_to_placeholder`): если
  заголовок воткнут текстбоксом — переносим в placeholder; титульные/divider и
  нестандартные (вертикальные/огромные/узкие) заголовки НЕ трогаем.

**Сознательно НЕ делаем** в пост-пассе (ломает чужие макеты вслепую):
bump до 16pt (overflow на код-/плотных боксах) и «поиск заголовка по позиции».
16pt-стандарт и autofit гарантированы только в NATIVE-пути.

### Canonical правило v2.0 — размер шрифта (Problem #5)

**Стандарт контента = 16pt**, комфортный минимум = **12pt**. Меньше 12pt — только
при сильной перегрузке слайда (сначала Overflow Strategy). Дефолты native-
рендереров: flow-блоки и подписи **16pt** (плотные схемы ужимать до 12pt);
**таблицы — дефолт 16pt с авто-уменьшением** (`_autofit_table_font`: подбирает
крупнейший размер, при котором текст влезает в ячейки, вниз до 12pt комфортно /
10pt крайне; форс через `table.font_size`); subtitle **13pt**. `validate_plan`
ужимает overflow не ниже 12pt; `brand_guardian` даёт WARN `size_below_comfortable`
при < 12pt и FAIL `size_too_small` при < 10pt.

### Canonical правило v2.0 — эмфаза через SemiBold-face (Problem #3)

Жирность задаётся **именем начертания** `SB Sans Display Semibold` (встроен в
шаблон → рендерится в PowerPoint везде), а **не bold-флагом**. Все рендереры
(`flow_renderer`, `table_renderer`, `kpi_renderer`, `image_renderer`,
`chart_native_pptx`) переводят `bold=true` в face Semibold + `bold=false`. Bold
(`b=1`) на `SB Sans Display` запрещён (даёт «фейковый» тяжёлый жирный) —
`brand_guardian` выдаёт WARN `bold_flag`.

**Превью (PNG):** LibreOffice не различает SemiBold (схлопывает в Regular),
поэтому `render_slides.py` для превью делает временную подмену
`SB Sans Display Semibold` → `SB Sans Display` + bold-флаг (только во временной
копии для рендера). В итоге в PNG эмфаза видна жирной, а сам `.pptx` остаётся
настоящим SemiBold. На валидацию (`brand_guardian` читает `.pptx`) это не влияет.

### Canonical правило v2.0 — контраст текста (Problem #2)

**Текст ВСЕГДА `#222222`** на белом фоне И на зелёной плашке. Запрещено:
- ❌ зелёный текст на белом (вкл. крупную KPI-цифру 199pt и divider-цифру —
  они тоже `#222222`, user-decision 2026-05-29);
- ❌ белый текст на зелёной плашке (`white-on-green`).

Белый текст допустим **только на тёмном** фоне (графит `#222222` / чёрный).

**Акцент** делается не цветом текста, а **цветным ЭЛЕМЕНТОМ**: зелёная
плашка-подчёркивание (KPI), зелёный divider, footer-плашка, заливка блока
(с графитовым текстом). В `kpi_renderer` главный показатель помечается зелёной
плашкой над цифрой (`_add_accent_bar`); в `flow_renderer` зелёная плашка `green`
теперь несёт графитовый текст. `brand_guardian` выдаёт WARN `colored_text` на
любой зелёный/белый текст вне допустимых случаев.

### Canonical правило v1.8 — table_native (zebra) + anti-distortion

**Регулярные таблицы** (≥3 cols × ≥3 rows, без merged cells) рендерятся как
**native PowerPoint table** (zebra style slide 56). Пользователь редактирует
данные / границы / стили прямо в PowerPoint.

**Стиль zebra (slide 56):**
- Header row: без заливки, bold 12pt #222222, без border
- Body rows: чередуются `#F2F2F2` / белый
- Vertical separators 0.5pt `#C8C8C8` между колонками
- НЕТ горизонтальных границ
- Текст: left + top alignment (нерушимо), SB Sans Display, 11pt body / 12pt header
- Поля ячеек: L/R 12px, T/B 8px
- Первая колонка 1.4× шире (для row labels) — опц. (`first_col_wider: true`)

**Anti-distortion stop+ask (v1.8):** при обнаружении объекта, который может
быть **искажён или утерян** при применении canonical правил — Slide Classifier
ОБЯЗАН остановиться и спросить пользователя. Триггеры: merged cells, RACI,
roadmap, color-coded ячейки, иконки-маркеры, custom annotations, brand-объекты
клиента. Подробности — `agents/02-slide-classifier.md` и memory
`feedback_anti_distortion_safety.md`. Запрет: решать самостоятельно.

### Canonical правило v1.6 — editable schemas (обновлено v1.7)

**Все схемы и блок-диаграммы должны быть редактируемыми (`flow_diagram_native`).**
Скилл рисует блоки, стрелки, пунктирные группы и декор как обычные PowerPoint
shapes. Пользователь в PowerPoint может двигать блоки, менять текст, перекидывать
стрелки — без перевыпуска файла. PNG-вставка схемы как картинки — **запрещена**.

**⭐ v1.7 — Распознавание схем в картинках:** если source `.pptx` содержит
embedded PNG/JPG, который является схемой (блоки + стрелки + подписи) — Slide
Classifier обязан **реконструировать** её как `flow_diagram_native` (не оставлять
картинкой). Эвристика «схема vs. не схема» и алгоритм реконструкции — см.
`agents/02-slide-classifier.md` → раздел «Распознавание схем в картинках».
Если есть сомнение — спросить у пользователя в чате.

**Стиль (canonical v1.7):**
- Блоки: серые `#F2F2F2`, текст `#222222`. Текст внутри: `align=left`,
  `vanchor=top`. Поля 12px со всех сторон + 16px снизу.
- Стрелки: тёмно-серые `#434343`, 1pt, открытая галочка (`type='arrow'`)
  размер 8 (`w='lg', len='med'`). **Только горизонтальные или вертикальные** —
  диагонали запрещены кодом.
- Группировка фаз: пунктирный rect `#888888`, label 10pt SemiBold по центру
- Заголовок: 20pt SemiBold CAPS, top-left (35, 60) — общий стиль content слайдов
- **Safe-area**: блоки должны лежать в `SAFE_TOP=140 .. SAFE_BOTTOM=660`,
  `SAFE_LEFT=35 .. SAFE_RIGHT=1245` — **совпадает с направляющими PowerPoint и
  левым краем заголовка** (заголовок и фреймы по ОДНИМ границам). Константы
  доступны как `flow_renderer.SAFE_*`. **Доводи фреймы до правого края 1245** —
  не оставляй пустого места справа. Панели (`groups`) у границ автоматически
  притягиваются к направляющим (`snap_panel_to_safe`).
- **Дивайдер под заголовком НЕ добавляется по умолчанию** (правило 2026-06-01).
  Заголовок всегда в safe space через TITLE-placeholder. Линию-разделитель
  включать только явно `flow_config["top_separator"]=true` (или
  `table_config`) — для отделения абзацев в специфических типах слайда.

**Когда растягивать схему на всю safe-area** (правило 2026-05-06):
- Растягиваем — если схема **≥2 ряда** И **≥3 колонки**.
- Не растягиваем (естественная высота, центрировать по вертикали) — для линейного
  pipeline в 1 ряд или для простых ≤2-колоночных схем.

**Чеклист после рендера (обязательный):** проверить каждый блок — весь ли текст
помещается? Если нет: расширить frame → уменьшить шрифт (мин 10pt) → пересмотреть
композицию. Подробно — в agents/02-slide-classifier.md.

## Слой валидации (4 уровня)

| # | Уровень | Что ловит | Что НЕ ловит |
|---|---|---|---|
| 1 | `validate_plan.py` | Plan-level: overflow, missing canonical | Реальное содержимое .pptx |
| 2 | `brand_guardian.py` | XML-level: цвета/шрифты/размеры/composition | Визуальные проблемы рендера |
| 3 | `visual_validator.py` | Pixel-level: AR, edges, dominant colors | Семантику и логику |
| 4 | **LLM Visual Verifier** | **Семантика, логика, читаемость** (как человек) | — финальный рубеж |

**КРИТИЧНО**: PASS от 1+2+3 НЕ ГАРАНТИРУЕТ production-quality. Только LLM Verifier даёт окончательный READY.

Все работают через `python-pptx` + `pyyaml` + LibreOffice (`/opt/homebrew/bin/soffice`).

## Принципы работы

- **Не изобретай контент**, которого нет в исходнике (исключение: сценарий 2 "бриф → новая")
- **Используй только layouts из шаблона** (через idx, не через имя)
- **Цвета только из палитры** Cloud.ru (см. brand-rules.md §2)
- **Шрифт только SB Sans** (Display/Text/Interface) или Verdana как fallback
- **Микромодуль 2px**, отступы кратны 10
- **Декомпозиция текста:** не более 7 строк подряд, иначе разбивай на блоки/колонки
- **Запрет ритма:** не использовать одинаковую категорию layouts >2 раз подряд

## Быстрый старт для пользователя

```
> Сверстай презентацию /input/draft.pptx по бренду Cloud.ru
```

Skill активируется по triggers в description. Если нужен явный вызов — `Используй skill cloud-ru-slides для верстки этого pptx`.

## Обновление шаблона (несколько раз в год)

Когда дизайнеры присылают новую версию шаблона — **замены `.pptx` НЕ достаточно**. От шаблона зависят производные артефакты, которые сами не обновляются:

| Артефакт | Как обновляется |
|---|---|
| `brand/template-layouts-dump.json` (координаты placeholder'ов) | **авто** — `sync_template.py --apply` |
| `template/png/*` (эталоны для visual diff) | **авто** — `sync_template.py --apply --render` |
| `brand/template-version.json` (индексы donor-слайдов: blank/kpi/table) | **руками** — сверить, что слайды 30/51/43/44/53/54 всё ещё нужного типа |
| `brand/donor-slot-map.yaml` (карта слотов доноров) | **руками/LLM** — если менялись доноры |
| `brand/template-analysis.md` (семантический каталог) | **руками/LLM** — категории новых layouts |

**Процедура:**
```bash
# 1. положить новый файл по пути из template-version.json → template_path
# 2. посмотреть, что изменилось (без записи):
python3 scripts/sync_template.py
# 3. применить механическую пересборку (dump + бэкап старого):
python3 scripts/sync_template.py --apply --render
# 4. выполнить пункты «ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ» из отчёта
# 5. прогнать 4 baseline-кейса (регресс-тест) перед коммитом
```

Отчёт сам подсветит добавленные/удалённые layouts, «съехавшие» donor-индексы и новые категории на ревью. **Полностью автоматически скилл под новый шаблон НЕ подстраивается** — семантику (какой слайд теперь «чистый белый донор») может подтвердить только человек/LLM.

## Ограничения

- Размер шаблона `.pptx` (29MB) превышает лимит файлов скилла → шаблон НЕ кладётся внутрь. Подтягивается с диска через `scripts/template_path.py` (env `CLOUD_RU_TEMPLATE` / `template-version.json → template_path`); подгрузка в чат — fallback. **Работает только там, где файл достижим** (локальный агент-режим Claude); в облачной песочнице claude.ai диск десктопа недоступен — там шаблон по-прежнему нужно прикладывать.
- Брендбук PDF (27MB) тоже не в скилле — все правила извлечены в `brand/brand-rules.md`
- Для длинных презентаций (>30 слайдов) может потребоваться обработка батчами

## Версия

- v0.1 — 2026-05-01 — smoke-test на `slide_graph_humanity.pptx`
- v0.4 — 2026-05-02 — slot-addressing + style-preserving replace
- v0.5 — 2026-05-02 — donor-slot-map v0.5 + canonical правила
- v0.6 — 2026-05-02 — generation_method (clone/add_layout) per donor
- v0.7 — 2026-05-02 — Brand Guardian (палитра, шрифты, размеры, эмодзи, overflow)
- v0.8 — 2026-05-02 — image handling, тест на Test—part.pptx (12 слайдов)
- **v0.9 — 2026-05-02** — production pipeline:
  - slot-map v0.6: `safe_max_chars` (~70% max) + `canonical_color/size_pt/bold`
  - `validate_plan.py`: auto-add canonical overrides, auto-apply STRATEGY 3 при overflow
  - Brand Guardian + composition checks (title pressed, subtitle green, KPI consistency)
  - Layout Designer v0.9 algorithm (selection by safe_max_chars, fallback strategies)
  - SKILL.md: real LLM pipeline с 10 шагами

См. `LEARNINGS.md` для накопленного опыта итераций и известных ограничений.
