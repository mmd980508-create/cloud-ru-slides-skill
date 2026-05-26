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
- **Обязательно (для финальной вёрстки):** файл `Cloud.ru_Template_2026.pptx` (29MB, не входит в скилл — подгружается в чат)
- **Опционально:** референсные слайды (PNG), брендбук (если нужны нестандартные правила)

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

[7] python3 scripts/build_v6.py output/plan_validated.json template/Cloud.ru_Template_2026.pptx output/result.pptx

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
| **`kpi_renderer.py`** | **Рисует KPI shapes на blank canvas (3 числа равного размера, 1 accent green)** |
| **`image_renderer.py`** | **Image-as-content auto-fit (mode: fit/fill, опц. caption в серой плашке)** |
| **`chart_native_pptx.py`** ⭐ | **DEFAULT для charts (v1.4+). Редактируемая PowerPoint chart через `pptx.chart.add_chart()` — пользователь меняет данные через Edit Data → Excel. Поддержка: area_stacked/area_100/bar/bar_stacked/line/pie** |
| `chart_engine.py` | Legacy: Matplotlib chart redraw в PNG (canonical pastel palette). Использовать ТОЛЬКО когда нужны custom annotations или прозрачные overlapping areas, которые native chart не умеет |
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
```

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

## Ограничения

- Размер шаблона `.pptx` (29MB) превышает лимит файлов скилла → шаблон НЕ кладётся внутрь, пользователь подгружает его при работе
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
