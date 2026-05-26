# LEARNINGS — журнал ошибок и фиксов

> Каждая итерация добавляет сюда: что не работало → что сделали → ссылка на коммит/файл.

## v0.1 — 2026-05-01 — Первая сборка скилла

### ✅ Что заработало
- Парсинг `Cloud.ru_Template_2026.pptx` через python-pptx — извлечены все 102 layouts с координатами placeholder'ов в EMU
- Парсинг сырого `slide_graph_humanity.pptx` — корректно извлёк title + body + image
- `build_pptx.py` собирает .pptx со слайдами на нужных layouts шаблона Cloud.ru
- `kill_widows.py` корректно ставит nbsp после коротких слов и между числом+единицей, тире, кавычки
- Файл результата открывается через `Presentation()` без ошибок, layouts применены, текст на месте

### ⚠️ Известные проблемы

#### LRN-001: SlideLayout не hashable
**Симптом:** `TypeError: cannot use 'pptx.slide.SlideLayout' as a dict key`
**Фикс:** использовать `id(lay)` вместо `lay` в lookup-словарях
**Файл:** `scripts/parse_pptx.py:45,57`

#### LRN-002: Duplicate slide files в результирующем .pptx
**Симптом:** при сохранении — `UserWarning: Duplicate name: 'ppt/slides/slide1.xml'`
**Причина:** `build_pptx.py` удаляет элементы из `sldIdLst` (XML-manifest), но физические файлы slide*.xml внутри ZIP остаются. Когда python-pptx добавляет новые слайды, они получают имена slide1, slide2, slide3 — конфликтуют с физическими остатками.
**Воздействие в v0.1:** файл открывается и работает, но содержит мусорные XML внутри (≈25MB вместо ~3MB)
**Фикс v0.2 (РЕАЛИЗОВАН в `build_v2_demo.py:remove_existing_slides`):**
```python
def remove_existing_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)
```
**Результат:** v2_demo_aitest2.pptx — **9.6 MB** (было 25MB в v0.1), нет warning'ов.

#### LRN-003: Тестовая презентация имеет свой master
**Наблюдение:** `slide_graph_humanity.pptx` использует `layout='Blank'` из своего master, не из Cloud.ru. Это норма для драфтов — мы не должны полагаться на исходный layout, только на текстовое содержимое.
**Что делать:** Brief Reader работает только с текстом, не с layout исходника. Slide Classifier выбирает категорию по `intent` (содержимому), Layout Designer выбирает idx из 102 Cloud.ru layouts.

### 📋 Чек-лист smoke-теста (v0.1)

- [x] `parse_pptx.py` извлекает структуру сырого .pptx
- [x] `build_pptx.py` создаёт .pptx с layouts из шаблона Cloud.ru
- [x] `kill_widows.py` применяет русскую типографику
- [x] Результат открывается, текст в нужных placeholders
- [ ] Полный pipeline через все 9 агентов (TODO v0.2)
- [ ] Brand Guardian валидирует цвета/шрифты результата (TODO v0.2)
- [ ] Удаление мусорных XML после добавления слайдов (LRN-002)
- [ ] Поддержка инфографики (Infographic Maker → shapes в PPT)

### 🎯 Что нужно для v0.2

1. ~~**Фикс duplicate slides** (LRN-002)~~ ✅ ИСПРАВЛЕНО в `build_v2_demo.py`
2. ~~**Реальная вёрстка** (не просто placeholder fill)~~ ✅ ДЕМО В `build_v2_demo.py` (8 слайдов из 5, явные стили SB Sans Display/Text, цвета палитры, KPI shapes 80pt Green)
3. **Реальный pipeline через LLM-агентов** — сейчас build_v2_demo использует ручной план под конкретный файл; нужно сделать обобщённый skill prompt-based pipeline
4. **Поддержка изображений** — копировать images из draft в новые placeholder'ы
5. **Поддержка таблиц** — пока пропускаем
6. **Тест на больших презентациях** — `Test.pptx` (22MB) и `Marketing_cloude.pptx` (51MB)

## v0.2 demo — 2026-05-01 — Реальная вёрстка

### Кейс
`Презентации на тест/ai test 2.pptx` — 5 сырых слайдов (все на одном layout, эмодзи в body, длинные заголовки).

### Стратегия 5 → 8 слайдов
| # | Layout idx | Layout name | Что |
|---|---|---|---|
| 1 | 52 | Титул / три строки_2 | Заголовок 40pt + спикер |
| 2 | 8 | Раздел / Зелёный 1 | "01 / КОНТЕКСТ" 100pt |
| 3 | 31 | 3 колонки | 3 факта с green-заголовками + body 14pt |
| 4 | 23 | Важная информация | Цитата 60pt |
| 5 | 54 | Раздел / тёмный | "02 / РЕЗУЛЬТАТЫ" 100pt Green |
| 6 | 10 | Контент / Белый 1 | 3 KPI цифры **через shapes** 80pt Green + подписи |
| 7 | 28 | 4 текстовых блока | Механизмы (без эмодзи) |
| 8 | 95 | Зелёный / Логотип | Финал |

### Ключевые техники в build_v2_demo.py
- `style_run()` — explicit `font.name / size / bold / color.rgb` на каждом run
- `add_text_box()` — создание text-box с явными стилями (для KPI)
- `add_filled_rect()` — прямоугольники без скруглений и теней (брендбук)
- `remove_existing_slides()` — корректное удаление через `drop_rel`

### Что НЕ работает в v0.2 demo
- Hardcoded под конкретный файл — план не генерируется агентами автоматически
- Не копируются images из исходника (только текст)
- Эмодзи дропнуты вручную, не автоматически (в Copy Editor)

### 🧪 Test artifacts

- `output/parsed_slide_graph.json` — результат parse_pptx на тестовом draft
- `output/smoke_plan.json` — ручной план для смоук-теста (имитация выхода агентов)
- `output/smoke_result.pptx` — собранный 3-слайдовый результат на Cloud.ru шаблоне

---

## v0.19 — 2026-05-04 — Canonical rules из шаблона

**Что:** извлёк правила из guide-слайдов самого `Cloud.ru_Template_2026.pptx` (slide 5, 21-35, 43, 62-66, 81-83, 87-88) — 14 секций в `brand/template-canonical-rules.md`.

**Главное:** template содержит **инструкции от дизайнеров Cloud.ru**, я раньше использовал эти слайды просто как доноры, игнорируя их инструктивный текст.

**Уроки в memory:**
- `feedback_canonical_priority_over_intuition.md` — Brand Guardian уже знает правила, не перетолковывать
- `feedback_kpi_digit_limit.md` — donor 43/44 199pt frame = max 2 цифры
- `feedback_table_donor_needs_data.md` — donor 53/54 имеют PNG-заглушку

## v0.20 — 2026-05-04 — Native renderers (KPI/image/chart)

**Проблема:** donors 43/44 (KPI) и 53/54 (table) — это **guide-слайды** с фиксированным placeholder content (не доноры для произвольного content). Использование как clone source ломает композицию.

**Решение — Native rendering:**
- `scripts/kpi_renderer.py` — рисует KPI shapes на blank canvas (3 цифры равного размера, accent green)
- `scripts/image_renderer.py` — auto-fit image на canvas с aspect preserve (mode: fit/fill)
- `scripts/chart_engine.py` — matplotlib redraw в canonical pastel palette (изучено по эталонам slides 45/46/47/50)
- `scripts/visual_validator_v2.py` — PIL pixel-level: unfilled placeholders, off-palette > 15%
- `scripts/build_v9.py` — поддерживает `slide_type: "kpi_native" | "image_native" | "chart_native"`

**Уроки в memory:**
- `feedback_no_false_positive_pass.md` — НИКОГДА не объявлять PASS без визуальной проверки

## v1.3 — 2026-05-04 — Polish + e2e workflow

**Что сделано:**
- Slide 3 chart: упрощён до 3 серий (canonical эталон slide 50) + chart_engine canonical pastel palette
- Slide 6 finale: убран unused picture-shape из donor 25 (`remove_shapes: [2]`)
- Caption под chart в серой плашке (как эталон slide 49)
- Image_renderer: fit/fill modes
- Slide Classifier (агент 02): auto-detect slide_type через decision tree
- Validate_plan: bug fix для native types (был FAIL без clone_from_slide)
- Cleanup output: 466MB → 48MB (158 → 14 файлов)
- E2E emulated workflow: 12 шагов (G1+01+02+04+03+07+validate+build+render+brand+visual+LLM verdict) на slide_graph

**Финальные тесты (все после rebuilds + честных glазами verdicts):**
- slide_graph: 6 sl, chart_native canonical 3-color, **5-dim 4.8/5**
- ai_test2: 8 sl, kpi_native, **5-dim 4.8/5**
- testpart: 13 sl, kpi_native white+dark + image_native, голубой placeholder убран, **5-dim 4.7/5**
- Marketing: 8 sl, kpi_native, **5-dim 4.8/5**
- e2e: 6 sl, **LLM Verifier verdict 4.8/5 READY**

**4 уровня валидации:** validate_plan (WARN OK) → brand_guardian (PASS 100/100) → visual_validator_v2 (PASS 0 issues) → **LLM Visual Verifier (per-slide 5-dim)**.

## v1.3.1 — 2026-05-05 — MVP up на новом окружении (mdmolotkova)

**Контекст:** новый пользователь, новое окружение `/Users/mdmolotkova/Desktop/Презентации в ИИ 2/`. v1.3 baseline воспроизведён.

### Окружение установлено
- Python 3.9.6 + `python-pptx 1.0.2`, `Pillow 11.3.0` (уже было).
- Доустановлены через `pip3 install --user`: `pyyaml 6.0.3`, `matplotlib 3.9.4`, `numpy 2.0.2`, `lxml 6.0.2`.
- LibreOffice 26.2.3 — установлен через DMG в `~/Applications/LibreOffice.app` (без brew/sudo, /Applications защищена SIP-permission).
- Поскольку `pdftoppm` (poppler) отсутствует, доустановлен `pymupdf 1.26.5` как замена.

### Изменения в коде
- `scripts/render_slides.py`:
  - `find_soffice()` — добавлен путь `~/Applications/LibreOffice.app/Contents/MacOS/soffice`.
  - Добавлена функция `pdf_to_pngs()`: использует `pdftoppm` если есть, иначе fallback на `PyMuPDF` (fitz).
  - Это снимает жёсткую зависимость от системного poppler.

### Smoke-test pipeline на slide_graph_humanity.pptx
- `parse_pptx.py` → 1 input slide (исходник перегружен — title + 4 пункта body + цитата).
- Plan вручную (имитация LLM agents): donor 21 (content_text_white) + donor 42 (callout_white) — 2 слайда на выходе.
- `build_v9.py` → `output/result_slidegraph_smoke.pptx` (2 слайда, 0 pictures).
- `render_slides.py` (через LibreOffice + PyMuPDF) → 2 PNG, 96dpi.
- `brand_guardian.py` → **PASS 100/100**.
- `visual_validator_v2.py` → **PASS 0 issues**.
- `validate_plan.py` → **FAIL slide[1].body**: 371 chars > 200 max (overflow 1.85x). Корректная сработка правил — нужен split или другой donor с большим slot. Не баг, а ожидаемая дисциплина STRATEGY 2.
- LLM Visual Verifier (мои глаза): оба слайда брендово (cloud.ru logo, типографика SB Sans, decorative shapes на цитате), читаемо, без артефактов. Score ~4.5-4.7/5.

### Итог MVP
Pipeline полностью функционален end-to-end: parse → plan → validate → build → render → 3-level auto-validation → LLM verdict. Все 4 валидатора возвращают понятные verdict-ы.

### Backlog known limitations (унаследовано)
- Длинные презентации (30+ слайдов) — не тестированы (`Test.pptx` 21 sl).
- Donor 86 теряет subtitle (1 slot вместо 2).
- Image enrichment не отличает chart от photo.
- Table flow только для donor 53/54 (PNG-заглушка нужна).
- Smart fallback для overflow ≥30% — нужен split/другой donor; auto-split в Slide Classifier ещё не интегрирован в скриптовый pipeline (только в LLM-промпте агента 02).

## v1.4 — 2026-05-05 — Editable native PowerPoint chart

**Проблема:** существующий `chart_native` (через `chart_engine.py`) рендерит matplotlib → PNG → вставляет картинкой. Пользователь не может править данные в PowerPoint.

**Решение:** новый `slide_type: chart_pptx_native` через `pptx.chart.add_chart()` — настоящий PowerPoint chart с привязанным датасетом. Пользователь жмёт ПКМ → «Изменить данные» → открывается Excel-окно.

**Реализация:**
- Новый модуль [scripts/chart_native_pptx.py](pptx-skill/scripts/chart_native_pptx.py):
  - `add_chart_to_slide(slide, cfg, l, t, w, h, dark)` — primitives.
  - `render_chart_pptx_slide(slide, cfg, dark)` — полная сборка (header 32pt + chart + caption в серой плашке).
  - Поддержка: `area_stacked`, `area_100`, `bar`, `bar_stacked`, `line`, `pie`.
  - Canonical palette: GREEN зарезервирован под `accent_idx`, остальные серии распределяются по `NON_ACCENT_COLORS` (graphite, pastel blue, pastel yellow, pastel mint, gray) — обеспечивает различимость.
  - Для `area_100` — ось Y в процентах через `value_axis.tick_labels.number_format = "0%"`.
- В [build_v9.py](pptx-skill/scripts/build_v9.py):
  - Добавлен импорт `render_chart_pptx_slide`.
  - В списке native types добавлен `chart_pptx_native` (использует blank donor, как kpi/image).
  - Соответствующий бранч в pipeline между `image_native` и `chart_native`.

**Тест на humanity/slide_graph_humanity.pptx:**
- 5 серий (Сельское хоз → Реляционный сектор), 1850–2050, accent_idx=3.
- Вердикты: brand_guardian PASS 100/100, visual_validator_v2 WARN (off-palette 31.4% — pastel yellow/mint не в первом ряду; canonical §2 разрешает их для charts, но валидатор не различает контекст — known limitation).
- LLM verdict: 4.7/5 — chart редактируемый, читаемый, accent виден к 2050, серии различимы.

**Уроки в memory:**
- `feedback_chart_palette_pastel_allowed.md` — visual_validator_v2 WARN на pastel в charts — false positive по pixel-level правилам, не блокер.
- `feedback_chart_title_template_style.md` — header слайда с native chart должен совпадать стилем с content slides шаблона: 20pt SemiBold CAPS, top-left (35, 60), не 32pt не-CAPS.

## v1.5 — 2026-05-05 — Audit follow-up: brand validators + palette infra

Контекст: пользователь провёл сторонний аудит → 5 файлов в `~/Desktop/audit/`. Реализовали выборочно P0+P1 (P0-2 orchestrator пропущен — pipeline через чат), P1-8 (EN localization) пропущен — кейсы только русские.

### Этап 1 — критика
- **1.1** `validate_plan.py:41` — `chart_pptx_native` добавлен в список native types. Раньше план с editable chart проходил по случайности (silent-pass, build_v9 потом падал).
- **1.2** `visual_validator_v2.py` — добавлены `PALETTE_CHART_EXTENSION` (pastel-цвета для chart-серий) и параметр `--plan`. Когда передан plan.json — для slide_type ∈ {chart_pptx_native, chart_native} используется расширенная палитра, false-positive WARN на пастельных диаграммах больше не выдаётся. Backward-compat: без `--plan` старое поведение.
- **1.3** Counter «слишком много зелёного» на divider-слайдах — fact-checked: проблемы нет, `MAX_GREEN_AREA_PCT=30` объявлена, но нигде не используется. Мёртвая константа. Аудит ошибся.
- **1.4** `e2e_pipeline.sh` — переписан под актуальные скрипты (`build_v9`, `visual_validator_v2 --plan`) и новую модель «pipeline через чат». LLM-этапы делаются в чате, bash-скрипт прогоняет детерминированную часть от plan.json до отчётов одной командой. `bash scripts/e2e_pipeline.sh <plan.json> <out_dir> [template]`.

### Этап 2 — закрыть дыры в проверке бренда
- **2.1** `check_arrow_colors` в `brand_guardian.py` — детектит connector-линии и AUTO_SHAPE с `ARROW` в типе. Если line/fill близок к фирменному зелёному (tolerance 30 RGB) — **violation** `green_arrow`. Canonical §3 (slide 66): стрелки серые, не зелёные.
- **2.2** `_count_significant_digits` + warn в `kpi_renderer.py` — для hero-цифры 199pt считаем значащие цифры (без `, . - ~ K M x`). 3+ → stderr WARN с подсказкой использовать сокращение или сменить donor 47.
- **2.3** Счётчик тёмных слайдов ≤40% — пропущен по решению пользователя.
- **2.4** Два правила в `brand_guardian.py`:
  - **A** `extended_without_green` — slide-level WARN: на слайде есть extended-цвета (Magenta/Carrot/Coral/etc.) и НЕТ фирменного зелёного.
  - **B** `colored_text` — text-run-level WARN. Whitelist для текста: graphite (#222222 и оттенки чёрного) + grays. На тёмных слайдах дополнительно разрешены белый и зелёный (#26D07C и оттенки). Любой другой цвет → WARN. Tolerance ±20 RGB. `_is_dark_slide` детектит по layout name (содержит "тёмн/dark/graphite") или background luminance < 128.
- **2.5** `find_soffice` в `render_slides.py` — добавлены Windows-пути (`C:\Program Files\LibreOffice\program\soffice.exe` и x86), Linux (`/usr/bin/soffice`), `soffice.exe` через PATH.
- **2.6** `brand/palette.json` — единый источник истины для палитры. 7 секций: base / base_alts / extended / chart_extension / text_neutral / text_dark_extra / green_for_arrows / unfilled_flags. Подключено в `brand_guardian.py`, `visual_validator_v2.py`, `chart_native_pptx.py` через helper `_load_palette()`. Hardcoded fallback оставлен на случай отсутствия файла.
- **2.7** `brand/template-version.json` — маппинг slide-индексов: blank_donors (white=30, dark=22), guide_donors. `kpi_renderer.py:_load_template_version()` грузит JSON, константы `BLANK_DONOR_WHITE/DARK` подтягиваются динамически. При обновлении шаблона править только JSON.
- **2.8** `tests/regression.py` + `tests/baselines/humanity/` (plan.json + expected.json). Прогон: запустить e2e_pipeline в temp-папке, сверить brand_guardian (verdict + score_avg + violations) и visual_validator_v2 (verdict + total_issues) с expected. `python3 tests/regression.py [case_name]`. Baseline humanity сейчас фиксирует: 4 слайда / brand PASS 100/100 / visual PASS 0 issues.

### Уроки в memory
- `feedback_audit_p0_p1_done.md` — какие пункты аудита внедрены, какие отложены.
- Обновлён `project_environment_status.md` — pyyaml/matplotlib/pymupdf уже установлены.

## v1.6 — 2026-05-26 — Flow diagrams / схемы (native, editable)

**Проблема:** до v1.6 скилл умел только charts (`chart_pptx_native`), KPI
(`kpi_native`) и images (`image_native`). Схемы / блок-диаграммы / pipeline
рисовать был не способен — приходилось вставлять PNG, что неоптимально:
не редактируется в PowerPoint, не консистентно с brand-стилем.

**Контекст:** в соседней сессии (заблокирована из-за `Image is too large >2000px`
после рендера в PNG) пользователь экспериментально вручную написал
`output/flow_diagrams/build_flow.py` — 492-строчный скрипт, который рисует
2 эталонные схемы (SGR + Pipeline аналитики) через примитивы python-pptx:
серые блоки, чёрные стрелки с маленьким наконечником, пунктирные группы
Phase 1/2, зелёные L-уголки декор. Результаты `sgr_schema.png` и
`analytics_pipeline.png` сохранены в `references/flow_diagrams/`.

**Решение:** новый `slide_type: "flow_diagram_native"` через модуль
`scripts/flow_renderer.py`. Примитивы (`add_block`, `add_arrow`,
`add_dashed_rect`, `add_label`, `add_header`, `add_top_separator`,
`add_decor_diagonals`) вынесены из эксперимента в переиспользуемые
функции + высокоуровневая `render_flow_diagram_slide(slide, config, dark)`.

**Реализация:**
- [scripts/flow_renderer.py](pptx-skill/scripts/flow_renderer.py) — новый модуль:
  - Низкоуровневые примитивы: `add_block(fill: gray|white|green|dark)`,
    `add_label`, `add_arrow(with_head, dashed)`, `add_dashed_rect`,
    `add_header(dark)`, `add_top_separator`, `add_decor_diagonals`.
  - Высокоуровневая `render_flow_diagram_slide(slide, flow_config, dark)`:
    собирает header → blocks (с id) → groups (dashed rects + labels) →
    arrows (by ref `from`/`to` либо по координатам) → labels → decor.
  - Палитра грузится из `brand/palette.json` (единый источник истины).
  - Standalone CLI: `python3 flow_renderer.py <flow_config.json> <template.pptx> <out.pptx>`.
- [scripts/build_v9.py](pptx-skill/scripts/build_v9.py):
  - Импорт `render_flow_diagram_slide` (try/except → `FLOW_RENDERER_AVAILABLE`).
  - `flow_diagram_native` добавлен в список native types (использует blank donor).
  - Новая ветка в pipeline между `chart_pptx_native` и `chart_native`.
- [scripts/validate_plan.py](pptx-skill/scripts/validate_plan.py):
  - `flow_diagram_native` принят как native type (требует поля `flow`).
  - Дополнительные проверки: блоки в canvas 1280×720, минимальный размер
    блока (w≥60, h≥24), arrow-refs `from`/`to` соответствуют block.id.
- [agents/02-slide-classifier.md](pptx-skill/agents/02-slide-classifier.md):
  - Триггеры для авто-детекции: intent ∈ {schema, flow, pipeline, process,
    architecture}, key_phrase содержит «схема/архитектура/pipeline/процесс»,
    источник содержит SmartArt или прямоугольники + connectors.
  - Routing decision tree обновлён (приоритет ниже chart, выше KPI).
  - Композиционные подсказки (≤ 8 блоков → иначе split, gap 22px, etc.).
- [SKILL.md](pptx-skill/SKILL.md):
  - Таблица скриптов: новый `flow_renderer.py` ⭐.
  - Раздел native slide types: JSON-схема `flow_diagram_native`.
  - Canonical правило v1.6 — editable schemas, PNG-вставка запрещена.
- [references/flow_diagrams/](pptx-skill/references/flow_diagrams/) — эталонные
  PNG (sgr_schema, analytics_pipeline) + README с canonical критериями PASS.
- [tests/baselines/flow_diagram/](pptx-skill/tests/baselines/flow_diagram/) —
  регрессионный тест.

**Что НЕ перенесено в репо:**
- Сам `flow_diagrams.pptx` (10MB) и `build_flow.py` — остались в `output/`
  (gitignored). `build_flow.py` — это эксперимент с захардкоженными
  координатами под 2 конкретные схемы; правильный путь — JSON-конфиг через
  новый renderer. Эталонные PNG сохранены в `references/`.

**Стиль (canonical):** идентичен build_flow.py, который пользователь утвердил
визуально. Палитра подтянута из `brand/palette.json`, не дублирована.

**Уроки в memory:**
- Эксперименты с native rendering из output-папки нужно вытаскивать в
  `scripts/<feature>_renderer.py` через паттерн «примитивы + высокоуровневая
  функция, принимающая JSON-config». Такой паттерн уже работает для
  KPI/chart/image — flow стал четвёртым.
- Координаты блоков в эксперименте захардкожены под конкретные слайды;
  в продакшене ответственность за layout-расчёт лежит на Layout Designer
  (агент 04) или на самом Slide Classifier при детекции flow.

