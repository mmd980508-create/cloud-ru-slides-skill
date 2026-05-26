# HANDOFF — Cloud.ru Slides Skill

> **Для нового агента/разработчика.** Это входное окно для продолжения разработки скилла Cloud.ru Slides Skill после v1.7.
>
> Прочитать перед первым коммитом. После прочтения — переходи к [Что делать первым](#что-делать-первым).

---

## TL;DR за 60 секунд

- **Что:** Claude.app skill для авто-вёрстки `.pptx` по бренду **Cloud.ru 2.0**
- **Версия:** v1.7 (2026-05-26) — `cloud-ru-slides-skill-v1.7.zip`
- **Состояние:** **Working baseline.** 4 кейса прошли LLM-визуальную проверку 4.7-4.8/5
- **Архитектура:** Per-slide LLM design loop + native renderers (KPI/image/chart) + 4 валидатора (validate_plan → brand_guardian → visual_validator_v2 → LLM Visual Verifier)
- **Главный файл:** `SKILL.md` (frontmatter + workflow)
- **Шаблон Cloud.ru:** `../template/Cloud.ru_Template_2026.pptx` (29MB, **НЕ внутри skill** — лимиты Claude.app)
- **Memory:** 22 файла в `~/.claude/projects/-Users-gmmelnikov-Desktop-----------------/memory/` — обязательно прочитать!

---

## 1. Контекст проекта

### Кто пользователь
Глеб Мельников — моушен-дизайнер Cloud.ru. Не разработчик, но не боится терминала. Работает на macOS, использует Max-подписку. Главное: **визуальное качество финального .pptx > соответствие коду**.

### Что хочет
Автоматическая вёрстка `.pptx` по бренду Cloud.ru. Три сценария:
1. **Драфт → вёрстка** (главный): сырой `.pptx` от коллег → красивая презентация
2. **Бриф → презентация** (markdown/docx → .pptx с нуля)
3. **Аудит** (готовая → отчёт о соответствии бренду)

### Главные правила пользователя (из memory!)
- **Template первичен, брендбук вторичен** — `Cloud.ru_Template_2026.pptx` = главный ориентир
- **SB Sans Display только** (Regular + SemiBold)
- **«PERFECT!» без визуальной проверки = бан** — НИКОГДА не объявлять PASS без LLM-визуального ревью
- **Brand Guardian PASS ≠ хороший дизайн** — финал решает LLM-визуальная критика

---

## 2. Архитектура за 3 минуты

### Pipeline v0.17 — Per-slide LLM Design Loop

```
PHASE 1 — GLOBAL ANALYSIS (один раз):
  G1. parse_pptx / parse_md / parse_docx + extract_images
  01. Brief Reader     — общий бриф презентации
  02. Slide Classifier — категория + slide_type для каждого слайда (с auto-detect)

PHASE 2 — PER-SLIDE LOOP (для каждого слайда):
  04. Layout Designer  — выбирает donor (или native render) с design thinking
  03. Content Distrib  — распределяет text по slots
  07. Copy Editor      — kill_widows + анти-эмодзи
  S4. validate_plan    — gate: canonical sizes/colors auto-add
  S5. build_v9         — собирает один слайд (clone donor / kpi_native / image_native / chart_native)
  S6. render_slides    — PNG через LibreOffice
  S7. 10. LLM Visual Verifier — LLM смотрит PNG глазами (8 designer criteria)
  S8. If verdict ≠ READY → fix one issue → goto S5 (max 3 iterations)

PHASE 3 — GLOBAL VERIFICATION:
  V1. brand_guardian       — палитра/шрифты/композиция XML-level
  V2. visual_validator_v2  — PIL pixel-level (unfilled, off-palette)
  V3. 09. Verifier         — объединение verdict-ов → READY/NEEDS_REWORK
```

### Ключевая инновация v1.3 — Native renderers
Donors 43/44 (KPI) и 53/54 (table) — **guide-слайды** в шаблоне (показывают примеры размеров), не доноры. Native rendering = рисуем shapes на чистом canvas:

| Native type | Скрипт | Как использовать |
|---|---|---|
| `kpi_native` | `scripts/kpi_renderer.py` | 1-3 цифры равного размера, 1 accent green |
| `image_native` | `scripts/image_renderer.py` | Auto-fit (mode: fit/fill) с aspect preserve |
| **`chart_pptx_native`** ⭐ (v1.4 default) | `scripts/chart_native_pptx.py` | **Редактируемая PowerPoint chart через `add_chart()` — пользователь меняет данные через Edit Data → Excel** |
| `chart_native` (legacy) | `scripts/chart_engine.py` | Matplotlib redraw в canonical pastel palette (PNG, не редактируется). Только для custom annotations |

Plan слайда:
```json
{"slide_type": "kpi_native", "kpi": {
  "title": "...",
  "numbers": [{"value": "12", "desc": "...", "accent": false}, ...]
}}

// v1.4 — DEFAULT для всех графиков и диаграмм. РЕДАКТИРУЕМАЯ chart.
{"slide_type": "chart_pptx_native", "chart": {
  "type": "area_100|area_stacked|bar|bar_stacked|line|pie",
  "title": "...", "caption": "...",
  "x": [...], "series": [{"name": "...", "data": [...]}, ...],
  "accent_idx": N
}}
```

### Canonical правило v1.4 — editable charts
**Все графики и диаграммы — редактируемые** (`chart_pptx_native`). Не PNG.
Пользователь должен иметь возможность открыть «Изменить данные» в PowerPoint
и править цифры через встроенный Excel. См. `brand/template-canonical-rules.md` §3
(подраздел «Редактируемые диаграммы»).

---

## 3. Карта файлов

### `pptx-skill/` (то что попадает в Claude.app skill zip)

```
SKILL.md                  ← главный файл (frontmatter + workflow). Читать первым
HANDOFF.md                ← этот файл
README.md                 ← user-facing документация
LEARNINGS.md              ← журнал ошибок и фиксов (v0.1 → v1.3, 13 учений)

agents/                   (11 markdown инструкций для LLM-агентов)
  01-brief-reader.md
  02-slide-classifier.md       ← auto-detect slide_type + split overcrowded (v1.3)
  03-content-distributor.md
  04-layout-designer.md
  05-icon-picker.md
  06-infographic-maker.md
  07-copy-editor.md
  08-brand-guardian.md
  09-verifier.md
  10-llm-visual-verifier.md    ← 8 designer criteria + scale 1-5 (v1.3 final)
  auto-fix-protocol.md         ← 6-step chain для overflow

brand/                    (правила бренда)
  brand-rules.md                  ← общие правила (палитра/шрифты)
  template-canonical-rules.md     ← ⭐ КРИТИЧНО: правила извлечённые из guide-слайдов шаблона (slide 5/21-35/43/62-66/87-88)
  DESIGN.md                       ← философия дизайна
  donor-slot-map.yaml             ← каталог 36+ donors с slot mapping
  donor-profiles.json             ← profiling каждого donor (PNG-stripping и т.д.)
  donor-frames-dump.json          ← machine dump координат
  template-analysis.md            ← семантический каталог 102 layouts
  template-layouts-dump.json      ← XML дамп всех layouts
  template-slides-catalog.json    ← каталог 88 эталонных слайдов

scripts/                  (Python utilities)
  parse_pptx.py / parse_md.py / parse_docx.py
  extract_images.py
  build_v9.py                  ← главный builder (clone + native types)
  build_v8.py / v7 / v6 / v5   ← legacy (для совместимости)
  kpi_renderer.py              ← native KPI
  image_renderer.py            ← native image with auto-fit
  chart_engine.py              ← native chart (matplotlib)
  validate_plan.py             ← gate перед build
  brand_guardian.py            ← XML-level валидация
  visual_validator_v2.py       ← PIL pixel-level валидация
  render_slides.py             ← .pptx → PNG (LibreOffice)
  kill_widows.py               ← русская типографика

dictionaries/             (для kill_widows)
  short-words-ru.txt
  whitelist.txt

input/                    ← сюда пользователь кладёт draft.pptx
output/                   ← результаты build (cleaned после v1.3)
```

### Артефакты вне skill (не в zip)

```
../template/Cloud.ru_Template_2026.pptx   ← 29MB — пользователь подгружает в чат вместе со скиллом
../template/png/                           ← 88 эталонных PNG (для LLM визуального сравнения)
../template/рефы/                          ← дополнительные рефы
../брендбук/Брендбук Cloud.ru 2.0_cond.pdf ← 27MB PDF брендбука
../Презентации на тест/                    ← 5 тестовых .pptx
../Плохие генеративные презентации/        ← negative examples
```

---

## 4. Memory — обязательно перенести

**Путь:** `~/.claude/projects/-Users-gmmelnikov-Desktop-----------------/memory/`

22 файла. Они **критичны** для нового агента — содержат накопленные user feedback и lessons. Без них новый агент повторит ошибки.

### Самые важные (читать в первую очередь)
| Файл | Зачем |
|---|---|
| `MEMORY.md` | Index всех memory entries |
| `feedback_no_false_positive_pass.md` | НИКОГДА не объявлять PASS без визуального ревью |
| `feedback_designer_style_validation.md` | Финал = LLM designer critique (8 criteria), НЕ code |
| `feedback_template_priority.md` | Template первичен, брендбук вторичен |
| `feedback_canonical_priority_over_intuition.md` | Brand Guardian уже знает правила |
| `feedback_overflow_strategy.md` | 4 стратегии при overflow (donor → split → -20-30% → copy edit) |
| `feedback_split_overcrowded_slides.md` | Когда разбивать на 2 слайда |
| `feedback_kpi_digit_limit.md` | KPI 199pt = max 2 цифры |
| `feedback_table_donor_needs_data.md` | Donor 53/54 имеет PNG-заглушку |
| `feedback_visual_loop_required.md` | render → diff с эталоном — обязательно |
| `feedback_clone_donor_strategy.md` | Clone donor + replace text > add_slide(layout) |

### Контекст (откуда мы пришли)
- `user_profile.md` — кто Глеб, что важно
- `project_cloud_slides_skill.md` — общий контекст проекта
- `brand_assets_summary.md` — что есть в репо
- `reference_vault_project.md` — vault folder note path

---

## 5. Что делать первым (новый агент)

### Шаг 1 — прочитать в этом порядке
1. **Этот HANDOFF.md** (целиком)
2. **`memory/MEMORY.md`** + 11 ключевых feedback (см. таблицу выше) — 30 мин
3. **`SKILL.md`** — главный workflow
4. **`brand/template-canonical-rules.md`** — правила Cloud.ru
5. **`LEARNINGS.md` tail** (последние 100 строк, v1.3 секция) — что уже сделано

### Шаг 2 — воспроизвести baseline
```bash
cd "/Users/gmmelnikov/Desktop/Презентации в ИИ/pptx-skill"

# Парсинг тестового draft
python3 scripts/parse_pptx.py "../Презентации на тест/ai test 2.pptx" output/parsed.json

# Build (нужен ручной plan — пример: output_archive есть в backup)
# python3 scripts/build_v9.py output/plan.json ../template/Cloud.ru_Template_2026.pptx output/result.pptx

# Render для визуальной проверки
python3 scripts/render_slides.py output/result.pptx output/render/

# Валидация
python3 scripts/brand_guardian.py output/result.pptx output/brand_report.json
python3 scripts/visual_validator_v2.py output/render/ output/visual_report.json

# ⚠️ КРИТИЧНО: после всех PASS — LLM ОБЯЗАТЕЛЬНО смотрит каждый PNG глазами
# (см. agents/10-llm-visual-verifier.md)
```

### Шаг 3 — известные проблемы для починки

| Severity | Issue | Где |
|---|---|---|
| 🔴 high | Длинные презентации (30+ slides) не тестированы | `Test.pptx` 21 sl, 37 pictures |
| 🟡 med | Donor 86 теряет subtitle (1 slot вместо 2) | `donor-slot-map.yaml` |
| 🟡 med | Image enrichment не отличает chart от photo | `layout_designer.py` |
| 🟡 med | Table flow только для donor 53/54 | `build_v9.py` |
| 🟢 low | Donor 78 требует ровно 3 unique descs | `donor-slot-map.yaml` |
| 🟢 low | KPI слайды дублируют "%" в descriptions | Copy Editor правило |

См. `LEARNINGS.md` секция "Backlog known limitations".

---

## 6. Ритуалы перед каждой итерацией

### Read-only
1. Прочитать `memory/MEMORY.md` для контекста
2. Прочитать relevant feedback файлы

### Перед изменением кода/правил
1. Прогнать **регресс на 4 кейсах** (slide_graph / ai_test2 / testpart / Marketing)
2. Брать .pptx из `../Презентации на тест/`
3. Финал — LLM Visual Verifier per-slide

### Перед коммитом
1. Update `LEARNINGS.md` — что сделано, что не работало
2. Если новое user feedback → memory entry
3. Repack zip: `cd .. && zip -r cloud-ru-slides-skill-vX.Y.zip pptx-skill/ -x "pptx-skill/output/*"`

### Версионирование
- patch (v1.3.1) — bug fix без новых donors/категорий
- minor (v1.4) — новые donors, новые правила
- major (v2.0) — архитектурный сдвиг

---

## 7. Roadmap (что планировалось)

### v1.4 (next minor)
- Stress-test на длинных презентациях (Test.pptx 21 sl, Marketing 50MB)
- Расширить slot-map donor 86 + 78
- Smart fallback для KPI с автогенерацией taglines
- Distinguish chart vs photo в image enrichment

### v2.0 (architectural)
- Real LLM-orchestrated pipeline в Claude.app (vs текущая скриптовая batch + emulated workflow)
- HTML-engine pivot (есть начатый `html-engine/` — заброшен в v0.18, возможно вернуться)
- Интеграция с Notion/Figma как источниками
- Auto-feedback loop: при NEEDS_REWORK → автоматический re-do без user

---

## 8. Гитчи (подводные камни)

### Технические
- **`p.slides[i:j]` не работает** — баг python-pptx 1.x. Использовать `list(p.slides)[i:j]`
- **Donors с PNG-заглушками** (53/54, 79, 48) — `remove_before_fill: [shape_idx]` обязателен
- **clone_slide() в build_v8/v9** — не поддерживает дубли donor нативно. Workaround в коде
- **LibreOffice render** иногда обрезает большие presentations — для smoke-test OK
- **Имя файла Test — part.pptx** содержит NBSP (`\xa0`) перед тире — экранировать

### Дизайнерские
- **Зелёный = акцент 5-10%, не доминанта** (canonical §2)
- **Стрелки серые, НЕ зелёные** (canonical §3, slide 66)
- **Тёмные слайды ≤40%** общей презентации
- **199pt KPI = max 2 цифры**, иначе обрезается
- **«Один слайд = одна мысль»** (canonical §7 пункт 2)

### Процесс
- **PASS от 3 авто-валидаторов ≠ готово.** LLM Visual Verifier ОБЯЗАТЕЛЕН
- **Не удалять donor 43/44 из шаблона** — они нужны как guide reference, но НЕ как clone source
- **Output cleanup перед каждой сессией** — иначе 50+ MB мусора накапливается

---

## 9. Вопросы к пользователю

Если новый агент сомневается — уточняет у Глеба:
- «Этот слайд должен быть title или divider?» (если semantic ambiguous)
- «Картинка — главный контент или акцент?» (image_native vs accent)
- «Можно разбить на 2 слайда?» (когда overflow >30%)
- «Сменить donor если визуально плохо?» — да, всегда

---

## 10. Связанные документы

### Внутри pptx-skill/
- `SKILL.md` — главный workflow
- `LEARNINGS.md` — история v0.1 → v1.3
- `README.md` — user-facing
- `brand/template-canonical-rules.md` — canonical правила

### Vault (Obsidian)
- `~/Documents/2nd brain/01 Projects/Cloud.ru Slides Skill/Cloud.ru Slides Skill.md`
- `~/Documents/2nd brain/01 Projects/Все проекты.md` (карта проектов)
- `~/Documents/2nd brain/01 Projects/Экосистема Claude.md` (хронология)

### Memory
- `~/.claude/projects/-Users-gmmelnikov-Desktop-----------------/memory/`

---

## Подпись и контакт

**v1.5 завершено:** 2026-05-05
**Финальный zip:** `../cloud-ru-slides-skill-v1.5.zip`

**Что нового в v1.5 (короткая сводка):**
- Brand validators расширены: зелёные стрелки, цветной текст, доп. цвета без зелёного, лимит цифр в KPI 199pt
- False-positive WARN на charts убран (`visual_validator_v2.py --plan`)
- Единый файл палитры `brand/palette.json`
- Версионирование шаблона `brand/template-version.json` (slide-индексы вынесены из кода)
- Cross-platform пути для LibreOffice (Windows + Linux)
- Regression test suite `tests/regression.py`
- `e2e_pipeline.sh` обновлён под актуальные скрипты

**v1.3 завершено:** 2026-05-04 — baseline качества (4 кейса, 4.7-4.8/5 LLM verdict)
**v1.4 завершено:** 2026-05-05 — editable PowerPoint charts

**Если новый агент видит «PERFECT!» в моих старых сообщениях** — это было до v0.18. После v0.18 правило: **никогда не объявлять PASS без визуального ревью каждого слайда глазами**. Это правило фиксировано в memory `feedback_no_false_positive_pass.md`.

Удачи следующему!
