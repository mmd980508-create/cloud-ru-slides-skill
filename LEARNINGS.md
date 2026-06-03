# LEARNINGS — журнал ошибок и фиксов

> Каждая итерация добавляет сюда: что не работало → что сделали → ссылка на коммит/файл.

## v2.4.2 — 2026-06-02 — Point 3: token-diff/конформанс свёрнут в brand_guardian

Из дизайн-аудита п.3 (token-diff: сверка слайда с бренд-контрактом). По решению
user — НЕ плодить сущности: всё сложено в существующий `brand_guardian`, а не в
отдельный файл (черновой `token_diff.py` удалён).

### ✅ Что добавилось в brand_guardian (это были ПРОБЕЛЫ — гард их не ловил)
- **Shape-ось** (`check_shape_geometry`): скругления `roundRect/snip…` → WARN
  (допустимо только как метафора), эффекты `effectLst/outerShdw/glow…` на фигурах
  → FAIL. Раньше гард проверял текст/цвет/стрелки, а форму фигур — нет.
- **Italic / underline** в text-run → WARN (раньше ловил только bold-флаг).
- **`axis_report(pptx)`** — группирует все находки гарда по осям
  **Color / Typography / Shape** (формат MATCH/DIVERGE для scorecard/печати).
  Единая точка чеков — не дублирует логику палитры/шрифтов.

### Проверено
- off-brand `before.pptx` → FAIL (ловит roundRect, effectLst, Arial, italic, не-палитру);
- нормализованный `after.pptx` → PASS;
- реальный сгенерённый деком → 94.3/100, новые чеки БЕЗ ложных срабатываний
  (enforce_canonical уже снял эффекты, доноры прямоугольные). FAIL там — от
  пред-существующего `color_off_palette`, не от новых проверок.

### Тест
`tests/point3_tokendiff/build.py` — Cloud.ru-scorecard БЫЛО/СТАЛО (данные из
`brand_guardian.axis_report`).

## v2.4.1 — 2026-06-02 — FIX: clone_slide портил файл (PowerPoint repair + слетали картинки)

### ⚠️ Симптом
Любой сгенерённый деком в PowerPoint (Mac) открывался с окном «обнаружена проблема
с содержимым → Восстановить», а после — слетали ВСЕ картинки из шаблона.
LibreOffice открывал без ошибок (он прощает), поэтому раньше не ловили.

### 🔬 Диагностика (бисекция через user в PowerPoint)
- TEST-A: чистый `Presentation(template).save()` → открывается ОК (python-pptx не виноват).
- TEST-B: clone + удаление оригиналов → ломается. → баг в `clone_slide`.
Две причины, маскировавшие друг друга:
1. **rId-перенумерация (главная, «слетают картинки»).** `clone_slide` копировал
   slide-XML блобом (с ИСХОДНЫМИ rId), а связи воссоздавал через `relate_to`, который
   **перенумеровывает rId** по порядку обхода. В итоге `<a:blip r:embed="rId3">`
   указывал на `slideLayout` вместо изображения → PowerPoint «не смог прочитать
   часть содержимого» и удалял картинку. (Наивный dangling-чек это не ловил: rId
   существовал, просто указывал не туда.)
2. **orphan-слайд (окно «Восстановить»).** Копировалась связь с `notesSlide` донора;
   notesSlide ссылается НАЗАД на оригинал-слайд → оригинал «застревал» в пакете
   (есть в файле, нет в sldIdLst) → PowerPoint repair.

### ✅ Фикс (`build_v9.py` + `build_v7/v8`)
`clone_slide` теперь копирует связи **СОХРАНЯЯ оригинальные rId** (вставка в
`new_part.rels._rels[rId] = _Relationship(...)`, а не `relate_to`), и **НЕ копирует**
`notesSlide`-связь. Так rId-ссылки слайда остаются валидными, а оригинал-донор
становится недостижимым и корректно отбрасывается при save.
(Зависит от внутренностей python-pptx 1.0.2: `_rels` / `_Relationship` / RTM.)

### 🛡️ Регрессия-страховка
Новый `scripts/validate_deck.py` — структурный валидатор (orphan-слайды, blip→не-media,
dangling rId, missing target, stray notes, content-type). `build_v9` зовёт его после
save и печатает `✅ DECK VALIDATION: структурно чисто` / список проблем. Проверено:
ловит баг на старых деках, пропускает исправленные.

## v2.4 — 2026-06-02 — Токен-контракт + brand-normalizer + масс-иерархия цвета

Из дизайн-аудита внешних скиллов (designaudit на десктопе) внедряли по пунктам.
Дошли до п.2 (по решению — чекпоинт-коммит здесь).

### ✅ Point 1 — токен-контракт (единый источник констант)
- **`brand/design-tokens.yaml`** — машиночитаемый контракт НЕ-цветовых констант
  (типографика/сетка/геометрия/размещение лого+заголовка+копирайта/компоненты).
  Loader **`scripts/design_tokens.py`** сливает palette.json + tokens; parity-self-check
  (`--check`) сверяет токены с прежними хардкодами → 22/22 (миграция безопасна).
- **Логотип** контент-слайда = КОНСТАНТА из мастера шаблона: (1096,38, 149×29),
  right=1245, y=38 (= как у header). Раньше в токене был только finale — противоречие снято.
- **Копирайт-футер** = КОНСТАНТА из мастера: (35,696) 6pt #BFBFBF, точный текст.
- Чип 40×40 / 12pt; 2 варианта карточек (`default` chip-left / `text_heavy` chip-right,
  заголовок вровень с чипом + перенос при наложении, тело ниже динамически).
- **Тени УБРАНЫ намертво ВЕЗДЕ**: `scripts/effects_util.strip_effects` снимает и
  явные `effectLst/effectDag`, и тематический `<p:style><a:effectRef>` (idx=0) —
  причина «фантомных» теней на blank-канвасе. Подключено в `_no_effects` всех
  рендеров (flow/kpi/table), в коннекторы декор-стрелок и в `enforce_canonical`.
  Скан тест-деков = 0 эффектов.
- **Зелёное правило** сужено (user): запрещены только «недолинии» — висящие декор-
  черты, не привязанные к сетке/композиции/акценту (подчёркивание под заголовком).
  Green-детали/заливки/чипы/цифры/композиц-линии — ОК; бренд-стрелки ↗ green ИЛИ
  серые. Тонкая зелёная полоса как акцент — нельзя (слаба); акцент = заливка-блок.

### ✅ Point 2 — extraction-layer: нормализатор «чужого скина»
- **`scripts/brand_normalizer.py`** — снап импортированного оформления к токенам:
  скругления→прямые, эффекты→сняты, градиент→сплошной, чужой цвет→палитра,
  чужой шрифт→SB Sans Display, italic/underline→сняты, bold-флаг→Semibold; контраст
  делегирован `enforce_canonical` (white-on-green→графит). CLI + importable.
- **Масс-иерархия цвета** (брендбук + user 2026-06-02): White ≫ greys > **Green
  (главный акцент)** > Yellow/Purple/Blue (каждый ≤ Green); Graphite = текст.
  `design-tokens.color_mass` + `brand-rules.md`. Официальная шкала серых добавлена
  в palette.json: **Grey1 #F2F2F2 / Grey2 #D9D9D9 / Grey3 #434343**.
- Нормализатор учитывает массы: чужой насыщенный цвет → нейтраль по умолчанию,
  green — единственный акцент; слайд-кап **Σ(доп.цвет) ≤ green** (избыток → серый).
- **`scripts/color_mass_audit.py`** — измерение цвета по площади. На 3 реальных
  деках: White 74–84% (подтверждает «белого больше всего»).

### Тесты
`tests/point1_tokens/`, `tests/point1_greenlines/`, `tests/point2_normalizer/`
(build.py + рендеры; png/pptx в .gitignore).

### ⚠️ Отложено
- Композиция архетипов KPI/hero/numbered (плавающие пустоты, слабые акценты) → Point 7.
- Полный скрейпер референс-колоды (извлечение структуры) — v3.0.
- Точное измерение green-vs-additional по массе (живут в тексте/теме — метрика груба).

## v2.3 — 2026-06-01 — Пресеты дизайн-архетипов (из эталонных дек)

### ✅ Что сделали
Из сравнения драфт→финал (`brand/design-principles-from-decks.md`) собрали 5
готовых пресетов-композиций, чтобы драфт сразу выходил на уровне финала. Все
проверены рендером LibreOffice на blank-доноре 30.

- **`flow.preset`** (`render_flow_diagram_slide` диспатчит после header/subtitle):
  `numbered_rows` (5B), `card_grid` (4), `numbered_columns` (5A), `hero_statement` (6).
- **`table.preset="before_after"`** (8, `table_renderer.py`). Рисуется из плашек,
  не из PPT-таблицы.

### ⚠️ КРИТИЧЕСКАЯ ОШИБКА → ФИКС (correction 1+2, сверка по референсам 2026-06-01)
Первая версия пресетов была собрана по тексту принципов (где я САМ ошибся:
«белый номер на зелёном»), а не по реальным эталонным картинкам. Пользователь
указал 5 расхождений. Сверил 1-в-1 по референсам и переписал:

1. **Цифры/иконки на зелёном — ГРАФИТ, regular** (white-on-green запрещён).
   Было: белые bold-цифры. `add_green_chip` теперь графит regular.
2. **Иконка-галочка = контурный кружок + галочка графитом** (`_draw_check`).
   Было: голая line-art галочка не той формы. Теперь OVAL-обводка + ✓ внутри.
3. **3-уровневая иерархия**: заголовок SemiBold графит, тело `TEXT_GRAY #5C5C5C`
   regular. Добавил константу `TEXT_GRAY` + параметр `text_colors=[...]` в
   `add_block` (цвет по строкам). Раньше всё было одним графитом.
4. **Отступы между фреймами `PRESET_GAP=4`** (≤10px). Были крупные зазоры.
5. **Раскладки переписаны по эталонам:**
   - `numbered_rows` — чистый стиль по умолчанию (номер графитом + заголовок
     SemiBold в строку, тело серым ниже, `cols=2`), БЕЗ лент/чипов. Старый
     вариант с лентами+чипами доступен через `style:"band"`.
   - `card_grid` — заголовок SemiBold + тело серым; чип ОПЦИОНАЛЕН (check/number),
     при наличии чип+заголовок в ОДНУ строку, тело ниже.
   - `numbered_columns` — БЕЗ зелёной черты сверху (default), крупный (~56pt)
     **REGULAR** номер снизу, тело серым. Черта возвращается `rule:true`.
   - `before_after` — табы-заголовки ЦВЕТНЫЕ с текстом по ЛЕВОМУ краю во всю
     ширину колонки; лёгкая серая подложка под «Было»; тонкие серые (#CCC)
     разделители строк; метрика двухстрочная (название SemiBold + деталь серым).
     Убрал полноразмерную зелёную hero-плашку «Стало» — её нет в эталоне табличного
     before/after (она относится к hero_statement).

**Урок:** пресеты сверять ТОЛЬКО по реальным картинкам-референсам, не по
пересказу принципов по памяти.

### ⚠️ Нюанс: hero overflow
Фикс. высота зелёной плашки давала вытекание текста за блок. Решение: высота
считается под число строк (`_wrapped_lines`, char_w=0.66·font_px для CAPS), при
нехватке места кегль уменьшается до 32pt (`block_bottom_max=545`).

### Осталось (рисовать вручную по принципам)
Архетипы 1 (матрица N×M), 2 (персона+фото), 3-доп (узлы по заливке + phase-underline),
7 (код-панель с узким заголовком).

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

## v1.7 — 2026-05-26 — Flow diagrams: canonical composition rules

**Контекст:** работали над двумя реальными слайдами (SGR ReAct + Pipeline аналитики). Пользователь дал серию итераций по правилам дизайна → выработали canonical правила, теперь зафиксированы в `brand/template-canonical-rules.md` §3 и в `scripts/flow_renderer.py`.

### Изменения в `scripts/flow_renderer.py`
- `ARROW_COLOR = #434343` (раньше графит #222222 чёрный) — единый цвет всех стрелок.
- `ARROW_WIDTH_PT = 1.0` (default), убран по умолчанию `w_pt=1.0` в `add_arrow`.
- `add_arrow`: arrowhead `type="arrow"` (open chevron) + `w="lg" + len="med"` = PowerPoint size 8. Раньше было `triangle/sm/sm` (size 1).
- `add_arrow`: добавлена **валидация диагоналей** — `ValueError` при `x1 != x2 and y1 != y2`. Принудительно ломаные через 90°.
- `add_block`: default `align="left"` и `vanchor="top"` (раньше center/middle). Поля унифицированы 12/12/12/16 (нижнее больше).
- Добавлены safe-area константы: `SAFE_TOP=140, SAFE_BOTTOM=660, SAFE_LEFT=30, SAFE_RIGHT=1250`.

### Изменения в `brand/template-canonical-rules.md` §3
Подраздел «Flow-схемы и блок-диаграммы»:
- Правило растягивания: многорядная (>1 ряд И >3 колонок) → растянуть на safe-area; линейный pipeline → естественная высота + центрирование.
- Чеклист после вёрстки: текст влезает → формируем .pptx; нет → расширяем фреймы → уменьшаем шрифт ≥10pt → пересмотр композиции.
- Параметры стрелок (1pt, #434343, type=arrow size 8, только 90°).
- Выравнивание текста в блоках (LEFT + TOP).
- Карта размеров arrow в OOXML (size 1-9).

### Уроки в memory
- `feedback_flow_diagram_composition.md` — все правила композиции + карта стрелок.

### Анализ «почему был непредсказуемый результат»
До v1.7 я думал «контент → размер», добавлял блоки сверху вниз, и они занимали столько px сколько занимали. Получался пустой низ слайда. Правильно — «доступная область → распределить контент по ней». Введены safe-area константы — теперь композиция стабильная и предсказуемая.

### Расширение v1.7 (2026-05-26, основная сессия) — schema-in-image recognition + sync

**Проблема (наблюдалась 2026-05-26):** в реальном прогоне на драфте пользователя классификатор по умолчанию выбирал `image_native` для embedded PNG, который сам по себе был схемой (блоки + стрелки). Результат: схема оставалась статичной картинкой вместо редактируемого `flow_diagram_native`. Это прямое нарушение canonical v1.6 («все диаграммы должны быть редактируемыми»).

**Что сделано:**

- `agents/02-slide-classifier.md` — добавлен раздел **«Распознавание схем в картинках»**:
  - Триггер расширен: embedded PNG/JPG, который выглядит как схема (≥2 из: прямоугольные блоки с текстом / стрелки между ними / lanes-phases / направленный граф) → НЕ `image_native`, а `flow_diagram_native`.
  - LLM vision: классификатор смотрит на extracted PNG (через `extract_images.py`) и реконструирует структуру в JSON-конфиг (blocks/arrows/groups/labels).
  - Стиль применяется canonical-flow (gray fills, тёмно-серые стрелки, и т.п.) — цвета из исходной картинки игнорируются.
  - Если сомнение «схема или нет» — спросить пользователя.
  - Множественные схемы на одном слайде → split на N output-слайдов.

- `agents/02-slide-classifier.md` — также добавлены полный canonical-блок композиции v1.7 и обязательный пост-рендер чеклист.

- `SKILL.md` — раздел «Canonical правило v1.6 — editable schemas» обновлён до v1.7: пункт про schema-in-image, обновлённый стиль (текст LEFT/TOP, стрелки #434343 size 8, safe-area константы), правило растягивания, чеклист.

- `scripts/flow_renderer.py` — мелкие финальные правки поверх композиционных:
  - default `align="left"` и `vanchor="top"` в `render_flow_diagram_slide` (раньше «center»).
  - default `w_pt=None` в передаче `add_arrow` (использует `ARROW_WIDTH_PT` константу автоматически).
  - Header docstring обновлён под v1.7 canonical (раньше упоминал «triangle, w=sm len=sm»).

- `tests/baselines/flow_diagram/plan.json` — координаты переписаны под safe-area, размеры блоков увеличены под новые поля 12/16, добавлен подкомментарий «линейный pipeline в 1 ряд → НЕ растягиваем».

**Sync infrastructure (правило «закоммить» 2026-05-26):**
- При коммите в `pptx-skill/` теперь автоматом синкается **установленная копия** (`~/Library/Application Support/Claude/.../skills/cloud-ru-slides/`) и обновляется **zip-снимок** на десктопе. См. `feedback_commit_workflow.md`.

## v1.8 — 2026-05-26 — table_native (zebra) + anti-distortion safety

**Контекст:** при тесте v1.7 на реальном файле `cloud-ru-schema-ai-yellow.pptx`
(маркетинговая воронка из 28+ блоков) обнаружилось что:
1. Скилл не умеет делать **настоящие PowerPoint таблицы** — только плашки через flow_diagram_native, без возможности редактировать таблицу как таблицу в PowerPoint.
2. Скилл **самостоятельно искажает смысл** при упрощении: split схемы на 3 слайда был «принят без спроса», хотя пользователь хотела всё на одном слайде с табличной компоновкой.

Также пользователь утвердила выбор canonical стиля для таблиц по slide 56
шаблона Cloud.ru (zebra) и отказалась от прочих вариантов (53 gradient, 54
expanded-row, 55 kpi-table, 57 RACI отдельно, 58 roadmap отдельно).

### Изменения

**Новый slide_type `table_native`:**
- [scripts/table_renderer.py](pptx-skill/scripts/table_renderer.py) — новый модуль:
  - `render_table_native(slide, table_config, dark)` собирает native PowerPoint таблицу через `slide.shapes.add_table()`.
  - Zebra style как в slide 56: header row без заливки + bold 12pt, body rows чередуются `#F2F2F2` / белый.
  - Vertical separators 0.5pt `#C8C8C8` между колонками (только справа от ячеек, кроме последней). Горизонтальных границ нет.
  - Поля ячеек: L/R 12px, T/B 8px.
  - Текст: SB Sans Display, **align=left, vanchor=top** (canonical, не override).
  - Первая колонка 1.4× шире (опц. `first_col_wider: true`).
  - Авто-расчёт высот: если `h` не задана — заполнить до safe-bottom; `row_height` = (avail_h - header_h) / n_rows.
  - `_strip_default_table_style()` убирает дефолтный applied style чтобы наши explicit fills не перебивались темой.
  - Standalone CLI: `python3 table_renderer.py <config.json> <template.pptx> <out.pptx>`.

- [scripts/build_v9.py](pptx-skill/scripts/build_v9.py):
  - Import `render_table_native` (try/except → `TABLE_RENDERER_AVAILABLE`).
  - `table_native` в списке native types (blank donor).
  - Новая ветка обработки между `flow_diagram_native` и end of native handlers.

- [scripts/validate_plan.py](pptx-skill/scripts/validate_plan.py):
  - `table_native` принят как native (требует поля `table`).
  - Проверки: header/headers/data наличие, n_cols ≥ 3 (canonical триггер), n_rows ≥ 3, длина каждой строки == n_cols, bounds (x..x+w в 0..1280; y+h в 0..720).
  - **Center alignment warning** в flow_diagram_native: если block имеет явный `align: "center"` или `vanchor: "middle"` — WARN (canonical нерушим, требует обоснования через header-плашки).

- [agents/02-slide-classifier.md](pptx-skill/agents/02-slide-classifier.md):
  - Раздел про `table_native` triggers: ≥3 cols × ≥3 rows + явная шапка + intent comparison/pricing/spec.
  - **Anti-distortion stop+ask раздел** — обязательная остановка при merged cells / irregular grid / RACI / color-coded ячейках / nested tables.
  - **Распознавание таблиц в картинках** (по аналогии со схемами v1.7): эвристика «таблица или нет» через LLM vision → реконструкция в `table_native` headers + data. Если merged cells в картинке → stop+ask.
  - Routing decision tree обновлён: первым шагом проверка anti-distortion триггера.

**Anti-distortion safety rule:**
- [memory feedback_anti_distortion_safety.md](memory) — новое правило: при встрече нестандартного объекта (merged cells, RACI, roadmap, color-coded, brand-объекты клиента, hand-drawn) → СТОП → описать → объяснить риск → 2-4 варианта → ждать explicit-решения пользователя. Запрет: решать самостоятельно.
- MEMORY.md: добавлена ссылка.
- Triggers подробно (с примерами): табличные структуры, cell-level семантика, спец-шаблоны (RACI/roadmap/quadrant/funnel/org chart), графики, брендирование, семантика положения.

**Документация:**
- [SKILL.md](pptx-skill/SKILL.md):
  - Таблица скриптов: новый `table_renderer.py` ⭐.
  - JSON-схема `table_native`.
  - Раздел «Canonical правило v1.8 — table_native (zebra) + anti-distortion».
  - В list slide_types: добавлен table_native + anti-distortion-stop.

**Что НЕ поддерживается v1.8 (по решению пользователя):**
- slide 53 (зелёный gradient) — пропущен.
- slide 54 (expanded row) — экзотика, пропущена.
- slide 55 (KPI table) — покрывает `kpi_native`, не делаем дубликат.
- slide 57 (RACI matrix) — anti-distortion stop+ask (предлагаем альтернативу), не отдельный style.
- slide 58 (Roadmap timeline) — отдельный slide_type позже (не таблица по природе).

**Уроки в memory:**
- `feedback_anti_distortion_safety.md` — новое общее правило безопасности.

### Анализ «почему скилл сам решал split»
До v1.8 классификатор «знал» правило «≤8 блоков → split» и применял молча.
Это привело к разрыву смысла маркетинговой воронки на 3 слайда там, где
пользователь явно хотела всё на одном. Уроки v1.8:
1. «Перегруженность» не всегда повод для split — иногда это естественная плотность контента (например, табличный коллаж).
2. Любое крупное structural решение (split, переоформление, замена slide_type) при ambiguity → anti-distortion stop+ask, не silent decision.

## v1.9 — 2026-05-29 — Native PP bullets «Заполненный квадрат» (Problem #1)

**Контекст:** строки-списки в flow/table-блоках рисовались как текстовые маркеры.
Первая версия использовала символ `▪` (U+25AA) со шрифтом Arial — в части
рендереров (LibreOffice/macOS) глиф подхватывал **emoji-вариант** и давал
«эффект объёма» (цветной/3D квадрат). Пользователю нужны были штатные плоские
PP-маркеры «Заполненные квадратные маркеры», цвет по умолчанию `#434343`, без
эффектов.

### Изменения

**Auto-detect bullets в обоих рендерерах:**
- [scripts/flow_renderer.py](pptx-skill/scripts/flow_renderer.py) и
  [scripts/table_renderer.py](pptx-skill/scripts/table_renderer.py):
  - Строка, начинающаяся с `▪ ` / `■ ` / `• ` / `- ` / `* `, → native PP bullet,
    префикс вырезается из текста.
  - Маркер = штатный PowerPoint «Заполненный квадрат»: шрифт **Wingdings**,
    символ `§` (0xA7). Это ASCII-символ → emoji-фолбэка не бывает, всегда плоский
    залитый квадрат. НЕ использовать `▪` U+25AA (отсюда был «объём»).
  - Цвет маркера задаётся явно: `<a:buClr><a:srgbClr val="434343"/></a:buClr>`.
  - Hanging indent `marL=228600 / indent=-228600` (~24px).
  - Порядок дочерних элементов pPr по схеме OOXML: `buClr → buFont → buChar`.
  - Смешанные блоки: строки без префикса остаются без маркера
    (`_apply_no_bullet_to_paragraph` ставит `<a:buNone/>` и убирает indent).
  - Это настоящие native-буллеты — в PowerPoint редактируются через «Маркеры».

### Урок
Для символьных буллетов в OOXML использовать **Wingdings + ASCII-символ**, а не
Unicode-символы вроде `▪`/`■`/`●`: последние в кросс-платформенном рендере могут
получить emoji-presentation и «3D»-вид. Wingdings гарантирует плоский глиф.

## v2.0 — 2026-05-29 — Полировка стиля: контраст / SemiBold / стрелки / размер / заголовки (Problems #2–6) + фикс тёмного донора

Батч из 5 пунктов «довести вёрстку до бренда» + один найденный по пути баг.

### Problem #2 — запрет зелёного/белого текста (контраст)
- **Правило:** текст ВСЕГДА `#222222` на белом И на зелёном фоне. Белый текст —
  только на тёмном (#222222/чёрный). Зелёный как ЦВЕТ ТЕКСТА запрещён везде
  (вкл. крупные KPI/divider-цифры). Акцент — отдельным цветным ЭЛЕМЕНТОМ.
- [flow_renderer]: зелёная плашка `green` → текст графит (был белый).
- [kpi_renderer]: акцентная цифра → графит/белая (была зелёная) + зелёная
  плашка-подчёркивание `_add_accent_bar` над цифрой как акцент.
- [palette.json] + [brand_guardian]: убран `#343F48` (Graphite-Iron) — не из
  палитры Cloud.ru; тёмная плашка теперь `#222222`. `text_dark_extra` = только
  белый (зелёный текст не разрешён даже на тёмном).
- [brand_guardian]: WARN `colored_text` на зелёный/белый текст; распознаёт тёмную
  ПЛАШКУ (по luminance заливки) → белый текст на тёмной плашке легитимен.
- Документация: brand-rules §4, SKILL, агенты 04/06/08.

### Problem #3 — Bold → SemiBold
- Эмфаза = начертание **`SB Sans Display Semibold`** (встроено в шаблон), а не
  bold-флаг. Все рендереры (flow/table/kpi/image/chart) переведены: `bold=True`
  → face Semibold + `bold=false`.
- [render_slides]: LibreOffice не различает SemiBold (схлопывает в Regular),
  поэтому для PNG-превью делается временная подмена Semibold→Regular+bold (только
  во временной копии). В .pptx остаётся настоящий SemiBold.
- [brand_guardian]: WARN `bold_flag` на `b=1`.

### Problem #4 — фирменные декор-стрелки ↗
- Было: искажённые треугольники с диагональю ↘. Стало: фирменная стрелка ↗
  (древко вверх-вправо + уголок-наконечник), геометрия из `brand/icons/brand_arrow.svg`.
- Каждая стрелка = **нативная редактируемая фигура-группа PowerPoint** (`p:grpSp`
  из линий-коннекторов), а не картинка/SVG — двигается/перекрашивается сразу,
  без «Преобразовать в фигуру». Линия 1pt.
- Утилита `add_svg_picture` (вставка SVG как editable-вектора с PNG-fallback)
  оставлена в flow_renderer для возможных будущих нужд (логотипы/иконки).

### Problem #5 — размер шрифта: стандарт 16pt, мин 12pt
- Дефолты: flow-блоки/подписи **16pt** (было 11-13); subtitle 13 / url 12 /
  лейблы фаз 12 (было 9-11); таблицы — **дефолт 16pt с авто-уменьшением**
  (`_autofit_table_font`: подбирает крупнейший влезающий размер, 16→12 комфортно,
  до 10 крайне; форс `table.font_size`).
- [validate_plan]: overflow-ужатие не ниже 12pt (было 10).
- [brand_guardian]: WARN `size_below_comfortable` < 12pt, FAIL `size_too_small`
  < 10pt; стандарт 16pt.

### Problem #6 — заголовок в штатный TITLE-placeholder
- Было: каждый рендерер рисовал свой header-textbox (35,38 vs 35,60; 20 vs 32pt;
  CAPS/не-CAPS). Стало: единая `kpi_renderer.set_slide_title()` вписывает заголовок
  в **штатный TITLE-placeholder шаблона** (canonical 35,38 / 963×54 / 20pt
  SemiBold CAPS; графит/белый). `clean_slide_to_blank(keep_title=True)` сохраняет
  placeholder (очищая донорский текст). Нет placeholder → fallback на textbox той
  же геометрии. Все 5 рендереров переведены на set_slide_title.

### Фикс по пути — тёмный blank-донор
- `BLANK_DONOR_DARK` был 22, но слайд 22 — белый layout → `dark:true` давал белый
  фон (белый текст невидим). Тема шаблона: `dk1=#222222`, `tx1→dk1`. Настоящий
  тёмный blank-донор — **слайд 51** («Контент / темный 1», bg #222222, TITLE-ph в
  (35,38)). Обновлён `brand/template-version.json` (dark=51) + fallback в
  kpi_renderer. Тёмные flow/KPI теперь рендерятся на #222222 с читаемым белым.

### Урок
SemiBold/Medium в OOXML задаются ТОЛЬКО именем начертания (нет b-значения для
semibold), а LibreOffice их не различает в превью — для служебных PNG нужна
подмена. Цвет фона тёмных слайдов идёт из темы (`tx1→dk1`), а не из заливки
слайда — «тёмность» донора надо проверять по резолву темы, не по имени слайда.



## v2.1 — 2026-05-29 — enforce_canonical для clone-слайдов (контраст/вес/заголовки) + честность

**Контекст:** пользователь обновила скилл и протестила в чате — Problems #2/#5/#6
всё ещё видны. Разбор реального файла из чата показал: презентация на 100%
**clone-based** (клоны шаблонных доноров), а фиксы v2.0 живут в native-рендерерах
→ до clone-слайдов не доходят. Аудит: 2 green-text, 14 white-on-green, 82 runs
<16pt, 34 bold-флага.

### Что сделано
- **`scripts/enforce_canonical.py`** — финальный canonical-пасс по ВСЕМ слайдам,
  вшит в конец `build_v9`. Делает ТОЛЬКО безопасное:
  - цвет: зелёный/белый текст → `#222222` (белый остаётся на тёмном фоне/плашке);
  - вес: bold-флаг → начертание `SB Sans Display Semibold`;
  - размер: `<12pt → 12pt`;
  - **заголовок** контент-слайда → штатный TITLE-placeholder (35,38)/20pt SemiBold
    CAPS, **семантически** (`normalize_header_to_placeholder`): текстбокс-заголовок
    переносится в placeholder; титульные/divider/вертикальные/огромные — не трогаются.

### Грабли (важно — было две регрессии)
- Первая версия «умного» пасса **сама сломала макеты**: bump 12-15→16pt → код-блок
  за слайд (slide 8); поиск заголовка «по позиции» (любой крупный текст сверху) →
  схватил вертикальный сайдбар → заголовок в столбик (slide 6). ОТКАТИЛА.
- Урок: на clone-слайдах **нельзя** двигать/масштабировать вслепую. Заголовок —
  только СЕМАНТИЧЕСКИ (через title-placeholder), а не по позиции. Bump до 16pt и
  autofit — только в NATIVE-пути, не пост-фиксом.
- **Корень «скачущих заголовков»:** генерация иногда кладёт заголовок в отдельный
  ТЕКСТБОКС, а не в штатный TITLE-placeholder (который пустой) → разные позиции.
  Шаблон тут ни при чём — он единообразен. Фикс = всегда класть в placeholder.

### Архитектурный вывод (в памяти feedback_clone_vs_native_gap.md)
Гарантированный канон (16pt, autofit-без-overflow, единые заголовки) возможен
только в native-пути. Clone-слайды лечатся постфактум лишь по цвету/весу/заголовку.

### Поведение
Пользователь попросила ВСЕГДА быть честной про ограничения/регрессии и не делать
слепых правок (в памяти feedback_honest_no_blind_changes.md). Каждый затронутый
слайд проверять рендером.

## v2.1 (продолжение) — заголовки в placeholder + грид-композиция схем

Сессия с пользователем по реальному файлу из чата. Большие уроки.

### Заголовки контент-слайдов → в штатный TITLE-placeholder
- **Корень «скачущих заголовков»:** генерация кладёт заголовок в отдельный
  ТЕКСТБОКС (в произвольной позиции), а штатный title-placeholder пустой. Шаблон
  единообразен — виновата генерация.
- **Грабли:** трогать placeholder (set position/clear+rewrite) НЕЛЬЗЯ — в PowerPoint
  он схлопывается до минимальной ширины → заголовок встаёт ВЕРТИКАЛЬНЫМ столбиком.
  **LibreOffice-рендер это НЕ показывает** (поймали только в PowerPoint у юзера).
- **Решение (`enforce_canonical.normalize_header_to_placeholder`):** если заголовок
  уже в placeholder — НЕ трогать. Если в текстбоксе — `_ensure_title_placeholder`
  инстанцирует placeholder из layout (полная/наследуемая геометрия), вписываем ТОЛЬКО
  текст, текстбокс удаляем. Геометрия не трогается → не схлопывается.
- Аудит «элементы не потерялись»: единственное изменение `TEXT_BOX−1/PLACEHOLDER+1`
  (заголовок переехал), все стрелки/картинки/фигуры 1:1.

### Грид-композиция схем (flow_renderer)
- **autofit (сжатие текста) ОТВЕРГНУТ** пользователем: в схеме блоки ужимаются
  по-разному → разнобой кеглей = дизайн-ошибка. Надо: рамка под текст при ЕДИНОМ кегле.
- `compose_grid(blocks, area, cols, font_pt, gap)`: блоки логические (row/col/lines),
  колонки равной ширины, высота строки = под самый длинный блок (frame-to-text),
  единый кегль. `flow.grid=true`.
- **Branching** (пайплайн → N выходов): диагональ запрещена (canonical), поэтому
  `add_orthogonal_arrow` — Z-маршрут (вправо→верт→вправо). Авто-сторона по позициям
  ячеек. Без этого вертикальные/диагональные ветки «пропадали».
- **Микроправило:** резерв под наконечник на финальном сегменте (`ARROW_ENTRY_RESERVE`)
  + больший зазор колонок для branching (`GRID_GAP_BRANCHING`) — окончание стрелки видно.
- **Нагруженные схемы:** залитые серые ПАНЕЛИ (`groups[{"style":"panel"}]`,
  `add_filled_panel`, блоки внутри `fill:"white"`) вместо пунктира — чище.

### Поведение (в памяти)
Всегда честно про ограничения/регрессии (две регрессии за сессию: bump→overflow,
перенос заголовка вслепую→вертикаль). Не делать слепых правок макета. Проверять
рендером; помнить, что LibreOffice не показывает placeholder-схлопывание и normAutofit.

## v2.2 — 2026-06-01 — Авто-резолв шаблона + sync_template для обновлений

### Авто-подтягивание шаблона (убрали ручную загрузку в чат)
- `scripts/template_path.py` → `resolve_template()`: ищет шаблон по приоритету
  env `CLOUD_RU_TEMPLATE` → `template-version.json → template_path` → известные пути.
  В командах: `TPL=$(python3 scripts/template_path.py)`.
- **Ограничение среды:** работает только где файл достижим (локальный агент-режим).
  В облаке claude.ai диск десктопа недоступен → SKILL.md предписывает остановиться
  и попросить загрузить шаблон, если резолвер вернул ошибку.

### sync_template.py — пересборка при обновлении шаблона (неск. раз в год)
- Замены .pptx НЕ достаточно. Авто: `template-layouts-dump.json` + `template/png/`.
  Руками/LLM: donor-индексы, `donor-slot-map.yaml`, `template-analysis.md`.
- `render_slides.py` пишет `slide-NN.png`, а агент-верификатор ищет `Слайд{N}.png`
  → sync_template переименовывает после рендера. Эталоны теперь LibreOffice
  (консистентно с pipeline), а не PowerPoint-экспорт.

### Находка: индексы donor «съезжают» при обновлении шаблона
- Шаблон обновлён 05-26, дампы были от 05-01 (102→84 layouts, 88→91 слайдов).
- **KPI-блок сдвинулся +1:** kpi_white 43→44, kpi_dark 44→45 (слайд 43 стал
  дивайдером). Поймано визуальной сверкой рендера. Таблицы 53/54 НЕ сдвинулись,
  blank-доноры 30/51 валидны. → При каждом обновлении сверять доноры рендером.
- Регресс: нативный KPI на доноре 30 = чистый корректный слайд ✓.

### НЕ доделано (флаг)
- `template-analysis.md` описывает старые 102 layouts (вкл. 20 удалённых паттернов).
- `donor-slot-map.yaml` (clone-доноры по slide-index) не пере-валидирован послайдово.

## v2.2 — 2026-06-01 — Выравнивание фреймов по направляющим + убран дивайдер

Фидбэк по рендеру: (1) фреймы выходили за поля, заголовок и фреймы по разным
границам; (2) лишний дивайдер под заголовком; (3) фреймы не выровнены и справа
пустое место.

- **Safe-area унифицирована с направляющими PP:** `SAFE_LEFT 30→35`,
  `SAFE_RIGHT 1250→1245` в `flow_renderer.py` И `table_renderer.py`. Теперь
  совпадает с TITLE-placeholder (@35) и направляющими PowerPoint → заголовок и
  фреймы по ОДНИМ границам, нет выхода за поля. Grid-режим (`compose_grid`)
  автоматически доводит крайнюю колонку до 1245 (нет пустого места справа).
- **Дивайдер под заголовком убран из дефолта.** Авто-вызов `add_top_separator`
  (flow) и `_add_top_separator` (table) удалён. Включается только флагом
  `top_separator=true` в config — для отделения абзацев в спец. типах слайда.
  Заголовок всегда в safe space через TITLE-placeholder.
- **`snap_panel_to_safe(x,y,w,h,tol=24)`** (flow_renderer): притягивает края
  hand-coded панелей (`groups`) к направляющим, если они РЯДОМ (≤24px). Чинит
  выход за поля и недоведённые до края панели, НЕ трогая намеренные отступы
  (не двигаем чужой макет вслепую). Применяется к panel-фону и dashed-rect.
- Проверено рендером: 3×2 grid доходит до 1245 без дивайдера; панель с
  x=28/right=1255 (за полями) притянулась к 35/1245. ✓
