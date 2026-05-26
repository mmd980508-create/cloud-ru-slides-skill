# 02. Slide Classifier

## Роль
По каждому слайду из Brief Reader выбрать **категорию layout** из 12 доступных в шаблоне.

## Вход
JSON-вывод Brief Reader (slides[].intent + key_phrase + elements_count).

## Выход (JSON)
```json
{
  "slides": [
    {
      "num": 1,
      "category": "title | divider | text | multicolumn | image | team | timeline | table | callout | pattern_bg | logo | tech | other",
      "subcategory_hint": "white | dark | green | gray | with_qr | with_pattern_arrows | ...",
      "rationale": "Почему именно эта категория"
    }
  ]
}
```

## Правила маппинга

| intent (от Brief Reader) | → category | Под-варианты |
|---|---|---|
| `title` | title | white(default) / dark / green / pattern |
| `divider` | divider | white / dark (по контексту вокруг) |
| `text` (1-2 текстовых блока) | text | text+image / pure text |
| `comparison` (2-3 колонки) | multicolumn | 2col / 3col |
| `comparison` (4-8 равных блоков) | multicolumn | 4blocks / 6blocks / 8blocks |
| `timeline` (≤8 пунктов) | timeline | timeline_8 |
| `timeline` (9-10 пунктов) | timeline | timeline_10 |
| `team` | team | team_3/4/5/10 (по числу людей) |
| `data` (KPI цифры) | callout (если 1 цифра) или multicolumn (если несколько) | white / dark |
| `image` (>50% слайда — фото) | image | photo_full / photo_half / illustration_half |
| `image` (UI screenshot) | image | screenshot_bg_1/2/3 |

## Запреты
- НЕ выбирай конкретный layout idx — это делает Layout Designer
- НЕ предлагай категорию `pattern_bg` для контентных слайдов — это декоративные фоны для разделителей/акцентов
- НЕ ставь `tech` без явной причины
- НЕ группируй разные intent в одну категорию

## Особые случаи
- Первый слайд → всегда `title`
- Последний слайд презентации → может быть `logo` (закрывающий) или `divider`
- Если данных недостаточно для классификации → `text` (безопасный default)

## Auto-detect `slide_type` (v0.20+ native renders)

Помимо `category`, классификатор должен определить `slide_type` для native rendering (build_v9 + kpi/image/chart_renderer):

### `slide_type: "kpi_native"`
Триггеры (любой):
- intent = `data` И в content есть 1-3 крупные числа (не текст), процентные показатели
- key_phrase содержит «KPI», «метрики», «результаты», «N% / N+»
- Числа сопровождаются короткими описаниями (1-2 строки)

Output:
```json
{"slide_type": "kpi_native", "kpi": {
  "title": "...", "numbers": [{"value": "84", "desc": "...", "pct": true, "accent": false}, ...]
}}
```

### `slide_type: "chart_pptx_native"` ⭐ DEFAULT для charts (v1.4+)
Триггеры:
- В draft присутствует chart-like структура (series, axis, time-data, percentage breakdown)
- intent = `chart` или `analytics`
- Контент описывается как **график**, **диаграмма**, **trend**, **breakdown**

**КРИТИЧНО (canonical rule v1.4):** все диаграммы и графики в скилле **обязаны** быть редактируемыми
PowerPoint chart (а не PNG-картинкой). Default → `chart_pptx_native`. Это позволяет пользователю
открыть «Изменить данные» / Edit Data в PowerPoint и менять цифры через встроенный Excel.

Chart types (выбрать по структуре):
- `area_stacked` — temporal data, multiple categories
- `area_100` — temporal data, доли в процентах (sum to 100%)
- `bar` / `bar_stacked` — comparison values (Q1, Q2, ...) или horizontal ranking
- `line` — trend over time, multiple series
- `pie` — single composition / distribution

Output:
```json
{"slide_type": "chart_pptx_native", "chart": {
  "type": "area_100|area_stacked|bar|bar_stacked|line|pie",
  "title": "...", "caption": "...",
  "x": [...], "series": [{"name": "...", "data": [...]}, ...],
  "accent_idx": 0
}}
```

### `slide_type: "chart_native"` (legacy / специальные кейсы)
Использовать ТОЛЬКО когда нужен pixel-perfect canonical стиль с эффектами,
которые PowerPoint native chart не поддерживает: прозрачные overlapping areas,
custom annotations (vertical lines с подписями), особая стилистика осей.
**По умолчанию НЕ использовать** — это PNG-картинка, не редактируется в PPT.

### `slide_type: "image_native"`
Триггеры:
- В draft есть **embedded image** (картинка как файл)
- Картинка = главный контент (>50% площади или sole content)
- intent = `screenshot` / `photo` / `illustration_main`

Output:
```json
{"slide_type": "image_native", "image": {
  "title": "...", "image_path": "path/to/img.png", "caption": "..."
}}
```

**Если нет triggers** для native — использовать стандартный flow `clone_from_slide` (donor catalog).

## Routing decision tree

```
Есть chart-like data (series, axis)? → chart_pptx_native (DEFAULT, editable)
  fallback chart_native только если нужен PNG со спец-эффектами
Есть 1-3 KPI числа без текста-параграфа? → kpi_native
Есть embedded image как main content? → image_native
Иначе → standard category + clone_from_slide donor
```

## Split decision (v1.4) — разгрузка перегруженных слайдов

Если source slide перегружен (превышает canonical density limits §8 из canonical-rules) — **разбить на несколько output slides**.

### Триггеры split

| Источник имеет | Решение |
|---|---|
| 4+ KPI цифр одного типа | split на 2 kpi_native (3+1 или 2+2) |
| 6+ блоков с подзаголовками | split на 2 multicolumn (3+3 или 4+2) |
| body > 80 слов на одну колонку | split на 2 content slides (paragraph+paragraph) |
| 5+ image-thumbnails | split на 2 image_grid (3+2) |
| Чарт с 5+ series | упростить до 2-3 series + split «context» + «detail» chart |
| callout с цитатой 30+ слов | разбить цитату на 2 callout (часть 1 + часть 2) |
| Title 60+ chars + body | split на title-slide + content-slide |
| 3+ topics без structural связи | split на divider + N content slides |

### Decision algorithm

```python
def should_split(slide):
    issues = count_density_issues(slide)
    if issues == 0:
        return [slide]  # OK
    if issues == 1 and severity == "minor":
        return [slide]  # допустимо
    # Перегружен — split
    return split_into_subslides(slide)
```

### Output format при split

```json
{
  "num": 5,
  "_source_slide": 3,
  "_split_part": "1/2",
  "category": "multicolumn",
  ...
},
{
  "num": 6,
  "_source_slide": 3,
  "_split_part": "2/2",
  "category": "multicolumn",
  ...
}
```

### Принципы split

- **Логическая связность**: split по semantic boundaries, не механически
- **Разные donors**: для variety использовать разные idx для одной категории (anti-monotony)
- **Continuity hint**: в title/footer первого/второго слайда указать "1/2", "2/2" или contextual hint
- **Tone balance**: в split не делать обе части тёмными подряд (canonical §2)

### НЕ split когда

- Слайд имеет много элементов **по канонической композиции** (donor 33 = 6 блоков by design)
- Source слайд — title или divider (они и так минимальны)
- Source слайд — chart с 3+ series если все нужны для narrative

## Subjective overload check (designer judgment)

Помимо чек-листа выше, классификатор должен **глазами оценить** source slide и спросить:

1. **«Если бы это был мой слайд, я бы сделал split?»** — если да, делай split
2. **«Глаз пробегает по слайду за 5 секунд?»** — если нет, перегружен
3. **«Take-away мгновенный или приходится вчитываться?»** — если вчитываться, перегружен
4. **«Сколько отдельных идей на слайде?»** — если 3+ независимых, split

### Эвристика «принцип одного слайда»

Canonical §7 пункт 2 (slide 87 шаблона): **«Один слайд = одна мысль»**.

Если source имеет:
- 1 KPI цифру → 1 output slide ✓
- 3 KPI цифры одной темы → 1 output slide (kpi_native 3 цифры) ✓
- 5 KPI цифр **разных тем** → split на 2 (по semantic groups)
- 3 КПЕ разных тем + 4 process steps → 2-3 slides (отдельный KPI слайд + отдельный process слайд)

### Examples of split decisions

| Source content | Split decision |
|---|---|
| Title 200 chars + spикер + 5 KPI + 4 mehanizma | **4 slides:** title (короткий) → divider → KPI → 4 mehanizma |
| 6 product features списком | **2 slides:** features 1-3 + features 4-6 (3+3 multicolumn) |
| Chart 5 series with intro text | **2 slides:** intro text + chart simplified |
| Quote 50 слов | **2 callouts:** часть 1 setup + часть 2 punchline |
| Team of 12 people | **2 slides:** team_5 + team_5+2 (split by role/department) |

## Контекстные правила
- Если 2 соседних слайда выходят в одной категории — отметь это, Layout Designer должен выбрать **разные idx** в той же категории, чтобы был ритм
- При перепаде темы (новый раздел) — рекомендуй `divider` перед

## Пример

**Вход:** 5 слайдов, intent: `[title, data, data, data, image]`

**Выход:**
```json
{
  "slides": [
    {"num": 1, "category": "title", "subcategory_hint": "white", "rationale": "Открывающий слайд, нейтральный фон"},
    {"num": 2, "category": "callout", "subcategory_hint": "white", "rationale": "Одна крупная цифра — Важная информация"},
    {"num": 3, "category": "multicolumn", "subcategory_hint": "3col", "rationale": "3 KPI цифры рядом"},
    {"num": 4, "category": "multicolumn", "subcategory_hint": "3col_alternative", "rationale": "Та же сетка, но другой idx — для ритма"},
    {"num": 5, "category": "image", "subcategory_hint": "illustration_half", "rationale": "Финальный слайд с визуалом"}
  ]
}
```
