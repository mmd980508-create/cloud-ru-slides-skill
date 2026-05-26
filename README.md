# Cloud.ru Slides Skill

Claude.app skill для автоматической вёрстки `.pptx` по бренду Cloud.ru 2.0.

**Версия:** v1.7 (2026-05-26)

**Что нового в v1.7:**
- **Flow-схемы canonical правила:** стрелки 1pt / #434343 / открытая галочка size 8 (`w=lg, len=med` в OOXML); текст по левому+верхнему краю; safe-area константы.
- **Правило растягивания** — многорядная схема (>1 ряд И >3 колонок) → растянуть на всю safe-area; линейный pipeline → естественная высота + центрирование.
- **Чеклист после вёрстки** — встроен в canonical-rules §3 (текст влезает → формируем; нет → расширяем фреймы → уменьшаем шрифт ≥10pt → пересматриваем композицию).
- **Валидация диагоналей** — `add_arrow` в `scripts/flow_renderer.py` теперь кидает ValueError при попытке нарисовать диагональную стрелку.

**Главные фичи v1.5:**
- **8 расширенных brand-validators** — зелёные стрелки, цветной текст, доп. цвета без зелёного, лимит цифр в KPI 199pt, false-positive WARN на charts убран
- **Единый файл палитры** ([brand/palette.json](brand/palette.json)) — один источник истины для всех валидаторов
- **Версионирование шаблона** ([brand/template-version.json](brand/template-version.json)) — slide-индексы вынесены из кода
- **Cross-platform** — Windows / Linux / macOS пути для LibreOffice
- **Regression suite** ([tests/regression.py](tests/regression.py)) — автоматическая проверка от поломок
- **Editable PowerPoint charts** (унаследовано из v1.4) — диаграммы редактируются прямо в PowerPoint через «Изменить данные»
- **Native rendering** (KPI / image / chart_pptx_native / chart_native)
- **Honest 5-dim verdicts** — НЕ объявлять PASS без визуальной проверки каждого слайда

## Что внутри

```
pptx-skill/
├── SKILL.md                          # главный файл скилла (frontmatter "Use this skill when…")
├── README.md                         # этот файл
├── LEARNINGS.md                      # журнал ошибок и фиксов
├── agents/                           # 9 инструкций для AI-агентов
│   ├── 01-brief-reader.md
│   ├── 02-slide-classifier.md
│   ├── 03-content-distributor.md
│   ├── 04-layout-designer.md
│   ├── 05-icon-picker.md
│   ├── 06-infographic-maker.md
│   ├── 07-copy-editor.md
│   ├── 08-brand-guardian.md
│   └── 09-verifier.md
├── brand/
│   ├── brand-rules.md                # компактные правила бренда
│   ├── template-analysis.md          # семантический каталог 102 layouts
│   └── template-layouts-dump.json    # машинный дамп layouts (5400+ строк)
├── scripts/
│   ├── parse_pptx.py                 # сырой .pptx → JSON структуры
│   ├── build_pptx.py                 # JSON-план + Cloud.ru шаблон → новый .pptx
│   └── kill_widows.py                # русская типографика (nbsp, тире, кавычки)
├── dictionaries/
│   ├── short-words-ru.txt            # короткие слова для неразрывных пробелов
│   └── whitelist.txt                 # бренды/имена, которые не трогаем
├── input/                            # сюда драфты от пользователя (пустая)
├── output/                           # сюда результат (smoke_result.pptx уже здесь)
└── references/                       # доп. ассеты (иконки, рефы) — пустая
```

## Установка в Claude.app

1. Заархивируй папку `pptx-skill/` в `cloud-ru-slides.zip`:
   ```bash
   cd /Users/gmmelnikov/Desktop/Презентации\ в\ ИИ
   zip -r cloud-ru-slides.zip pptx-skill/ -x "*.DS_Store"
   ```
2. В Claude.app (web или desktop) открой Settings → Skills (или Capabilities)
3. Загрузи `cloud-ru-slides.zip`
4. Скилл активируется автоматически по triggers (см. SKILL.md frontmatter)

**ВАЖНО:** в чате при работе со скиллом всегда подгружай файл `Cloud.ru_Template_2026.pptx` (29MB) — он не входит в скилл из-за лимитов размера.

## Использование

### Сценарий 1: Сырой драфт → вёрстка
```
[пользователь подгружает draft.pptx и Cloud.ru_Template_2026.pptx]
> Сверстай эту презентацию по бренду Cloud.ru
```

### Сценарий 2: Бриф → новая презентация
```
[пользователь даёт markdown-бриф или текст]
> Создай презентацию по этому брифу. Используй Cloud.ru шаблон, я приложу его в чат.
```

### Сценарий 3: Аудит
```
[пользователь подгружает готовую презентацию]
> Проверь эту презентацию на соответствие брендбуку Cloud.ru
```

## Требования

- Claude.app (web или desktop) с подпиской, дающей доступ к Skills
- Доступ к code execution (для запуска Python-скриптов)
- `python-pptx` (предустановлен в sandbox)

## Ограничения текущей версии (v0.1)

- Smoke-test прошёл на 1-слайдовой `slide_graph_humanity.pptx` → 3-слайдовый результат
- Не реализован полный LLM-оркестратор (план генерируется агентами по очереди — детали в SKILL.md)
- Не реализована поддержка инфографики через shapes (TODO v0.2)
- Не реализован копи-текст изображений из исходника (TODO v0.2)
- Есть warning duplicate XML в результирующем .pptx — безвредно, фикс в v0.2 (LRN-002)

## Контакты / поддержка

См. `LEARNINGS.md` для накопленного опыта итераций. Vault folder note для проекта: `01 Projects/Cloud.ru Slides Skill/Cloud.ru Slides Skill.md`.

## Версия

v0.1 — 2026-05-01
