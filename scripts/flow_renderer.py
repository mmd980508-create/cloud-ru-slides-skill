#!/usr/bin/env python3
"""
flow_renderer.py — native PowerPoint flow-diagrams / schemas / process maps.

Зачем: до v1.6 скилл умел charts (chart_pptx_native) и KPI (kpi_native),
но не умел собирать схемы блок-стрелка-блок (SGR, pipeline, process map).
Альтернатива была — рисовать в Figma и вставлять PNG, что не редактируется.

Решение: модуль рисует схему из примитивов PowerPoint shapes на blank canvas.
Все элементы редактируемые: блок можно подвинуть, текст переписать, стрелку
перенаправить — обычными средствами PowerPoint.

Используется build_v9 при slide_type == "flow_diagram_native".

Стиль (canonical, согласован с эталонными слайдами Cloud.ru;
правило 2026-05-06 — feedback_flow_diagram_composition.md):
  - Блоки: серые #F2F2F2, текст SB Sans Display #222222.
  - Текст в блоках: align=LEFT, vertical_anchor=TOP. Поля: 12px со всех сторон,
    низ 16px.
  - Стрелки: единая толщина 1pt, цвет #434343 (тёмно-серый, не #222222 графит).
    Наконечник: открытая галочка (type='arrow'), размер 8 в PPT UI (w=lg,
    len=med). Только горизонтальные/вертикальные — диагонали запрещены.
  - Пунктирные рамки для группировки фаз: 1pt #888888 dash.
  - Заголовок: 20pt SemiBold CAPS, top-left (35, 60).
  - Decor (по бренду): зелёные L-уголки с диагональю — bottom corner.
  - Safe-area: SAFE_TOP=140, SAFE_BOTTOM=660, SAFE_LEFT=30, SAFE_RIGHT=1250.
    Все координаты блоков должны лежать внутри safe-area.

Координаты — в пикселях (slide 1280×720). EMU = px × 9525.

Config schema (передаётся через plan.json):
  {
    "slide_type": "flow_diagram_native",
    "dark": false,
    "flow": {
      "header": "Заголовок схемы",            # обязательное
      "subtitle": "...",                       # опц., 11pt под header'ом
      "subtitle_url": "https://...",           # опц., 9pt серая ссылка
      "blocks": [
        {
          "id": "b1",                          # опц., для arrows by ref
          "x": 175, "y": 180, "w": 235, "h": 50,
          "lines": ["Title", "subtitle text"],
          "font_sizes": [13, 11],              # опц.
          "bolds": [true, false],              # опц.
          "caps_first": false,                 # опц.
          "fill": "gray|green|dark|white",     # опц., default "gray"
          "align": "left|center|right",        # опц., default "left" (canonical)
          "vanchor": "top|middle"              # опц., default "top" (canonical)
        }
      ],
      "groups": [                              # опц., dashed рамки + лейблы
        {
          "label": "Phase 1",
          "x": 167, "y": 154, "w": 251, "h": 86,
          "label_pos": "top|bottom"            # default "top"
        }
      ],
      "arrows": [
        # вариант 1 — явные координаты
        {"x1": 410, "y1": 205, "x2": 432, "y2": 205,
         "with_head": true, "dashed": false},
        # вариант 2 — между блоками по id
        {"from": "b1", "to": "b2", "side": "right"}
      ],
      "labels": [                              # опц., свободные подписи
        {"x": 35, "y": 122, "w": 600, "h": 22,
         "text": "SGR · Schema-Guided Reasoning",
         "font_size": 11, "bold": false,
         "align": "left", "caps": false}
      ],
      "decor": {                               # опц., зелёные L-уголки
        "enabled": true,
        "count": 4,
        "x_start": 950, "y_start": 625,
        "size": 38, "gap": 12
      }
    }
  }
"""
import os
import json

from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn
from lxml import etree


EMU = 9525
SLIDE_W_PX = 1280
SLIDE_H_PX = 720


# ============================================================================
# Палитра (из brand/palette.json — единый источник истины)
# ============================================================================
def _load_palette():
    here = os.path.dirname(os.path.abspath(__file__))
    for path in (
        os.path.join(here, "..", "brand", "palette.json"),
        os.path.join(os.getcwd(), "brand", "palette.json"),
        os.path.join(os.getcwd(), "pptx-skill", "brand", "palette.json"),
    ):
        try:
            if os.path.isfile(path):
                with open(path, encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            continue
    return None


_PAL = _load_palette()


def _hex(hx):
    h = hx.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _from_palette(section, name, fallback):
    if _PAL and name in _PAL.get(section, {}):
        return _hex(_PAL[section][name])
    return _hex(fallback)


GRAPHITE = _from_palette("base", "Black", "#222222")
WHITE = _from_palette("base", "White", "#FFFFFF")
GRAY = _from_palette("base", "Gray", "#F2F2F2")
GREEN = _from_palette("base", "Green", "#26D07C")
GRAPHITE_DARK = _from_palette("base_alts", "Graphite-Iron", "#343F48")
DASH_GRAY = RGBColor(0x88, 0x88, 0x88)
SEPARATOR_GRAY = RGBColor(0xCC, 0xCC, 0xCC)

# Canonical для flow-схем (правило 2026-05-06):
ARROW_COLOR = RGBColor(0x43, 0x43, 0x43)   # единый цвет всех стрелок
ARROW_WIDTH_PT = 1.0                        # единая толщина всех стрелок


FONT = "SB Sans Display"


# Safe-area для 1280×720 Cloud.ru slide.
# Все координаты flow-схем должны рассчитываться относительно этих констант.
SAFE_TOP = 140       # под заголовком 20pt + разделителем
SAFE_BOTTOM = 660    # выше footer copyright
SAFE_LEFT = 30
SAFE_RIGHT = 1250
SAFE_W = SAFE_RIGHT - SAFE_LEFT   # 1220
SAFE_H = SAFE_BOTTOM - SAFE_TOP    # 520


# ============================================================================
# Низкоуровневые примитивы (можно импортировать отдельно для кастомных схем)
# ============================================================================
def px(v):
    """px → EMU."""
    return Emu(int(v) * EMU)


def _no_effects(shape):
    """Удалить drop-shadow / glow эффекты из shape (брендбук: no effects)."""
    spPr = shape._element.spPr
    for tag in ("effectLst", "effectDag"):
        for e in spPr.findall(qn(f"a:{tag}")):
            spPr.remove(e)


def _resolve_fill(fill_name):
    """Имя fill → (rgb_color, text_color)."""
    if fill_name in (None, "gray"):
        return GRAY, GRAPHITE
    if fill_name == "white":
        return WHITE, GRAPHITE
    if fill_name == "green":
        return GREEN, WHITE
    if fill_name == "dark":
        return GRAPHITE_DARK, WHITE
    # raw hex
    if isinstance(fill_name, str) and fill_name.startswith("#"):
        return _hex(fill_name), GRAPHITE
    return GRAY, GRAPHITE


def add_block(slide, x, y, w, h, lines,
              font_sizes=None, bolds=None,
              caps_first=False, fill="gray", align="left",
              vanchor="top",
              text_color=None):
    """Прямоугольник с многострочным контентом.

    Canonical (правило 2026-05-06):
    - Текст ВСЕГДА выравнивается по ЛЕВОМУ (align='left') и ВЕРХНЕМУ (vanchor='top') краю.
    - Поля одинаковые 12px, нижнее 16px.

    lines: list[str].
    font_sizes / bolds: list of same len. По умолчанию [13, 11, 11, ...] / [True, False, ...].
    caps_first: первая строка → CAPS.
    fill: "gray" (default) / "white" / "green" / "dark" / hex string.
    align: "left" (default, canonical) / "center" / "right".
    vanchor: "top" (default, canonical) / "middle".
    text_color: явный override (RGBColor); иначе из fill.
    """
    rgb_fill, default_text = _resolve_fill(fill)
    text = text_color if text_color is not None else default_text

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y), px(w), px(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb_fill
    shape.line.fill.background()
    _no_effects(shape)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP if vanchor == "top" else MSO_ANCHOR.MIDDLE
    tf.margin_left = Emu(12 * EMU)
    tf.margin_right = Emu(12 * EMU)
    tf.margin_top = Emu(12 * EMU)
    tf.margin_bottom = Emu(16 * EMU)

    if font_sizes is None:
        font_sizes = [13] + [11] * max(0, len(lines) - 1)
    if bolds is None:
        bolds = [True] + [False] * max(0, len(lines) - 1)

    align_enum = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
    }.get(align, PP_ALIGN.LEFT)

    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align_enum
        run = p.add_run()
        run.text = line.upper() if (caps_first and i == 0) else line
        run.font.name = FONT
        run.font.size = Pt(font_sizes[i] if i < len(font_sizes) else font_sizes[-1])
        run.font.bold = bolds[i] if i < len(bolds) else bolds[-1]
        run.font.color.rgb = text
    return shape


def add_label(slide, x, y, w, h, text,
              font_size=11, bold=False,
              align="left", anchor="top", caps=False,
              color=None):
    """Свободная текстовая подпись (без фона)."""
    box = slide.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE if anchor == "middle" else MSO_ANCHOR.TOP
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)

    align_enum = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
    }.get(align, PP_ALIGN.LEFT)

    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align_enum
        run = p.add_run()
        run.text = line.upper() if caps else line
        run.font.name = FONT
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = color if color is not None else GRAPHITE
    return box


def add_arrow(slide, x1, y1, x2, y2,
              with_head=True, w_pt=None, dashed=False, color=None):
    """Прямая стрелка/коннектор.

    Canonical (правило 2026-05-06):
    - Только горизонтальные или вертикальные. Диагонали запрещены (валидация ниже).
    - Толщина: единая ARROW_WIDTH_PT = 1.0 (w_pt override только для исключений).
    - Цвет: единый ARROW_COLOR = #434343.
    - Голова: открытая галочка (type='arrow') размер 8 (w=lg, len=med).
    """
    if x1 != x2 and y1 != y2:
        raise ValueError(
            f"add_arrow: диагональная стрелка ({x1},{y1})→({x2},{y2}) запрещена. "
            "Используй ломаную через 90° (несколько add_arrow с промежуточными точками)."
        )
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, px(x1), px(y1), px(x2), px(y2)
    )
    conn.line.color.rgb = color if color is not None else ARROW_COLOR
    effective_w = w_pt if w_pt is not None else ARROW_WIDTH_PT
    conn.line.width = Emu(int(effective_w * 12700))
    if dashed:
        ln = conn.line._get_or_add_ln()
        prstDash = etree.SubElement(ln, qn("a:prstDash"))
        prstDash.set("val", "dash")
    if with_head:
        ln = conn.line._get_or_add_ln()
        for tag in ("tailEnd", "headEnd"):
            for el in ln.findall(qn(f"a:{tag}")):
                ln.remove(el)
        tail_end = etree.SubElement(ln, qn("a:tailEnd"))
        # PowerPoint UI «открытая стрелка, размер 8» = (type=arrow, w=lg, len=med).
        # (w=lg, len=lg) = размер 9 (максимум).
        tail_end.set("type", "arrow")
        tail_end.set("w", "lg")
        tail_end.set("len", "med")
    return conn


def add_dashed_rect(slide, x, y, w, h, color=None, w_pt=1.0):
    """Пунктирная рамка (для группировки phase-блоков)."""
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y), px(w), px(h))
    rect.fill.background()
    rect.line.color.rgb = color if color is not None else DASH_GRAY
    rect.line.width = Emu(int(w_pt * 12700))
    ln = rect.line._get_or_add_ln()
    prstDash = etree.SubElement(ln, qn("a:prstDash"))
    prstDash.set("val", "dash")
    _no_effects(rect)
    return rect


def add_header(slide, text, dark=False):
    """Заголовок слайда — 20pt SemiBold CAPS, top-left (35, 60).
    Canonical стиль content-слайда шаблона Cloud.ru."""
    box = slide.shapes.add_textbox(px(35), px(60), px(1209), px(40))
    tf = box.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text.upper()
    run.font.name = FONT
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = WHITE if dark else GRAPHITE


def add_top_separator(slide, y=110, color=None):
    """Тонкая серая линия под заголовком (визуальный разделитель)."""
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, px(35), px(y), px(1245), px(y)
    )
    conn.line.color.rgb = color if color is not None else SEPARATOR_GRAY
    conn.line.width = Emu(int(0.5 * 12700))


def add_decor_diagonals(slide,
                        count=4, x_start=20, y_start=620,
                        size=70, gap=14, w_pt=1.4, color=None):
    """Зелёные L-уголки с диагональю — фирменный декор Cloud.ru.

    Ставится в bottom-left или bottom-right (свободный угол).
    """
    c = color if color is not None else GREEN
    for i in range(count):
        x = x_start + i * (size + gap)
        y = y_start
        h_line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, px(x), px(y), px(x + size), px(y)
        )
        v_line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, px(x + size), px(y), px(x + size), px(y + size)
        )
        d_line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, px(x), px(y), px(x + size), px(y + size)
        )
        for ln in (h_line, v_line, d_line):
            ln.line.color.rgb = c
            ln.line.width = Emu(int(w_pt * 12700))


# ============================================================================
# Высокоуровневая функция — сборка схемы по конфигу
# ============================================================================
def _block_anchor(block, side):
    """Точка на границе блока по стороне (right|left|top|bottom|center)."""
    x, y, w, h = block["x"], block["y"], block["w"], block["h"]
    if side == "right":
        return x + w, y + h // 2
    if side == "left":
        return x, y + h // 2
    if side == "top":
        return x + w // 2, y
    if side == "bottom":
        return x + w // 2, y + h
    # center
    return x + w // 2, y + h // 2


def render_flow_diagram_slide(slide, flow_config, dark=False):
    """Собирает flow-схему на готовом (очищенном) blank slide.

    Предполагается, что slide уже прошёл clean_slide_to_blank() — этим
    занимается build_v9 до вызова render_flow_diagram_slide.
    """
    if not isinstance(flow_config, dict):
        raise ValueError("flow_config должен быть dict")

    # 1. Header
    header = flow_config.get("header", "")
    if header:
        add_header(slide, header, dark=dark)
        add_top_separator(slide)

    # 1a. Subtitle / URL
    subtitle = flow_config.get("subtitle")
    if subtitle:
        add_label(slide, 35, 122, 800, 22, subtitle,
                  font_size=11, bold=False, align="left",
                  color=WHITE if dark else GRAPHITE)
    subtitle_url = flow_config.get("subtitle_url")
    if subtitle_url:
        add_label(slide, 35, 142, 800, 18, subtitle_url,
                  font_size=9, bold=False, align="left",
                  color=DASH_GRAY)

    # 2. Blocks — собираем по id для arrow refs
    blocks_by_id = {}
    for blk in flow_config.get("blocks", []):
        # Canonical default: align=left, vanchor=top (правило 2026-05-06).
        shape = add_block(
            slide,
            blk["x"], blk["y"], blk["w"], blk["h"],
            blk.get("lines", []),
            font_sizes=blk.get("font_sizes"),
            bolds=blk.get("bolds"),
            caps_first=blk.get("caps_first", False),
            fill=blk.get("fill", "gray"),
            align=blk.get("align", "left"),
            vanchor=blk.get("vanchor", "top"),
        )
        if "id" in blk:
            blocks_by_id[blk["id"]] = blk

    # 3. Groups — dashed рамка + label
    for grp in flow_config.get("groups", []):
        gx, gy, gw, gh = grp["x"], grp["y"], grp["w"], grp["h"]
        add_dashed_rect(slide, gx, gy, gw, gh)
        label_text = grp.get("label")
        if label_text:
            label_pos = grp.get("label_pos", "top")
            label_y = gy + 2 if label_pos == "top" else gy + gh - 18
            add_label(slide, gx, label_y, gw, 16, label_text,
                      font_size=10, bold=True, align="center")

    # 4. Arrows
    for arr in flow_config.get("arrows", []):
        if "from" in arr and "to" in arr:
            src = blocks_by_id.get(arr["from"])
            dst = blocks_by_id.get(arr["to"])
            if src is None or dst is None:
                continue
            side = arr.get("side", "right")
            x1, y1 = _block_anchor(src, side)
            # По умолчанию входим в противоположную сторону
            opposite = {"right": "left", "left": "right",
                        "top": "bottom", "bottom": "top"}.get(side, "left")
            x2, y2 = _block_anchor(dst, arr.get("to_side", opposite))
        else:
            x1, y1 = arr["x1"], arr["y1"]
            x2, y2 = arr["x2"], arr["y2"]
        # w_pt=None → используем ARROW_WIDTH_PT (canonical единая толщина).
        add_arrow(
            slide, x1, y1, x2, y2,
            with_head=arr.get("with_head", True),
            w_pt=arr.get("w_pt"),
            dashed=arr.get("dashed", False),
        )

    # 5. Labels (произвольные)
    for lab in flow_config.get("labels", []):
        add_label(
            slide,
            lab["x"], lab["y"], lab["w"], lab["h"],
            lab["text"],
            font_size=lab.get("font_size", 11),
            bold=lab.get("bold", False),
            align=lab.get("align", "left"),
            caps=lab.get("caps", False),
        )

    # 6. Decor (зелёные уголки)
    decor = flow_config.get("decor")
    if decor and decor.get("enabled"):
        add_decor_diagonals(
            slide,
            count=decor.get("count", 4),
            x_start=decor.get("x_start", 20),
            y_start=decor.get("y_start", 620),
            size=decor.get("size", 70),
            gap=decor.get("gap", 14),
            w_pt=decor.get("w_pt", 1.4),
        )


# ============================================================================
# CLI для standalone-теста
# ============================================================================
def _cli_main():
    """python3 flow_renderer.py <flow_config.json> <template.pptx> <out.pptx>"""
    import sys
    from pptx import Presentation

    if len(sys.argv) < 4:
        print("Usage: flow_renderer.py <flow_config.json> <template.pptx> <out.pptx>",
              file=sys.stderr)
        sys.exit(1)

    cfg_path, tpl_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
    with open(cfg_path, encoding="utf-8") as f:
        cfg = json.load(f)

    # Импорт через relative path (для standalone)
    here = os.path.dirname(os.path.abspath(__file__))
    import sys as _sys
    _sys.path.insert(0, here)
    from kpi_renderer import clean_slide_to_blank, BLANK_DONOR_WHITE, BLANK_DONOR_DARK
    from build_v9 import clone_slide

    dark = cfg.get("dark", False)
    flow = cfg.get("flow", cfg)  # допустим, если передан "flat" config

    p = Presentation(tpl_path)
    originals = list(p.slides)
    blank_idx = BLANK_DONOR_DARK if dark else BLANK_DONOR_WHITE
    new_slide = clone_slide(p, originals[blank_idx - 1])

    # удалить оригиналы
    sldIdLst = p.slides._sldIdLst
    for sid in list(sldIdLst)[:len(originals)]:
        rid = sid.attrib[qn("r:id")]
        try:
            p.part.drop_rel(rid)
        except Exception:
            pass
        sldIdLst.remove(sid)

    clean_slide_to_blank(new_slide)
    render_flow_diagram_slide(new_slide, flow, dark=dark)
    p.save(out_path)
    print(f"Saved {out_path}", file=sys.stderr)


if __name__ == "__main__":
    _cli_main()
