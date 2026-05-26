#!/usr/bin/env python3
"""
kpi_renderer.py — canonical KPI rendering для build_v9.

Зачем: donor 43/44 в шаблоне = guide-слайды (показывают примеры размеров),
не доноры для произвольных KPI. Их placeholders расположены обучающе
(199pt + 199pt + 44pt) — не для real data.

Решение: рисуем KPI shapes с нуля на blank donor с CLEAN GRID:
- 1 number → hero center
- 2 numbers → 50/50 grid
- 3 numbers → 33/33/33 grid (все одного размера)
- Title 32pt SemiBold top
- Numbers 130pt SemiBold (помещается без overflow)
- Descriptions 16pt Regular under numbers
- Accent: 1 of N number = #26D07C (canonical §12)

Layout (1280×720 slide):
  Title:           (35, 38, 1209, 53) px       32pt SemiBold #222222
  Number row:      y=220, height=280
    1 number:      x=440, w=400 (centered)
    2 numbers:     x=140, w=400 + x=740, w=400
    3 numbers:     x=35, w=400 + x=440, w=400 + x=840, w=400
  Desc row:        y=530, height=100
    Same x columns
  Optional %:      next to number, top-right corner of number box
"""
import os
import sys
import json
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


def _load_template_version():
    """Грузит brand/template-version.json (маппинг slide-индексов). Fallback — None."""
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(here, "..", "brand", "template-version.json"),
        os.path.join(os.getcwd(), "brand", "template-version.json"),
        os.path.join(os.getcwd(), "pptx-skill", "brand", "template-version.json"),
    ]
    for path in candidates:
        try:
            if os.path.isfile(path):
                with open(path, encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            continue
    return None


_TPL_VER = _load_template_version()


GRAPHITE = RGBColor(0x22, 0x22, 0x22)
GREEN = RGBColor(0x26, 0xD0, 0x7C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


# Canonical §4 (slide 43): фрейм 199pt помещает максимум 2 значащих цифры.
KPI_HERO_DIGIT_LIMIT = 2


def _count_significant_digits(value):
    """Считает значащие цифры в KPI value, игнорируя разделители и unit-suffix.

    Примеры:
      '12'    → 2 цифры
      '99%'   → 2 (% игнорируется — pct идёт отдельным флагом)
      '1.5'   → 2 (точка игнорируется, но 1+5=2)
      '1.5K'  → 2 (K — буква-сокращение, не считается)
      '150'   → 3 (overflow!)
      '~150'  → 3 (тильда игнорируется)
      '12,5'  → 3 (русская десятичная — 3 цифры в frame не помещаются)
    """
    if value is None:
        return 0
    s = str(value)
    digits = "".join(c for c in s if c.isdigit())
    return len(digits)


def _add_text_box(slide, left_px, top_px, width_px, height_px, text,
                  font_size_pt=14, font_name="SB Sans Display",
                  bold=False, color=GRAPHITE, align=PP_ALIGN.LEFT,
                  anchor=MSO_ANCHOR.TOP):
    """Add text box with canonical font/size/color."""
    EMU = 9525
    box = slide.shapes.add_textbox(
        Emu(left_px * EMU), Emu(top_px * EMU),
        Emu(width_px * EMU), Emu(height_px * EMU)
    )
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def clean_slide_to_blank(slide):
    """Удалить все shapes на slide. Layout-inherited (logo, footer) останутся.
    Возвращает чистый canvas для KPI rendering."""
    spTree = slide.shapes._spTree
    # Удаляем все sp (shapes) — keep nvGrpSpPr и grpSpPr (структурные)
    to_remove = []
    for child in list(spTree):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ('sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'):
            to_remove.append(child)
    for el in to_remove:
        spTree.remove(el)


def render_kpi(slide, kpi_config, dark=False):
    """
    Render KPI block on slide.

    kpi_config = {
        "title": "Что получилось",
        "numbers": [
            {"value": "12", "desc": "сократили AI-инициатив", "accent": False, "pct": False},
            {"value": "84", "desc": "вовлечённость", "accent": False, "pct": True},
            {"value": "5",  "desc": "инструментов", "accent": True, "pct": False}
        ]
    }

    dark: True if slide is dark (text white instead of graphite)
    """
    text_color = WHITE if dark else GRAPHITE
    n = len(kpi_config.get("numbers", []))

    if n == 0 or n > 3:
        raise ValueError(f"KPI supports 1-3 numbers, got {n}")

    # Title
    title = kpi_config.get("title", "")
    if title:
        _add_text_box(slide, 35, 38, 1209, 53, title,
                      font_size_pt=20, bold=True, color=text_color,
                      anchor=MSO_ANCHOR.MIDDLE)

    # Number boxes layout
    if n == 1:
        x_positions = [440]
        block_width = 400
    elif n == 2:
        x_positions = [140, 740]
        block_width = 400
    else:  # n == 3
        x_positions = [35, 440, 845]
        block_width = 400

    NUMBER_TOP = 200
    NUMBER_HEIGHT = 260
    DESC_TOP = 470
    DESC_HEIGHT = 120
    NUMBER_FONT = 130 if n >= 2 else 199  # single hero number может быть 199pt

    # Reduce font for 3 numbers if they are wide
    if n == 3:
        max_chars = max(len(x["value"]) for x in kpi_config["numbers"])
        if max_chars > 3:
            NUMBER_FONT = 100  # 4+ digits — smaller

    for i, num in enumerate(kpi_config["numbers"]):
        x = x_positions[i]
        is_accent = num.get("accent", False)
        color = GREEN if is_accent else text_color
        value = num["value"]
        has_pct = num.get("pct", False)

        # Canonical §4: 199pt frame ⇒ max 2 значащих цифры. Иначе overflow.
        # Применяем только к hero (199pt single number); для 130/100pt запас больше.
        if NUMBER_FONT >= 199:
            digits = _count_significant_digits(value)
            if digits > KPI_HERO_DIGIT_LIMIT:
                print(
                    f"WARN: KPI value '{value}' содержит {digits} значащих цифр "
                    f"при 199pt frame (canonical max {KPI_HERO_DIGIT_LIMIT}). "
                    f"Используй сокращение ('1.5K', '~{value[:2]}') или сменить layout (donor 47).",
                    file=sys.stderr,
                )

        # Number text box
        _add_text_box(slide, x, NUMBER_TOP, block_width, NUMBER_HEIGHT,
                      value, font_size_pt=NUMBER_FONT, bold=False, color=color,
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Optional % sign — small, top-right corner of number box
        if has_pct:
            pct_size = max(40, NUMBER_FONT // 3)  # % about 1/3 of number
            pct_x = x + block_width - 70
            pct_y = NUMBER_TOP + 20
            _add_text_box(slide, pct_x, pct_y, 60, 80, "%",
                          font_size_pt=pct_size, bold=True, color=color,
                          align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

        # Description below
        desc = num.get("desc", "")
        if desc:
            _add_text_box(slide, x, DESC_TOP, block_width, DESC_HEIGHT,
                          desc, font_size_pt=16, bold=False, color=text_color,
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)


# Constants for blank donor selection.
# Источник: brand/template-version.json → blank_donors (white/dark).
# Hardcoded fallback — на случай отсутствия файла.
_BLANK_DEFAULTS = {"white": 30, "dark": 22}
_blank_cfg = (_TPL_VER or {}).get("blank_donors", {}) if _TPL_VER else {}
BLANK_DONOR_WHITE = int(_blank_cfg.get("white", _BLANK_DEFAULTS["white"]))
BLANK_DONOR_DARK = int(_blank_cfg.get("dark", _BLANK_DEFAULTS["dark"]))


# Standalone test
if __name__ == "__main__":
    from pptx import Presentation

    p = Presentation("template/Cloud.ru_Template_2026.pptx")
    # Use slide 30 as clean blank donor — but FULLY clean shapes
    slide = list(p.slides)[BLANK_DONOR_WHITE - 1]
    clean_slide_to_blank(slide)

    # Render KPI
    render_kpi(slide, {
        "title": "ЧТО ПОЛУЧИЛОСЬ ЧЕРЕЗ 6 МЕСЯЦЕВ",
        "numbers": [
            {"value": "12", "desc": "сократили AI-инициатив", "pct": True},
            {"value": "84", "desc": "вовлечённость команды", "pct": True},
            {"value": "5",  "desc": "рабочих инструментов", "accent": True}
        ]
    }, dark=False)

    out = "pptx-skill/output/kpi_renderer_test.pptx"
    p.save(out)
    print(f"Saved {out}, used slide 30 as blank donor")
