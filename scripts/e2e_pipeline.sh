#!/usr/bin/env bash
# e2e_pipeline.sh — запускает детерминированную часть pipeline на готовом plan.json.
#
# LLM-этапы (Brief Reader, Slide Classifier, Layout Designer, Content Distributor,
# Copy Editor) делаются в чате с Claude — результат: plan.json. Этот скрипт берёт
# готовый план и прогоняет всё остальное одной командой.
#
# Этапы:
#   1. validate_plan.py        plan.json → plan_validated.json
#   2. build_v9.py             plan_validated.json + template → result.pptx
#   3. render_slides.py        result.pptx → render/slide-NN.png
#   4. brand_guardian.py       result.pptx → brand_report.json
#   5. visual_validator_v2.py  render/ → visual_report.json (с --plan, чтобы знать про charts)
#
# Usage:
#   bash scripts/e2e_pipeline.sh <plan.json> <output_dir> [template.pptx]
#
# Пример:
#   bash scripts/e2e_pipeline.sh output/humanity/plan.json output/humanity \
#                                ../template/Cloud.ru_Template_2026.pptx
#
# По умолчанию template = ../template/Cloud.ru_Template_2026.pptx (рядом с pptx-skill/).

set -e

PLAN="$1"
OUT="$2"
TEMPLATE="${3:-../template/Cloud.ru_Template_2026.pptx}"

if [ -z "$PLAN" ] || [ -z "$OUT" ]; then
    echo "Usage: bash scripts/e2e_pipeline.sh <plan.json> <output_dir> [template.pptx]" >&2
    echo "" >&2
    echo "  plan.json   — готовый план от LLM (формат build_v9: clone_from_slide или slide_type)" >&2
    echo "  output_dir  — куда складывать validated/result/render/reports" >&2
    echo "  template    — Cloud.ru шаблон (default: ../template/Cloud.ru_Template_2026.pptx)" >&2
    exit 1
fi

# Скрипт должен запускаться из pptx-skill/, чтобы относительные пути brand/ работали.
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PPTX_SKILL_DIR="$( dirname "$SCRIPT_DIR" )"
cd "$PPTX_SKILL_DIR"

if [ ! -f "$PLAN" ]; then
    echo "ERROR: план не найден: $PLAN" >&2
    exit 1
fi

if [ ! -f "$TEMPLATE" ]; then
    echo "ERROR: шаблон не найден: $TEMPLATE" >&2
    exit 1
fi

mkdir -p "$OUT"

PLAN_VALIDATED="$OUT/plan_validated.json"
RESULT="$OUT/result.pptx"
RENDER_DIR="$OUT/render"
BRAND_REPORT="$OUT/brand_report.json"
VISUAL_REPORT="$OUT/visual_report.json"

echo "================================================================"
echo "Cloud.ru Slides Skill — e2e pipeline"
echo "Plan:     $PLAN"
echo "Template: $TEMPLATE"
echo "Output:   $OUT"
echo "================================================================"

echo ""
echo "[1/5] validate_plan…"
python3 scripts/validate_plan.py "$PLAN" brand/donor-slot-map.yaml "$PLAN_VALIDATED" 2>&1 | tail -3 || true

echo ""
echo "[2/5] build_v9…"
python3 scripts/build_v9.py "$PLAN_VALIDATED" "$TEMPLATE" "$RESULT" brand/donor-slot-map.yaml 2>&1 | tail -1

echo ""
echo "[3/5] render_slides…"
rm -rf "$RENDER_DIR"
python3 scripts/render_slides.py "$RESULT" "$RENDER_DIR" 2>&1 | tail -1

echo ""
echo "[4/5] brand_guardian…"
python3 scripts/brand_guardian.py "$RESULT" "$BRAND_REPORT" 2>&1 | grep -E "Verdict|Score" | head -1 || true

echo ""
echo "[5/5] visual_validator_v2 (с --plan для chart-aware палитры)…"
python3 scripts/visual_validator_v2.py "$RENDER_DIR" "$VISUAL_REPORT" --plan "$PLAN_VALIDATED" 2>&1 | grep -E "Verdict" | head -1 || true

echo ""
echo "================================================================"
echo "Done. Артефакты в $OUT/"
echo "  plan_validated.json   — план с auto-add canonical полями"
echo "  result.pptx           — собранная презентация"
echo "  render/slide-NN.png   — PNG-рендеры всех слайдов"
echo "  brand_report.json     — XML-валидация бренда"
echo "  visual_report.json    — pixel-level валидация"
echo ""
echo "Финальный вердикт READY/NEEDS_REWORK выдаёт LLM Visual Verifier"
echo "(агент 10) глазами по PNG-рендерам — выполняется в чате с Claude."
echo "================================================================"
