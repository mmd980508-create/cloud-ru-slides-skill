#!/usr/bin/env python3
"""
validate_plan.py — gate перед build_v6.

Принимает plan.json + donor-slot-map.yaml. Делает:
1. Проверка что каждый donor существует в slot-map
2. Проверка что все slots в плане определены в donor schema
3. Авто-добавление canonical_color/canonical_size_pt/canonical_bold в slot_styles_override
   (даже если автор плана забыл — Layout Designer не должен ошибиться)
4. Флагает overflow: длина текста > safe_max_chars → warning, > max_chars → error
5. Сохраняет validated_plan.json с enriched overrides

Возвращает exit code:
    0 — OK (план готов к build)
    1 — WARN (есть overflow или потенциальные проблемы, но build возможен)
    2 — FAIL (есть критические ошибки, build остановлен)

Usage:
    python3 validate_plan.py <input_plan.json> <slot_map.yaml> [<output_validated_plan.json>]
"""
import sys
import json
import yaml


def load_inputs(plan_path, slot_map_path):
    with open(plan_path, encoding="utf-8") as f:
        plan = json.load(f)
    with open(slot_map_path, encoding="utf-8") as f:
        smap = yaml.safe_load(f)
    return plan, smap


def validate_slide(slide_idx, slide, donors):
    """Returns (enriched_slide, errors, warnings)."""
    errors, warnings = [], []

    # v0.20+: native slide types не используют clone_from_slide
    # (kpi_native, image_native, chart_native, chart_pptx_native — рендерятся через build_v9 с blank donor)
    slide_type = slide.get("slide_type")
    if slide_type in ("kpi_native", "image_native", "chart_native", "chart_pptx_native"):
        # Проверяем наличие соответствующих data блоков
        data_key = {
            "kpi_native": "kpi",
            "image_native": "image",
            "chart_native": "chart",
            "chart_pptx_native": "chart",
        }[slide_type]
        if data_key not in slide:
            errors.append(f"slide[{slide_idx}]: slide_type={slide_type} требует поля '{data_key}'")
        return slide, errors, warnings

    donor_id = slide.get("clone_from_slide")

    if donor_id not in donors:
        errors.append(f"slide[{slide_idx}]: donor {donor_id} НЕ найден в donor-slot-map.yaml")
        return slide, errors, warnings

    donor = donors[donor_id]
    donor_slots = donor.get("slots", {})

    # plan может иметь slot_styles_override (опционально)
    overrides = dict(slide.get("slot_styles_override", {}))

    plan_slots = slide.get("slots", {})

    for slot_name, content in plan_slots.items():
        if slot_name not in donor_slots:
            warnings.append(
                f"slide[{slide_idx}].{slot_name}: slot НЕ определён в donor {donor_id} schema"
            )
            continue

        donor_slot = donor_slots[slot_name]

        # Проверка длины
        text = content if isinstance(content, str) else (content.get("text", "") if isinstance(content, dict) else "")
        text_len = len(str(text))
        max_chars = donor_slot.get("max_chars")
        safe_max = donor_slot.get("safe_max_chars")

        if max_chars and text_len > max_chars:
            # Auto-apply STRATEGY 3 (brand-rules.md §14b): уменьшить размер на 20%
            # для overflow до 30% над max, чтобы текст влез
            overflow_ratio = text_len / max_chars
            if overflow_ratio <= 1.3:
                base_size = donor_slot.get("size_pt", 60)
                # Уменьшаем на ~20-25% (но не ниже 10pt)
                reduction_pct = min(25, max(15, int((overflow_ratio - 1.0) * 100 * 1.5)))
                new_size = max(10, round(base_size * (1 - reduction_pct / 100)))
                slot_override = dict(overrides.get(slot_name, {}))
                if "size_pt" not in slot_override:
                    slot_override["size_pt"] = new_size
                    overrides[slot_name] = slot_override
                warnings.append(
                    f"slide[{slide_idx}].{slot_name}: AUTO STRATEGY 3 — overflow {text_len}>{max_chars} "
                    f"(ratio {overflow_ratio:.2f}x), уменьшен размер {base_size}pt → {new_size}pt (-{reduction_pct}%)"
                )
            else:
                errors.append(
                    f"slide[{slide_idx}].{slot_name}: текст {text_len} chars > max {max_chars} (overflow {overflow_ratio:.2f}x). "
                    f">30% — нужен другой donor или split. STRATEGY 3 не применима"
                )
        elif safe_max and text_len > safe_max:
            warnings.append(
                f"slide[{slide_idx}].{slot_name}: текст {text_len} chars > safe {safe_max} (visual risk)"
            )

        # Авто-добавление canonical overrides
        slot_override = dict(overrides.get(slot_name, {}))
        added_fields = []

        if "canonical_color" in donor_slot and "color" not in slot_override:
            slot_override["color"] = donor_slot["canonical_color"]
            added_fields.append(f"color={donor_slot['canonical_color']}")
        if "canonical_size_pt" in donor_slot and "size_pt" not in slot_override:
            slot_override["size_pt"] = donor_slot["canonical_size_pt"]
            added_fields.append(f"size_pt={donor_slot['canonical_size_pt']}")
        if "canonical_bold" in donor_slot and "bold" not in slot_override:
            slot_override["bold"] = donor_slot["canonical_bold"]
            added_fields.append(f"bold={donor_slot['canonical_bold']}")

        if slot_override and slot_name not in overrides:
            overrides[slot_name] = slot_override
        elif slot_override and slot_name in overrides:
            overrides[slot_name].update(slot_override)

        if added_fields:
            warnings.append(
                f"slide[{slide_idx}].{slot_name}: auto-added canonical " + ", ".join(added_fields)
            )

    enriched = dict(slide)
    if overrides:
        enriched["slot_styles_override"] = overrides

    # Проверка canonical_rule на уровне donor (если описан)
    if "canonical_rule" in donor:
        warnings.append(
            f"slide[{slide_idx}] (donor {donor_id}): применено canonical правило — '{donor['canonical_rule'][:80]}'"
        )

    return enriched, errors, warnings


def main():
    if len(sys.argv) < 3:
        print("Usage: validate_plan.py <plan.json> <slot_map.yaml> [<output.json>]", file=sys.stderr)
        sys.exit(2)

    plan_path = sys.argv[1]
    smap_path = sys.argv[2]
    out_path = sys.argv[3] if len(sys.argv) > 3 else None

    plan, smap = load_inputs(plan_path, smap_path)
    donors = smap.get("donors", {})

    enriched_slides = []
    all_errors, all_warnings = [], []

    for i, slide in enumerate(plan.get("slides", []), 1):
        e_slide, errs, warns = validate_slide(i, slide, donors)
        enriched_slides.append(e_slide)
        all_errors.extend(errs)
        all_warnings.extend(warns)

    enriched_plan = dict(plan)
    enriched_plan["slides"] = enriched_slides
    enriched_plan["_validated_by"] = "validate_plan.py v0.9"
    enriched_plan["_validation"] = {
        "errors": all_errors,
        "warnings": all_warnings,
        "verdict": "FAIL" if all_errors else ("WARN" if all_warnings else "OK"),
    }

    print("=" * 60, file=sys.stderr)
    print(f"Plan validation — {plan_path}", file=sys.stderr)
    print("=" * 60, file=sys.stderr)
    print(f"Verdict: {enriched_plan['_validation']['verdict']}", file=sys.stderr)
    print(f"Slides: {len(enriched_slides)}, Errors: {len(all_errors)}, Warnings: {len(all_warnings)}", file=sys.stderr)
    print(file=sys.stderr)

    for err in all_errors:
        print(f"  ❌ {err}", file=sys.stderr)
    for w in all_warnings[:30]:  # limit verbose output
        print(f"  ⚠️  {w}", file=sys.stderr)
    if len(all_warnings) > 30:
        print(f"  ... +{len(all_warnings) - 30} more warnings", file=sys.stderr)

    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(enriched_plan, f, ensure_ascii=False, indent=2)
        print(f"\nEnriched plan → {out_path}", file=sys.stderr)

    sys.exit(2 if all_errors else (1 if all_warnings else 0))


if __name__ == "__main__":
    main()
