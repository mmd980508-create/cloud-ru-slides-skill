#!/usr/bin/env python3
"""
layout_designer.py — автоматический выбор donor для каждого slide-content.

Реализует алгоритм из agents/04-layout-designer.md v0.9:
  1. Для каждого классифицированного слайда выбрать donor по category
  2. Применить overflow strategy если контент не помещается
  3. Anti-monotony: ≤2 одинаковых category_equivalence подряд
  4. Variety enforcement: ≥60% уникальных donor в плане ≥5 слайдов

Вход: classified.json (Slide Classifier output) + donor-slot-map.yaml

Формат classified.json:
{
  "slides": [
    {
      "num": 1,
      "category": "title_open|divider|content_3col|kpi|callout|...",
      "title": "...",
      "subtitle": "...",
      "body": ["...", "..."],
      "subtitles": ["...", "...", "..."],   # для multicolumn
      "numbers": ["20", "437", "16"],        # для kpi
      "descriptions": ["...", "..."],
      "tone_hint": "dark|light|green",
      "image_path": "..."                    # если slide со скриншотом/фото
    }
  ]
}

Выход: plan.json готовый для validate_plan.py + build_v7.py

Usage:
    python3 layout_designer.py <classified.json> <slot_map.yaml> <output_plan.json>
"""
import sys
import os
import json
import yaml
from collections import Counter


def load_inputs(classified_path, slot_map_path):
    with open(classified_path, encoding="utf-8") as f:
        classified = json.load(f)
    with open(slot_map_path, encoding="utf-8") as f:
        smap = yaml.safe_load(f)
    return classified, smap


def fits_donor(content, donor_slot_def):
    """True если text-len контента ≤ safe_max_chars донора."""
    if not donor_slot_def:
        return False
    safe = donor_slot_def.get("safe_max_chars") or donor_slot_def.get("max_chars", 1000)
    return len(content or "") <= safe


def pick_donor_from_group(group_donors, donors_yaml, content_check_fn, recently_used, total_used):
    """Выбрать donor из группы:
      1. Кто помещает контент (content_check_fn)
      2. Кто НЕ был использован в последних 2 слайдах (anti-monotony)
      3. Round-robin по total_used: предпочесть наименее использованных (variety boost)
    """
    candidates = []
    for d in group_donors:
        donor = donors_yaml.get(d)
        if not donor:
            continue
        if not content_check_fn(donor):
            continue
        candidates.append(d)

    if not candidates:
        candidates = list(group_donors)

    # Anti-monotony: НЕ в последних 2
    fresh = [d for d in candidates if d not in recently_used[-2:]]
    pool = fresh if fresh else candidates

    # Round-robin: сортируем по числу использований (меньше — лучше)
    pool_sorted = sorted(pool, key=lambda d: (total_used.get(d, 0), pool.index(d)))
    return pool_sorted[0]


def design_slide(slide_content, donors_yaml, equiv_groups, recently_used, total_used):
    """Возвращает {"clone_from_slide": N, "slots": {...}} для plan."""
    cat = slide_content.get("category", "content_text")
    tone = slide_content.get("tone_hint", "light")

    # === Маршрутизация category → equivalence_group ===
    cat_to_group = {
        "title": "title_open" if tone != "dark" else "title_dark",
        "title_open": "title_open",
        "title_dark": "title_dark",
        "divider": "divider",
        "content_2col": "content_2col",
        "content_3col": "content_3col",
        "content_4block": "content_4block",
        "content_4subtitles": "content_4block",
        "content_text": "content_text",
        "callout": "callout",
        "kpi": "kpi",
        "kpi_3numbers": "kpi",
        "diagram": "diagram",
        "table": "table",
        "image_grid": "image_grid",
        "screenshot": "screenshot",
        "team": "team",
        "timeline": "timeline",
        "logo": "logo_finale",
        "logo_finale": "logo_finale",
    }

    group_name = cat_to_group.get(cat, "content_text")

    # === IMAGE ROUTING v0.17: image-as-content vs image-as-accent ===
    # ROOT CAUSE FIX: image на title-слайде ломает композицию, т.к. title-donors
    # имеют title-зону (4/5/6) с фиксированной геометрией. Большой image (chart/screenshot/diagram)
    # перекрывает title или попадает в маленький box. Решение:
    # 1) Detect image role (content vs accent)
    # 2) Если content — маршрутизировать на image-content donor (screenshot/image_grid)
    # 3) Для title с content-image — флаг _needs_title_split для caller (build separate slides)
    image_path = slide_content.get("image_path")
    image_role = slide_content.get("image_role")  # "content" | "accent" | None
    image_meta = slide_content.get("image_meta", {}) or {}
    image_is_large = (
        image_meta.get("width_px", 0) >= 600
        or image_meta.get("kind") in ("chart", "graph", "screenshot", "diagram")
    )
    if image_path and image_role is None:
        # Default: large image → content, small → accent
        image_role = "content" if image_is_large else "accent"

    if image_path and image_role == "content":
        if cat in ("content_text", "content_2col", "content_3col"):
            body_count = len(slide_content.get("body", []) or slide_content.get("subtitles", []))
            if body_count >= 3:
                group_name = "image_grid"
            else:
                group_name = "screenshot"
        elif cat in ("title_open", "title_dark", "title"):
            # Title с большим content-image — маршрутизируем на image-donor,
            # title text должен быть на ОТДЕЛЬНОМ слайде (split signal для caller).
            group_name = "screenshot"
            slide_content["_needs_title_split"] = True
        # для divider/callout/kpi/logo — image-content игнорируется (контекст не подходит)

    group = equiv_groups.get(group_name, [])

    if not group:
        # fallback
        group = [21]

    # === Content fit check ===
    title = slide_content.get("title", "")
    subtitle = slide_content.get("subtitle", "")

    def fits(donor):
        slots = donor.get("slots", {})
        # Проверяем title слот (если есть title в content)
        for slot_name in ("title", "title_caption", "title_main"):
            if slot_name in slots and title:
                if not fits_donor(title, slots[slot_name]):
                    return False
                break
        # Проверяем subtitle (если есть в content)
        for slot_name in ("subtitle", "subtitle_caption"):
            if slot_name in slots and subtitle:
                if not fits_donor(subtitle, slots[slot_name]):
                    return False
                break
        return True

    chosen = pick_donor_from_group(group, donors_yaml, fits, recently_used, total_used)

    # === FINALE PREFERENCE v0.18: smart selection ===
    # Donor 86 (photo_full_dark) — ВИЗУАЛЬНО РИЧЕЙШИЙ финал. Имеет фотофон через layout
    # + 3 опциональные caption-карточки. Подходит для простых finals «Спасибо!».
    # Donor 78 (logo_with_3d) — ТОЛЬКО для product showcase с 3 unique descs.
    # Donor 25 — fallback (минимальный Green).
    # Стратегия: photo > 3D > minimal.
    if group_name == "logo_finale":
        descs = slide_content.get("descriptions") or []
        has_real_descs = len(descs) >= 3 and len(set(descs)) == len(descs)
        if has_real_descs and 78 in group:
            chosen = 78  # продуктовая выкладка
        elif 86 in group and chosen != 86:
            chosen = 86  # эмоциональный финал с фотофоном — DEFAULT для простых finals

    # === Заполняем slots ===
    donor_def = donors_yaml.get(chosen, {})
    slot_defs = donor_def.get("slots", {})
    plan_slots = {}

    # title
    for slot_name in ("title", "title_caption", "title_main"):
        if slot_name in slot_defs and title:
            plan_slots[slot_name] = title
            break

    # subtitle
    for slot_name in ("subtitle", "subtitle_caption"):
        if slot_name in slot_defs and subtitle:
            plan_slots[slot_name] = subtitle
            break

    # === FINALE ENRICHMENT v0.17.1 ===
    # Donor 78 (logo_with_3d) выбирается ТОЛЬКО если есть 3 unique descs
    # (см. selection logic выше). Здесь просто маппим descriptions без дублирования.
    if chosen == 78:
        descs = slide_content.get("descriptions") or []
        for i, d in enumerate(descs[:3], 1):
            plan_slots[f"desc{i}"] = d

    # number (для divider'а)
    if "number" in slot_defs and slide_content.get("number"):
        plan_slots["number"] = slide_content["number"]

    # body / quote
    body_list = slide_content.get("body", [])
    if "quote" in slot_defs and body_list:
        plan_slots["quote"] = body_list[0]
    elif "body" in slot_defs and body_list:
        plan_slots["body"] = "\n".join(body_list[:3]) if isinstance(body_list, list) else body_list

    # multicolumn (3 col)
    subtitles = slide_content.get("subtitles", [])
    if subtitles and chosen == 34:
        for i, sub in enumerate(subtitles[:3], 1):
            plan_slots[f"sub{i}"] = sub
        bodies = slide_content.get("col_bodies") or body_list
        for i, b in enumerate(bodies[:3], 1):
            plan_slots[f"body{i}"] = b

    # multicolumn (4 block)
    if subtitles and chosen == 29:
        slot_pos = ["sub1_top_l", "sub2_top_r", "sub3_bot_l", "sub4_bot_r"]
        body_pos = ["body1_top_l", "body2_top_r", "body3_bot_l", "body4_bot_r"]
        for i, sub in enumerate(subtitles[:4]):
            plan_slots[slot_pos[i]] = sub
        bodies = slide_content.get("col_bodies") or body_list
        for i, b in enumerate(bodies[:4]):
            plan_slots[body_pos[i]] = b

    # KPI numbers
    numbers = slide_content.get("numbers", [])
    descriptions = slide_content.get("descriptions", [])
    if chosen in (43, 44) and numbers:
        slot_n = ["num_top_l", "num_top_r", "num_bottom"]
        slot_d = ["desc_top_l", "desc_top_r", "desc_bottom"]
        for i, n in enumerate(numbers[:3]):
            plan_slots[slot_n[i]] = str(n)
        for i, d in enumerate(descriptions[:3]):
            plan_slots[slot_d[i]] = d

    result = {
        "_donor_category": f"{group_name} (donor {chosen})",
        "clone_from_slide": chosen,
        "slots": plan_slots,
    }

    # Image insertion — ТОЛЬКО для image-friendly donors.
    # ROOT CAUSE FIX v0.17.2: divider/callout/logo не должны принимать image content,
    # т.к. они decorative_shell (паттерны/декор не оставляют места для картинки).
    # Картинка перекрывает title (testpart slide 5: "Безопасность" обрезана донат-диаграммой).
    IMAGE_FRIENDLY_DONORS = {
        4, 5, 6, 7,        # title doors — картинка-аксессуар справа от title
        21, 22, 73,        # screenshot doors
        39, 40, 79, 81, 86,  # image-content doors
    }
    if image_path and chosen in IMAGE_FRIENDLY_DONORS:
        image_geometry = {
            79: {"left_px": 440, "top_px": 220, "width_px": 400, "height_px": 360},
            81: {"left_px": 35, "top_px": 100, "width_px": 580, "height_px": 540},
            73: {"left_px": 107, "top_px": 144, "width_px": 853, "height_px": 536},
            39: {"left_px": 35, "top_px": 100, "width_px": 580, "height_px": 540},
            # title doors — картинка справа, не должна перекрывать title
            4:  {"left_px": 720, "top_px": 80, "width_px": 540, "height_px": 480},
            5:  {"left_px": 720, "top_px": 80, "width_px": 540, "height_px": 480},
            6:  {"left_px": 720, "top_px": 80, "width_px": 540, "height_px": 480},
            7:  {"left_px": 720, "top_px": 80, "width_px": 540, "height_px": 480},
            # fallback content text doors
            21: {"left_px": 720, "top_px": 80, "width_px": 540, "height_px": 480},
            22: {"left_px": 720, "top_px": 80, "width_px": 540, "height_px": 480},
        }
        geom = image_geometry.get(chosen, {"left_px": 720, "top_px": 80, "width_px": 540, "height_px": 480})
        result["pictures"] = [{"file": image_path, **geom}]

    # Table data — пробрасываем в plan для build_v7
    if slide_content.get("table_data"):
        result["table_data"] = slide_content["table_data"]

    return result


def enrich_with_images(classified, manifest_path):
    """Для каждого slide в classified добавляет image_path если в manifest есть БОЛЬШАЯ картинка.
    Мелкие иконки (<200x200 px) игнорируются — они декор, не контент.

    classified slides indexed by 'num' (1-based).
    """
    if not manifest_path or not os.path.exists(manifest_path):
        return classified
    manifest = json.load(open(manifest_path, encoding="utf-8"))

    # Находим images_dir
    base_dir = os.path.dirname(manifest_path)
    base_name = os.path.basename(manifest_path).replace("_manifest.json", "")
    images_dir_candidates = [
        os.path.join(base_dir, f"{base_name}_images"),    # v15_slide_graph_images (наш паттерн)
        os.path.join(base_dir, base_name),                 # без суффикса
        manifest_path.replace("_manifest.json", ""),
        os.path.join(base_dir, manifest.get("source", "").split("/")[-1].replace(".pptx", "_images")),
    ]

    images_dir = None
    for cand in images_dir_candidates:
        if os.path.isdir(cand):
            images_dir = cand
            break
    if not images_dir:
        print(f"WARN: images_dir not found. Tried: {images_dir_candidates}", file=sys.stderr)
        return classified

    # Группируем images по slide_num и фильтруем большие
    images_by_slide = {}
    for img in manifest.get("images", []):
        if img.get("width_px", 0) >= 200 and img.get("height_px", 0) >= 200:
            slide_num = img["slide_num"]
            if slide_num not in images_by_slide:
                images_by_slide[slide_num] = []
            full_path = os.path.join(images_dir, img["file"])
            if os.path.exists(full_path):
                images_by_slide[slide_num].append({
                    "path": full_path,
                    "width": img["width_px"],
                    "height": img["height_px"],
                })

    # Обогащаем classified
    enriched_slides = []
    for cs in classified.get("slides", []):
        num = cs.get("num")
        if num in images_by_slide and not cs.get("image_path"):
            cs = dict(cs)
            largest = max(images_by_slide[num], key=lambda i: i["width"] * i["height"])
            cs["image_path"] = largest["path"]
            cs["_image_auto_added"] = True
        enriched_slides.append(cs)

    return {**classified, "slides": enriched_slides}


def design_plan(classified, smap):
    """Главная функция. Превращает classified в plan.
    Для длинных презентаций (≥15 sl) применяет дополнительный variety boost."""
    donors_yaml = smap.get("donors", {})
    equiv_groups = dict(smap.get("category_equivalence", {}))  # копия

    n_slides = len(classified.get("slides", []))

    # === VARIETY BOOST для длинных презентаций ===
    if n_slides >= 12:
        # Расширяем группы добавлением соседних
        equiv_groups["title_open"] = [4, 5, 6, 40]  # +40
        equiv_groups["divider"] = [12, 13, 62, 10]  # +10
        equiv_groups["content_text"] = [21, 22, 81]  # +81 для разнообразия
        equiv_groups["callout"] = [41, 42]
        equiv_groups["content_4block"] = [29]
        # При длинной презентации divider может быть равноценным с title_open для второстепенных переходов

    plan_slides = []
    recently_used = []  # последние 5 donor для anti-monotony
    total_used = Counter()  # сколько раз каждый donor использован → round-robin

    for content in classified.get("slides", []):
        slide_plan = design_slide(content, donors_yaml, equiv_groups, recently_used, total_used)
        plan_slides.append(slide_plan)
        donor_n = slide_plan["clone_from_slide"]
        recently_used.append(donor_n)
        recently_used = recently_used[-5:]
        total_used[donor_n] += 1

    # === Variety check ===
    used_counter = Counter(s["clone_from_slide"] for s in plan_slides)
    n_total = len(plan_slides)
    n_unique = len(used_counter)
    variety_pct = round(n_unique / n_total * 100, 1) if n_total else 0

    return {
        "_designed_by": "layout_designer.py v0.13",
        "_variety": {
            "total_slides": n_total,
            "unique_donors": n_unique,
            "variety_pct": variety_pct,
            "donor_usage": dict(used_counter.most_common()),
        },
        "slides": plan_slides,
    }


def main():
    if len(sys.argv) < 4:
        print("Usage: layout_designer.py <classified.json> <slot_map.yaml> <output_plan.json> [manifest.json]",
              file=sys.stderr)
        sys.exit(1)

    classified_path = sys.argv[1]
    smap_path = sys.argv[2]
    out_path = sys.argv[3]
    manifest_path = sys.argv[4] if len(sys.argv) > 4 else None

    classified, smap = load_inputs(classified_path, smap_path)
    if manifest_path:
        classified = enrich_with_images(classified, manifest_path)
        n_with_img = sum(1 for s in classified["slides"] if s.get("image_path"))
        print(f"  Enriched with images: {n_with_img} slides have image_path", file=sys.stderr)
    plan = design_plan(classified, smap)

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(plan, f, ensure_ascii=False, indent=2)

    v = plan["_variety"]
    print(f"\n{'=' * 60}", file=sys.stderr)
    print(f"Layout Designer — {classified_path}", file=sys.stderr)
    print(f"{'=' * 60}", file=sys.stderr)
    print(f"Slides designed: {v['total_slides']}", file=sys.stderr)
    print(f"Unique donors:   {v['unique_donors']}", file=sys.stderr)
    print(f"Variety:         {v['variety_pct']}%", file=sys.stderr)
    print(f"Donor usage:     {v['donor_usage']}", file=sys.stderr)
    print(f"\nPlan → {out_path}", file=sys.stderr)


if __name__ == "__main__":
    main()
