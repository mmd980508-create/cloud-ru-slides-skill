#!/usr/bin/env python3
"""
visual_validator.py — пиксельный анализ rendered PNG для визуальной валидации.

Проверяет КАЖДЫЙ rendered PNG на:
1. Aspect ratio — должно быть 16:9 (±1%). Иначе render artefact.
2. Empty slide — если >97% pixels одного цвета — пустой/сломанный.
3. Edge text overflow — текст не должен касаться краёв (top/bottom/left/right 8px зона).
4. Color palette dominance — основные цвета из брендовой палитры.

Usage:
    python3 visual_validator.py <render_dir/> [report.json]

Returns:
    0 — PASS (всё ок)
    1 — WARN (потенциальные проблемы)
    2 — FAIL (критические render-артефакты или overflow)
"""
import sys
import os
import json
import glob
from collections import Counter
from PIL import Image

PALETTE_HEXES = {
    "26D07C": "Green", "CFF500": "Yellow", "A068FF": "Purple", "C0E0FC": "Blue",
    "222222": "Black", "FFFFFF": "White", "F2F2F2": "Gray",
    "000000": "Black-pure", "343F48": "Graphite-Iron",
}


def check_aspect_ratio(img):
    """Aspect ratio должен быть 16:9 = 1.778."""
    w, h = img.size
    actual = w / h
    expected = 16 / 9
    diff_pct = abs(actual - expected) / expected * 100

    if diff_pct < 1.0:
        return None
    if diff_pct < 3.0:
        return {"type": "aspect_ratio", "level": "WARN",
                "msg": f"AR={actual:.3f} (16:9={expected:.3f}, diff {diff_pct:.1f}%) — небольшое отклонение"}
    return {"type": "aspect_ratio", "level": "FAIL",
            "msg": f"AR={actual:.3f} (16:9={expected:.3f}, diff {diff_pct:.1f}%) — render artefact"}


def check_empty_slide(img, threshold_pct=97):
    """Если >97% пикселей одного цвета (±5 RGB) — слайд пустой/сломанный."""
    img_small = img.resize((128, 72)).convert("RGB")
    pixels = list(img_small.getdata())
    counts = Counter(pixels)
    most_common = counts.most_common(1)[0][1]
    total = len(pixels)
    pct = most_common / total * 100
    if pct > threshold_pct:
        return {"type": "empty_slide", "level": "WARN",
                "msg": f"{pct:.1f}% pixels одного цвета — возможно пустой слайд"}
    return None


def estimate_bg_brightness(img):
    """Возвращает 'light' / 'dark' / 'mixed' на основе center-pixel и dominants."""
    w, h = img.size
    # Sample center 50% region (avoid edges)
    cx0, cy0 = w // 4, h // 4
    cx1, cy1 = 3 * w // 4, 3 * h // 4
    center = img.crop((cx0, cy0, cx1, cy1)).resize((50, 28)).convert("RGB")
    pixels = list(center.getdata())
    avg_brightness = sum((r + g + b) / 3 for r, g, b in pixels) / len(pixels)
    if avg_brightness > 180:
        return "light"
    if avg_brightness < 80:
        return "dark"
    return "mixed"


def detect_text_pixels(strip, bg_type):
    """Возвращает % пикселей которые ЯВЛЯЮТСЯ text (контрастирующий bg).
    На светлом bg — ищем тёмные pixels.
    На тёмном bg — ищем светлые pixels."""
    pixels = list(strip.convert("RGB").getdata())
    text_count = 0
    for r, g, b in pixels:
        avg = (r + g + b) / 3
        if bg_type == "light" and avg < 80:
            text_count += 1
        elif bg_type == "dark" and avg > 180:
            text_count += 1
        elif bg_type == "mixed" and (avg < 60 or avg > 220):
            text_count += 1
    return text_count / len(pixels) * 100


def check_edge_overflow(img, edge_px=8):
    """Проверка краевых полос на наличие text-pixels (с учётом bg color)."""
    issues = []
    w, h = img.size
    bg_type = estimate_bg_brightness(img)

    strips = {
        "top":    img.crop((0, 0, w, edge_px)),
        "bottom": img.crop((0, h - edge_px, w, h)),
        "left":   img.crop((0, 0, edge_px, h)),
        "right":  img.crop((w - edge_px, 0, w, h)),
    }

    # Tolerance per side: right side имеет логотип Cloud.ru (~12-15%), top — декоративные паттерны
    side_thresholds = {
        "top":    {"warn": 20, "fail": 40},
        "bottom": {"warn": 12, "fail": 30},
        "left":   {"warn": 12, "fail": 30},
        "right":  {"warn": 25, "fail": 50},  # лого справа намеренно
    }

    for side, strip in strips.items():
        text_pct = detect_text_pixels(strip, bg_type)
        thr = side_thresholds[side]
        if text_pct > thr["fail"]:
            level = "FAIL"
        elif text_pct > thr["warn"]:
            level = "WARN"
        else:
            continue
        issues.append({
            "type": f"edge_overflow_{side}",
            "level": level,
            "msg": f"{side} edge: {text_pct:.1f}% контрастных pixels (bg={bg_type}) — возможный text overflow"
        })
    return issues


def detect_dominant_colors(img, top_n=5):
    """Top-N доминантных цветов в slide."""
    img_small = img.resize((100, 56)).convert("RGB")
    pixels = list(img_small.getdata())
    # quantize: group similar colors
    quantized = [(r // 16 * 16, g // 16 * 16, b // 16 * 16) for r, g, b in pixels]
    counts = Counter(quantized).most_common(top_n)
    total = len(pixels)
    return [(f"{r:02X}{g:02X}{b:02X}", count / total * 100) for (r, g, b), count in counts]


def check_palette_dominance(img):
    """Основные цвета должны быть из брендовой палитры (или близки к ней)."""
    dominants = detect_dominant_colors(img, top_n=3)
    palette_match_count = 0
    for hex_color, pct in dominants:
        # Проверка на близость к палитре (RGB distance < 50)
        r1, g1, b1 = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        for ph, name in PALETTE_HEXES.items():
            r2, g2, b2 = int(ph[0:2], 16), int(ph[2:4], 16), int(ph[4:6], 16)
            d = ((r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2) ** 0.5
            if d < 60:
                palette_match_count += 1
                break

    if palette_match_count == 0:
        return {"type": "palette_dominance", "level": "WARN",
                "msg": f"Top-3 цвета {[h for h,_ in dominants]} не близки к палитре"}
    return None


def validate_slide_png(png_path):
    """Полная визуальная валидация одного PNG."""
    img = Image.open(png_path)
    issues = []

    # Все checks
    for check_fn in (check_aspect_ratio, check_empty_slide, check_palette_dominance):
        result = check_fn(img)
        if result:
            issues.append(result)

    # Edge checks (multi-issue)
    issues.extend(check_edge_overflow(img))

    # Score: 100 - 25*FAIL - 10*WARN
    fails = sum(1 for i in issues if i["level"] == "FAIL")
    warns = sum(1 for i in issues if i["level"] == "WARN")
    score = max(0, 100 - 25 * fails - 10 * warns)

    return {
        "file": os.path.basename(png_path),
        "size": img.size,
        "score": score,
        "issues": issues,
        "dominant_colors": detect_dominant_colors(img, top_n=3),
    }


def main():
    if len(sys.argv) < 2:
        print("Usage: visual_validator.py <render_dir/> [report.json]", file=sys.stderr)
        sys.exit(2)

    render_dir = sys.argv[1]
    report_path = sys.argv[2] if len(sys.argv) > 2 else None

    pngs = sorted(glob.glob(os.path.join(render_dir, "slide-*.png")))
    if not pngs:
        print(f"No slide-*.png found in {render_dir}", file=sys.stderr)
        sys.exit(2)

    report = {
        "render_dir": render_dir,
        "n_slides": len(pngs),
        "slides": [],
        "summary": {"fails": 0, "warns": 0, "score_avg": 0},
    }

    for png in pngs:
        r = validate_slide_png(png)
        report["slides"].append(r)
        report["summary"]["fails"] += sum(1 for i in r["issues"] if i["level"] == "FAIL")
        report["summary"]["warns"] += sum(1 for i in r["issues"] if i["level"] == "WARN")

    if report["slides"]:
        report["summary"]["score_avg"] = round(
            sum(s["score"] for s in report["slides"]) / len(report["slides"]), 1
        )

    if report["summary"]["fails"] > 0:
        report["verdict"] = "FAIL"
        ec = 2
    elif report["summary"]["warns"] > 0:
        report["verdict"] = "WARN"
        ec = 1
    else:
        report["verdict"] = "PASS"
        ec = 0

    print(f"\n{'=' * 60}", file=sys.stderr)
    print(f"Visual Validator — {render_dir}", file=sys.stderr)
    print(f"{'=' * 60}", file=sys.stderr)
    print(f"Verdict: {report['verdict']}  |  Score avg: {report['summary']['score_avg']}/100", file=sys.stderr)
    print(f"Slides: {report['n_slides']}  |  FAIL: {report['summary']['fails']}  |  WARN: {report['summary']['warns']}", file=sys.stderr)
    print(file=sys.stderr)

    for s in report["slides"]:
        if s["issues"]:
            print(f"--- {s['file']} (score {s['score']}, size {s['size']}) ---", file=sys.stderr)
            for issue in s["issues"]:
                icon = "❌" if issue["level"] == "FAIL" else "⚠️ "
                print(f"  {icon} {issue['type']}: {issue['msg']}", file=sys.stderr)
            colors_str = ", ".join(f"#{h} ({p:.0f}%)" for h, p in s["dominant_colors"])
            print(f"     dominant: {colors_str}", file=sys.stderr)

    if report_path:
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        print(f"\nFull report → {report_path}", file=sys.stderr)

    sys.exit(ec)


if __name__ == "__main__":
    main()
