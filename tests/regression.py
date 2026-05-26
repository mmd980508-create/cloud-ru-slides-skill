#!/usr/bin/env python3
"""
regression.py — прогон baseline-кейсов через полный pipeline и сверка с
ожидаемыми значениями. Возвращает exit 0 (PASS) / 1 (regression).

Usage:
    python3 tests/regression.py [case_name]

Без аргументов — гоняет все baseline в tests/baselines/.
С аргументом — гоняет только указанный кейс (например `humanity`).

Структура baseline:
    tests/baselines/<case>/plan.json       — input
    tests/baselines/<case>/expected.json   — ожидаемые метрики

Тестер:
    1. Запускает scripts/e2e_pipeline.sh <plan.json> <tmp_out>
    2. Читает brand_report.json и visual_report.json
    3. Сравнивает с expected.json
    4. PASS / FAIL per case
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
BASELINES_DIR = os.path.join(HERE, "baselines")
PIPELINE_SCRIPT = os.path.join(SKILL_DIR, "scripts", "e2e_pipeline.sh")


GREEN = "\033[32m"
RED = "\033[31m"
YELLOW = "\033[33m"
RESET = "\033[0m"


def run_case(case_name):
    """Returns dict {case, ok, errors, details}."""
    case_dir = os.path.join(BASELINES_DIR, case_name)
    plan_path = os.path.join(case_dir, "plan.json")
    expected_path = os.path.join(case_dir, "expected.json")

    if not os.path.isfile(plan_path):
        return {"case": case_name, "ok": False,
                "errors": [f"plan не найден: {plan_path}"]}
    if not os.path.isfile(expected_path):
        return {"case": case_name, "ok": False,
                "errors": [f"expected не найден: {expected_path}"]}

    with open(expected_path, encoding="utf-8") as f:
        expected = json.load(f)

    # Прогон в temp dir
    with tempfile.TemporaryDirectory(prefix=f"regression_{case_name}_") as tmpdir:
        proc = subprocess.run(
            ["bash", PIPELINE_SCRIPT, plan_path, tmpdir],
            cwd=SKILL_DIR,
            capture_output=True,
            text=True,
            timeout=600,
        )
        if proc.returncode != 0:
            return {"case": case_name, "ok": False,
                    "errors": [f"e2e_pipeline.sh упал (exit {proc.returncode}): {proc.stderr[-500:]}"]}

        brand_path = os.path.join(tmpdir, "brand_report.json")
        visual_path = os.path.join(tmpdir, "visual_report.json")
        if not os.path.isfile(brand_path):
            return {"case": case_name, "ok": False,
                    "errors": ["brand_report.json не создан"]}
        if not os.path.isfile(visual_path):
            return {"case": case_name, "ok": False,
                    "errors": ["visual_report.json не создан"]}

        with open(brand_path, encoding="utf-8") as f:
            brand = json.load(f)
        with open(visual_path, encoding="utf-8") as f:
            visual = json.load(f)

    errors = []
    details = {}

    # n_slides
    expected_n = expected.get("n_slides")
    actual_n = brand.get("n_slides")
    details["n_slides"] = {"expected": expected_n, "actual": actual_n}
    if expected_n is not None and actual_n != expected_n:
        errors.append(f"n_slides: ожидалось {expected_n}, получено {actual_n}")

    # brand_guardian
    bg_exp = expected.get("brand_guardian", {})
    bg_verdict_exp = bg_exp.get("verdict")
    bg_verdict_act = brand.get("verdict") or brand.get("summary", {}).get("verdict")
    bg_score = brand.get("summary", {}).get("score_avg")
    bg_violations = brand.get("summary", {}).get("violations_total", 0)
    details["brand_guardian"] = {
        "verdict": bg_verdict_act,
        "score_avg": bg_score,
        "violations_total": bg_violations,
    }
    if bg_verdict_exp and bg_verdict_act != bg_verdict_exp:
        errors.append(f"brand_guardian.verdict: ожидался {bg_verdict_exp}, получен {bg_verdict_act}")
    score_min = bg_exp.get("score_avg_min")
    if score_min is not None and bg_score is not None and bg_score < score_min:
        errors.append(f"brand_guardian.score_avg: {bg_score} < min {score_min}")
    viol_max = bg_exp.get("violations_max")
    if viol_max is not None and bg_violations > viol_max:
        errors.append(f"brand_guardian.violations: {bg_violations} > max {viol_max}")

    # visual_validator_v2 (с --plan)
    vv_exp = expected.get("visual_validator_v2_with_plan", {})
    vv_verdict_exp = vv_exp.get("verdict")
    vv_verdict_act = visual.get("verdict")
    vv_issues = sum(len(s.get("issues", [])) for s in visual.get("slides", []))
    details["visual_validator_v2"] = {
        "verdict": vv_verdict_act,
        "total_issues": vv_issues,
    }
    if vv_verdict_exp and vv_verdict_act != vv_verdict_exp:
        errors.append(f"visual_validator_v2.verdict: ожидался {vv_verdict_exp}, получен {vv_verdict_act}")
    iss_max = vv_exp.get("total_issues_max")
    if iss_max is not None and vv_issues > iss_max:
        errors.append(f"visual_validator_v2.total_issues: {vv_issues} > max {iss_max}")

    return {"case": case_name, "ok": len(errors) == 0,
            "errors": errors, "details": details}


def main():
    selected = sys.argv[1] if len(sys.argv) > 1 else None

    if not os.path.isfile(PIPELINE_SCRIPT):
        print(f"{RED}ERROR: pipeline скрипт не найден: {PIPELINE_SCRIPT}{RESET}", file=sys.stderr)
        sys.exit(2)

    if not os.path.isdir(BASELINES_DIR):
        print(f"{RED}ERROR: нет папки baselines: {BASELINES_DIR}{RESET}", file=sys.stderr)
        sys.exit(2)

    if selected:
        cases = [selected]
    else:
        cases = sorted(d for d in os.listdir(BASELINES_DIR)
                       if os.path.isdir(os.path.join(BASELINES_DIR, d)))

    if not cases:
        print(f"{YELLOW}Нет baseline-кейсов в {BASELINES_DIR}{RESET}")
        sys.exit(0)

    print(f"\n{'=' * 70}")
    print(f"Cloud.ru Slides Skill — regression suite")
    print(f"Cases: {', '.join(cases)}")
    print(f"{'=' * 70}\n")

    results = []
    for case in cases:
        print(f"--- Running {case} ---")
        r = run_case(case)
        results.append(r)
        if r["ok"]:
            print(f"{GREEN}✓ {case} PASS{RESET}")
            for k, v in r.get("details", {}).items():
                print(f"    {k}: {v}")
        else:
            print(f"{RED}✗ {case} FAIL{RESET}")
            for e in r["errors"]:
                print(f"    {RED}!{RESET} {e}")
            for k, v in r.get("details", {}).items():
                print(f"    {k}: {v}")
        print()

    failed = [r for r in results if not r["ok"]]

    print(f"{'=' * 70}")
    if failed:
        print(f"{RED}FAILED: {len(failed)} / {len(results)} cases{RESET}")
        for r in failed:
            print(f"  ✗ {r['case']}")
        sys.exit(1)
    else:
        print(f"{GREEN}ALL PASS: {len(results)} / {len(results)} cases{RESET}")
        sys.exit(0)


if __name__ == "__main__":
    main()
