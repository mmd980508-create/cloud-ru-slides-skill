#!/usr/bin/env python3
"""
render_html.py — HTML slide → PNG через Playwright Chromium.

Usage:
    python3 render_html.py <slide.html> <output.png> [--w 1280 --h 720] [--vars JSON]
"""
import sys
import os
import json
import argparse
import re
import tempfile
import shutil


def fill_template(html, vars_dict):
    """Простая Jinja-замена {{ var }} + {% if %}…{% endif %}."""
    def if_replace(match):
        cond_var = match.group(1).strip()
        block = match.group(2)
        if "{% else %}" in block:
            true_block, false_block = block.split("{% else %}", 1)
        else:
            true_block, false_block = block, ""
        return true_block if vars_dict.get(cond_var) else false_block
    html = re.sub(r"\{%\s*if\s+(\w+)\s*%\}(.*?)\{%\s*endif\s*%\}", if_replace, html, flags=re.DOTALL)
    html = re.sub(r"\{\{\s*(\w+)\s*\}\}", lambda m: str(vars_dict.get(m.group(1).strip(), "")), html)
    return html


def render(html_path, output_png, width=1280, height=720, vars_dict=None):
    from playwright.sync_api import sync_playwright

    with open(html_path, encoding="utf-8") as f:
        html = f.read()
    if vars_dict:
        html = fill_template(html, vars_dict)

    # Подготавливаем temp dir с симлинками templates/ + fonts/
    tmp_dir = tempfile.mkdtemp(prefix="render_html_")
    src_dir = os.path.dirname(os.path.abspath(html_path))
    project_root = os.path.dirname(src_dir)  # html-engine/
    for sub in ("templates", "fonts"):
        src = os.path.join(project_root, sub)
        dst = os.path.join(tmp_dir, sub)
        if os.path.isdir(src):
            os.symlink(src, dst)

    tmp_html = os.path.join(tmp_dir, "slide.html")
    html_fixed = html.replace('href="../templates/', 'href="templates/')
    html_fixed = html_fixed.replace('src="../fonts/', 'src="fonts/')
    with open(tmp_html, "w", encoding="utf-8") as f:
        f.write(html_fixed)

    abs_png = os.path.abspath(output_png)
    os.makedirs(os.path.dirname(abs_png), exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport={"width": width, "height": height},
                                   device_scale_factor=2)
        page = ctx.new_page()
        page.goto(f"file://{tmp_html}")
        page.wait_for_load_state("networkidle")
        page.screenshot(path=abs_png, full_page=False, omit_background=False, clip={"x":0,"y":0,"width":width,"height":height})
        browser.close()

    shutil.rmtree(tmp_dir, ignore_errors=True)
    print(f"Rendered: {abs_png}", file=sys.stderr)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("html")
    ap.add_argument("output_png")
    ap.add_argument("--w", type=int, default=1280)
    ap.add_argument("--h", type=int, default=720)
    ap.add_argument("--vars", help='JSON string with template variables', default="{}")
    args = ap.parse_args()
    render(args.html, args.output_png, args.w, args.h, json.loads(args.vars))


if __name__ == "__main__":
    main()
