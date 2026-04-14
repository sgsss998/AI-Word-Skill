#!/usr/bin/env python3
"""
Write docs/images/readme-compare-autogen-quicklook.png (macOS).

This is a lightweight Quick Look collage for developers. The README hero image
(`readme-hero-word-desktop.png`) is a separate, full Word desktop capture.

Requires: python-docx, Pillow; system qlmanage (Quick Look) for .docx thumbnails.
Run from repo root: python scripts/render_readme_compare_figure.py
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

REPO = Path(__file__).resolve().parents[1]
DEMO = REPO / "demo"
OUT = DEMO / "out"
IMG = REPO / "docs" / "images"


def main() -> int:
    if sys.platform != "darwin":
        print("This script uses macOS qlmanage; run on a Mac or update the thumbnail step.", file=sys.stderr)
        return 1

    subprocess.run(
        [sys.executable, str(REPO / "scripts" / "build_demo_template.py")],
        cwd=REPO,
        check=True,
    )
    subprocess.run(
        [
            sys.executable,
            str(REPO / "scripts" / "compare_sop_vs_paragraph_text.py"),
            "--template",
            str(DEMO / "demo-template.docx"),
            "--out-dir",
            str(OUT),
        ],
        cwd=REPO,
        check=True,
    )

    IMG.mkdir(parents=True, exist_ok=True)
    sop_docx = OUT / "compare-sop-rewrite-paragraph.docx"
    bad_docx = OUT / "compare-bad-paragraph-text.docx"
    for f in (sop_docx, bad_docx):
        subprocess.run(
            ["qlmanage", "-t", "-s", "1600", "-o", str(IMG), str(f)],
            cwd=REPO,
            check=True,
        )

    sop_png = IMG / "compare-sop-rewrite-paragraph.docx.png"
    bad_png = IMG / "compare-bad-paragraph-text.docx.png"
    sop = Image.open(sop_png)
    bad = Image.open(bad_png)
    # README convention: pitfall on the left, SOP on the right (matches hero screenshot).
    left, right = bad, sop
    w1, w2 = left.width, right.width
    h = max(left.height, right.height)
    pad, label_h, gap, footer_h = 24, 56, 16, 28
    total_w = pad * 2 + w1 + gap + w2
    total_h = pad * 2 + label_h + h + footer_h
    canvas = Image.new("RGB", (total_w, total_h), (245, 245, 245))
    draw = ImageDraw.Draw(canvas)
    try:
        font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 22)
        font_sm = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 14)
    except OSError:
        font = ImageFont.load_default()
        font_sm = font

    draw.text((pad, pad), "Anti-pattern: paragraph.text = (wipes runs)", fill=(120, 20, 20), font=font)
    draw.text(
        (pad + w1 + gap, pad),
        "SOP: rewrite_paragraph (keep first run rPr)",
        fill=(15, 80, 15),
        font=font,
    )
    y0 = pad + label_h
    canvas.paste(left, (pad, y0))
    canvas.paste(right, (pad + w1 + gap, y0))
    draw.text(
        (pad, y0 + h + 8),
        "Same template & replacement text · macOS Quick Look preview of .docx first page",
        fill=(80, 80, 80),
        font=font_sm,
    )
    final = IMG / "readme-compare-autogen-quicklook.png"
    canvas.save(final, "PNG", optimize=True)
    for p in (sop_png, bad_png):
        p.unlink(missing_ok=True)
    print("Wrote", final)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
