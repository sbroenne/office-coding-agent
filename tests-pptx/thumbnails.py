#!/usr/bin/env python3
"""Create thumbnail grid from PowerPoint slides using Spire.Presentation.

Usage:
    python3 tests-pptx/thumbnails.py tests-pptx/test.pptx
    python3 tests-pptx/thumbnails.py tests-pptx/test.pptx --cols 4
"""

import argparse
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from spire.presentation import Presentation
from spire.presentation.common import Stream

THUMB_W = 480
THUMB_H = 270
COLS = 3
PAD = 16
BORDER = 2
FONT_SIZE = 18
JPEG_Q = 92


def render_slides(pptx_path: Path, tmp_dir: Path):
    """Render each slide to PNG via Spire, return list of (path, label)."""
    prs = Presentation()
    prs.LoadFromFile(str(pptx_path))
    slides = []
    for i in range(prs.Slides.Count):
        out = tmp_dir / f"slide{i + 1}.png"
        img_stream = prs.Slides[i].SaveAsImageByWH(THUMB_W * 2, THUMB_H * 2)
        fs = Stream(str(out))
        img_stream.CopyTo(fs)
        fs.Close()
        img_stream.Close()
        slides.append((out, f"Slide {i + 1}"))
    prs.Dispose()
    return slides


def make_grid(slides, cols, out_path):
    """Stitch slide thumbnails into a labeled grid."""
    rows = (len(slides) + cols - 1) // cols
    label_h = FONT_SIZE + 8
    cell_w = THUMB_W + PAD
    cell_h = THUMB_H + label_h + PAD
    grid_w = cols * cell_w + PAD
    grid_h = rows * cell_h + PAD

    grid = Image.new("RGB", (grid_w, grid_h), "white")
    draw = ImageDraw.Draw(grid)
    try:
        font = ImageFont.load_default(size=FONT_SIZE)
    except Exception:
        font = ImageFont.load_default()

    for idx, (img_path, label) in enumerate(slides):
        r, c = divmod(idx, cols)
        x = PAD + c * cell_w
        y = PAD + r * cell_h

        # Label
        bbox = draw.textbbox((0, 0), label, font=font)
        tw = bbox[2] - bbox[0]
        draw.text((x + (THUMB_W - tw) // 2, y), label, fill="black", font=font)

        # Thumbnail
        ty = y + label_h
        with Image.open(img_path) as img:
            img.thumbnail((THUMB_W, THUMB_H), Image.Resampling.LANCZOS)
            w, h = img.size
            tx = x + (THUMB_W - w) // 2
            tty = ty + (THUMB_H - h) // 2
            grid.paste(img, (tx, tty))
            if BORDER:
                draw.rectangle(
                    [(tx - BORDER, tty - BORDER), (tx + w + BORDER - 1, tty + h + BORDER - 1)],
                    outline="gray", width=BORDER,
                )

    grid.save(str(out_path), quality=JPEG_Q)
    return out_path


def main():
    parser = argparse.ArgumentParser(description="PPTX → thumbnail grid")
    parser.add_argument("input", help=".pptx file")
    parser.add_argument("--cols", type=int, default=COLS)
    parser.add_argument("-o", "--output", help="Output path (default: <input>_grid.jpg)")
    args = parser.parse_args()

    pptx = Path(args.input)
    if not pptx.exists():
        print(f"Not found: {pptx}", file=sys.stderr)
        sys.exit(1)

    out = Path(args.output) if args.output else pptx.with_name(pptx.stem + "_grid.jpg")

    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        slides = render_slides(pptx, Path(tmp))
        if not slides:
            print("No slides found", file=sys.stderr)
            sys.exit(1)
        make_grid(slides, args.cols, out)
        print(f"✓ {len(slides)} slides → {out}")


if __name__ == "__main__":
    main()
