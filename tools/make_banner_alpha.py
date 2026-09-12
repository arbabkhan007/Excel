#!/usr/bin/env python3
"""
Give the cover banners a transparent middle.

Excel draws inserted pictures *above* the cell grid, so text typed in the cells
underneath an opaque picture can never be seen.  The Start Here tab puts its
title in cells under the banner; for that to work the flat cream/white middle
of the artwork has to be see-through.

This script converts ``assets/banner_<theme>.jpg`` (flat background) into
``assets/banner_<theme>.png`` with an alpha channel: everything inside the
title window that matches the background colour becomes transparent, with a
one pixel feather so the watercolour edges stay soft.

    /tmp/venv/bin/python tools/make_banner_alpha.py
"""

import os

from PIL import Image, ImageFilter

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(HERE, "assets")

# title window as a fraction of the image (matches the C:H cell overlay)
WIN = (0.06, 0.02, 0.68, 0.98)
TOLERANCE = 26


def background_color(im):
    w, h = im.size
    px = im.load()
    samples = [px[int(w * 0.5), int(h * f)] for f in (0.02, 0.5, 0.98)]
    r = sum(s[0] for s in samples) // 3
    g = sum(s[1] for s in samples) // 3
    b = sum(s[2] for s in samples) // 3
    return r, g, b


def convert(src, dst):
    im = Image.open(src).convert("RGB")
    w, h = im.size
    bg = background_color(im)
    alpha = Image.new("L", (w, h), 255)
    ap = alpha.load()
    px = im.load()
    x1, y1, x2, y2 = (int(w * WIN[0]), int(h * WIN[1]),
                      int(w * WIN[2]), int(h * WIN[3]))
    for y in range(y1, y2):
        for x in range(x1, x2):
            r, g, b = px[x, y]
            if abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2]) < TOLERANCE * 3:
                ap[x, y] = 0
    alpha = alpha.filter(ImageFilter.GaussianBlur(1.1))
    out = im.convert("RGBA")
    out.putalpha(alpha)
    out.save(dst, optimize=True)
    return out.size, os.path.getsize(dst) // 1024


def main():
    for stem in ("banner_festive", "banner_minimal"):
        src = os.path.join(ASSETS, stem + ".jpg")
        dst = os.path.join(ASSETS, stem + ".png")
        if not os.path.exists(src):
            print("missing", src)
            continue
        size, kb = convert(src, dst)
        print("%-16s %s -> %s (%d KB)" % (stem, size, dst, kb))
        os.remove(src)


if __name__ == "__main__":
    main()
