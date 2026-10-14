#!/usr/bin/env python3
"""Copy the Quran font into site/assets/fonts and draw the app icons and share image.

Icons use the KFGQPC Hafs font for the word «بصائر», gold on the brand blue.
Needs Pillow built with libraqm (Arabic shaping). Usage: python3 tools/site/make_assets.py
"""
import shutil
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, features

ROOT = Path(__file__).resolve().parents[2]
SITE = ROOT / "site"
FONT = ROOT / "data/fonts/UthmanicHafs_V22.ttf"
BRAND, GOLD, PAPER = (43, 79, 140), (217, 180, 106), (245, 246, 248)
WORD = "بَصَائِر"


def word(draw, box, size_ratio, fill):
    x0, y0, x1, y1 = box
    size = int((y1 - y0) * size_ratio)
    font = ImageFont.truetype(str(FONT), size, layout_engine=ImageFont.Layout.RAQM)
    draw.text(((x0 + x1) / 2, (y0 + y1) / 2), WORD, font=font, fill=fill, anchor="mm", direction="rtl", language="ar")


def icon(size, safe=1.0, radius=0.22):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    if safe < 1:  # maskable: full bleed background, content kept inside the safe zone
        d.rectangle([0, 0, size, size], fill=BRAND)
    else:
        d.rounded_rectangle([0, 0, size - 1, size - 1], radius=int(size * radius), fill=BRAND)
    pad = size * (1 - safe) / 2 + size * 0.08
    d.ellipse([pad, pad, size - pad, size - pad], outline=GOLD, width=max(2, size // 64))
    word(d, (pad, pad, size - pad, size - pad), 0.42, GOLD)
    return img


def og():
    img = Image.new("RGB", (1200, 630), BRAND)
    d = ImageDraw.Draw(img)
    d.rectangle([40, 40, 1160, 590], outline=GOLD, width=4)
    word(d, (0, 120, 1200, 420), 0.62, GOLD)
    sub = ImageFont.truetype(str(FONT), 46, layout_engine=ImageFont.Layout.RAQM)
    d.text((600, 470), "القرآن الكريم ومصاحفه وعلومه", font=sub, fill=PAPER, anchor="mm", direction="rtl", language="ar")
    return img


def main():
    assert features.check("raqm"), "Pillow needs libraqm for Arabic shaping"
    (SITE / "assets/fonts").mkdir(parents=True, exist_ok=True)
    shutil.copy(FONT, SITE / "assets/fonts/UthmanicHafs_V22.ttf")
    try:
        from fontTools.ttLib import TTFont
        f = TTFont(FONT)
        f.flavor = "woff2"
        f.save(SITE / "assets/fonts/UthmanicHafs_V22.woff2")
    except Exception as e:  # brotli missing: the CSS falls back to the .ttf
        print("woff2 skipped:", e)
    icons = SITE / "icons"
    icons.mkdir(exist_ok=True)
    icon(192).save(icons / "icon-192.png")
    icon(512).save(icons / "icon-512.png")
    icon(512, safe=0.8).save(icons / "maskable-512.png")
    icon(180, radius=0).convert("RGB").save(icons / "apple-touch-icon.png")
    og().save(icons / "og.png")
    print("assets written to", SITE)


if __name__ == "__main__":
    main()
