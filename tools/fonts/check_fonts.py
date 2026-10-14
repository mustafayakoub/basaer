#!/usr/bin/env python3
"""Shape every ayah of each riwaya with candidate fonts and count failures.

A failure is a missing glyph (.notdef), a dotted circle (U+25CC) that the
shaper inserts when a mark has no base it can attach to, or a placeholder
glyph: KFGQPC fonts map characters they do not support (e.g. ے and ۟ in the
Hafs font) to one shared dotted-circle outline, which shapers accept silently.
A font is fit for a riwaya only when the whole text shapes with zero failures.

Reads data/fonts/riwaya_fonts.json, scans data/fonts/, and writes
tools/fonts/font_report.json. Needs: pip install uharfbuzz fonttools

Usage: python3 tools/fonts/check_fonts.py
"""
import hashlib
import json
import re
from collections import defaultdict
from pathlib import Path

import uharfbuzz as hb
from fontTools.ttLib import TTFont

ROOT = Path(__file__).resolve().parents[2]
FONTS = ROOT / "data/fonts"
TAG = re.compile(r"<[^>]*>")


def placeholder_glyphs(path):
    """Glyph ids whose outline is shared by 5+ unrelated characters."""
    tt = TTFont(path)
    if "glyf" not in tt:
        return set()
    glyf, order = tt["glyf"], tt.getGlyphOrder()
    by_outline = defaultdict(set)
    for cp, name in tt.getBestCmap().items():
        g = glyf[name]
        if g.numberOfContours:
            key = hashlib.md5(repr((g.numberOfContours, list(g.getCoordinates(glyf)[0]))).encode()).digest()
            by_outline[key].add(name)
    names = set().union(*[v for v in by_outline.values() if len(v) >= 5]) if by_outline else set()
    return {order.index(n) for n in names}


def load(path):
    font = hb.Font(hb.Face(hb.Blob.from_file_path(str(path))))
    bad = placeholder_glyphs(path) | {0}
    dotted = font.get_nominal_glyph(0x25CC)
    if dotted:
        bad.add(dotted)
    return font, bad


def failures(font, bad, text):
    buf = hb.Buffer()
    buf.add_str(text)
    buf.guess_segment_properties()
    hb.shape(font, buf)
    return sum(1 for g in buf.glyph_infos if g.codepoint in bad)


def family(path):
    try:
        return TTFont(path, lazy=True)["name"].getDebugName(4) or ""
    except Exception:
        return ""


def main():
    spec = json.loads((FONTS / "riwaya_fonts.json").read_text(encoding="utf-8"))
    fonts = {p.name: family(p) for p in sorted(FONTS.iterdir()) if p.suffix.lower() in (".ttf", ".otf")}
    report = {"fonts_in_repo": fonts, "riwayat": {}}
    for key, r in spec["riwayat"].items():
        entry = {"data": r["data"], "expected_family": r["expected_family"]}
        if not r["data"]:
            entry["status"] = "data_missing"
            report["riwayat"][key] = entry
            continue
        ayat = json.loads((ROOT / r["data"]).read_text(encoding="utf-8"))
        field = "tajweedText" if "tajweedText" in ayat[0] else "aya_text"
        texts = [TAG.sub("", a[field]) for a in ayat]
        matches = [f for f, fam in fonts.items() if r["expected_family"].lower() in fam.lower()]
        candidates = matches or [spec["fallback_font"]]
        results = {}
        for name in candidates:
            font, dotted = load(FONTS / name)
            bad = [(a.get("sura_no", a.get("surahNo")), a.get("aya_no", a.get("ayahNo")))
                   for a, t in zip(ayat, texts) if failures(font, dotted, t)]
            results[name] = {"failing_ayat": len(bad), "first": bad[:5]}
        entry["tested"] = results
        ok = [n for n in matches if results[n]["failing_ayat"] == 0]
        entry["status"] = "ok" if ok else ("font_fails" if matches else "font_missing")
        entry["font"] = ok[0] if ok else None
        report["riwayat"][key] = entry
        fb = results.get(spec["fallback_font"])
        print(f"{r['ar']:8} {entry['status']:13} " + (f"fallback {spec['fallback_font']}: {fb['failing_ayat']} failing ayat" if fb else f"font={entry['font']}"))
    (ROOT / "tools/fonts/font_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
