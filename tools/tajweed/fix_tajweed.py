#!/usr/bin/env python3
"""Repair the madd muttasil/munfasil tags (6/18) in the tajweed text.

The source file mixes the two tags: 666 spans open with one and close with
the other, and 445 spans tagged 6 (muttasil) are followed by a hamza in the
next word (e.g. «قَالُوٓاْ إِنَّا»), which makes them munfasil. Tags cannot be
trusted here, so every 6/18 span is re-classified from its written context:

  hamza later in the same written word  -> 6  (muttasil)
  hamza at the start of the next word   -> 18 (munfasil)

A silent alif (<13>…</13>) right after the span is skipped, as in «قَالُوٓاْ».
Spans whose context is unclear keep their closing tag and are reported.

It also swaps three Tanzil marks for their KFGQPC equivalents, which the
KFGQPC Hafs font (the font hafsData_v2-0.json is encoded for) can draw:
  U+06EB → U+06EC  ishmam in «تَأْمَ۬نَّا» 12:11
  U+06DF → U+06E0  alif dropped in wasl, read in waqf (لَّـٰكِنَّا، ٱلظُّنُونَا، قَوَارِيرَا)
  U+06E3 → U+06DC  small seen in «ٱلْمُصَۜيْطِرُونَ» 52:37

Usage: python3 tools/tajweed/fix_tajweed.py [--in PATH] [--out PATH] [--report PATH]
"""
import argparse
import json
import re
import unicodedata
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
HAMZA = set("ءأإؤئآٔ")
SPAN = re.compile(r"<(6|18)>(.*?)</(6|18)>")
TAG = re.compile(r"<[^>]*>")
LEADING_SILENT_ALIF = re.compile(r"^((?:<[^>]*>)*)<13>.*?</13>")
KFGQPC_MARKS = str.maketrans({"۫": "۬", "۟": "۠", "ۣ": "ۜ"})


def context_kind(text, end):
    """Return 6, 18 or None from the letters that follow a madd span."""
    rest = LEADING_SILENT_ALIF.sub(r"\1", text[end:], count=1)
    for ch in TAG.sub("", rest):
        if ch in HAMZA:
            return 6
        if ch == " ":
            return 18
        if unicodedata.category(ch) == "Mn" or ch in "‌ـ":
            continue
        return None
    return None


def fix_text(text, stats, where, unresolved):
    out, last = [], 0
    for m in SPAN.finditer(text):
        opened, closed = int(m.group(1)), int(m.group(3))
        kind = context_kind(text, m.end())
        if kind is None:
            kind = closed
            unresolved.append({"sura": where[0], "aya": where[1], "span": TAG.sub("", m.group(0))})
        stats[(opened, closed, kind)] += 1
        out.append(text[last:m.start()])
        out.append(f"<{kind}>{m.group(2)}</{kind}>")
        last = m.end()
    out.append(text[last:])
    return "".join(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="src", default=ROOT / "data/ayat/quran_text_with_tajweed.json")
    ap.add_argument("--out", default=ROOT / "data/ayat/quran_text_with_tajweed_v2.json")
    ap.add_argument("--report", default=ROOT / "tools/tajweed/fix_report.json")
    args = ap.parse_args()

    ayat = json.loads(Path(args.src).read_text(encoding="utf-8"))
    stats, unresolved = Counter(), []
    marks = 0
    for a in ayat:
        text = fix_text(a["tajweedText"], stats, (a["surahNo"], a["ayahNo"]), unresolved)
        swapped = text.translate(KFGQPC_MARKS)
        marks += sum(x != y for x, y in zip(text, swapped))
        a["tajweedText"] = swapped

    Path(args.out).write_text(json.dumps(ayat, ensure_ascii=False, indent=1), encoding="utf-8")

    changed = sum(n for (o, c, k), n in stats.items() if not (o == c == k))
    report = {
        "spans": sum(stats.values()),
        "changed": changed,
        "by_case": {f"open{o}_close{c}_to{k}": n for (o, c, k), n in sorted(stats.items())},
        "kfgqpc_marks_swapped": marks,
        "unresolved_kept_closing_tag": unresolved,
    }
    Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"spans={report['spans']} changed={changed} unresolved={len(unresolved)} -> {args.out}")


if __name__ == "__main__":
    main()
