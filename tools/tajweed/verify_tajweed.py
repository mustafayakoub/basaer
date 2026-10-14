#!/usr/bin/env python3
"""Check a tajweed text file for structural and textual soundness.

  1. every tag is one of 1..18, opened and closed in order, never nested;
  2. madd 6/18 agree with their written context (see fix_tajweed.py);
  3. madd 4 (arid) sits only in the last word of the ayah;
  4. with tags removed, the letters match hafsData_v2-0.json ayah by ayah,
     after folding the two files' encoding conventions together.

Exits 1 when a check fails. Usage:
  python3 tools/tajweed/verify_tajweed.py [PATH]
"""
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from fix_tajweed import SPAN, TAG, context_kind  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
VALID = {str(n) for n in range(1, 19)}
WAQF_SAJDA = "۩"


def letters(text):
    """Fold Tanzil and KFGQPC Uthmani encodings to a comparable letter string."""
    t = TAG.sub("", text).replace(" ", " ")
    t = re.sub(r"\s*[ﰀ-﷿٠-٩]+\s*$", "", t)  # ayah-number glyph
    t = re.sub(r"[ىٮـ](?=[ً-ْٓ-ٟ]*ٰ)", "", t)  # alif maqsura before dagger alif
    t = re.sub("[اآٱ]", "", t).replace("ى", "ي")  # long alif is spelled differently in the two encodings
    t = re.sub("[أإؤئء]", "ء", t)
    return re.sub(r"[^ء-غف-ي]", "", t)


def main():
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data/ayat/quran_text_with_tajweed_v2.json"
    ayat = json.loads(path.read_text(encoding="utf-8"))
    hafs = json.loads((ROOT / "data/ayat/hafsData_v2-0.json").read_text(encoding="utf-8"))
    problems = Counter()
    samples = {}

    def flag(kind, where):
        problems[kind] += 1
        samples.setdefault(kind, where)

    if len(ayat) != 6236:
        flag("ayah_count", len(ayat))
    for a, h in zip(ayat, hafs):
        where = f"{a['surahNo']}:{a['ayahNo']}"
        if (a["surahNo"], a["ayahNo"]) != (h["sura_no"], h["aya_no"]):
            flag("ayah_order", where)
            continue
        t = a["tajweedText"]
        stack = []
        for m in re.finditer(r"<(/?)(\d+)>", t):
            if m.group(2) not in VALID:
                flag("unknown_tag", where)
            if not m.group(1):
                if stack:
                    flag("nested_tag", where)
                stack.append(m.group(2))
            elif not stack or stack.pop() != m.group(2):
                flag("mismatched_close", where)
        if stack:
            flag("unclosed_tag", where)
        for m in SPAN.finditer(t):
            if m.group(1) != m.group(3) or context_kind(t, m.end()) not in (None, int(m.group(1))):
                flag("madd_6_18_context", where)
        plain = TAG.sub("", t).rstrip(" " + WAQF_SAJDA)
        for m in re.finditer(r"<4>", t):
            if " " in TAG.sub("", t[m.end():]).strip().rstrip(WAQF_SAJDA).strip():
                flag("madd4_not_at_end", where)
        if plain and letters(t) != letters(h["aya_text"]):
            flag("letters_differ_from_hafs", where)

    for k, n in problems.most_common():
        print(f"{k}: {n}  (first: {samples[k]})")
    print("OK" if not problems else "FAILED", "-", path.name)
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    main()
