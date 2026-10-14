#!/usr/bin/env python3
"""Build the JSON files the static pages in site/ read.

  site/data/surahs.json        surah cards from data/surahs/suar_quran_Mustafa_Yakoub.xlsx
  site/data/tajweed/N.json     tajweed text per surah (from quran_text_with_tajweed_v2.json)
  site/data/tajweed-rules.json tag meanings, colours and definitions
  site/data/stats.json         content statistics for the "site statistics" panel
  site/data/sources.json       copied from data/sources.json, plus the works cited in the cards

Run from anywhere: python3 tools/site/build_data.py   (needs openpyxl)
"""
import json
import re
import shutil
from collections import Counter
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "site/data"

# xlsx column index -> card field. Columns 18 (duplicate of 15) and 24 (empty) are skipped.
FIELDS = {
    0: "n", 1: "name_en", 2: "name", 3: "ayat", 6: "words", 7: "letters", 8: "juz_hizb",
    9: "page_from", 10: "page_to", 12: "pages", 13: "category", 14: "type", 15: "revelation_order",
    16: "revelation_place", 17: "revealed_after", 19: "stage", 20: "context", 21: "purpose",
    23: "reason_db", 25: "reason_brief", 26: "virtues_summary", 27: "hasanat", 28: "virtues_extra",
    29: "virtues", 31: "guidance", 32: "virtues_sahih", 33: "virtues_daif", 34: "virtues_sayings",
    35: "naskh", 36: "remarks_en", 37: "rukus",
}
CITE = re.compile(r"\[([^\[\]]{2,160}?)\]")
GRADES = {"موضوع", "ضعيف", "حديث مرسل", "مرسل", "صحيح", "حسن"}
SAME_WORK = {
    "الثعلبي في تفسيره": "الثعلبي",
    "لمحات الأنوار": "لمحات الأنوار للغافقي",
    "النسائي في السنن الكبرى": "النسائي في الكبرى",
    "فضائل القرآن لأبي عبيد": "أبو عبيد في فضائل القرآن",
    "البيهقي في الشعب وغيره": "البيهقي في الشعب",
    "ابن حبيب": "ابن حبيب في رغائب القرآن",
    "ابن صخر": "ابن صخر في فضائل القرآن",
    "متفق عليه": "متفق عليه (البخاري ومسلم)",
}


def clean(value):
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, str):
        v = value.replace("_x000D_", "").replace("\r", "").strip()
        return v or None
    return value


def cited_works(text, surah_names):
    """Names of works cited in [..] brackets, skipping Quran references like [البقرة: 255]."""
    works = []
    for c in CITE.findall(text or ""):
        name = re.split(r"[:،(]|\s[\d٠-٩]", c.strip())[0].strip()
        name = re.sub(r"^(رواه|أخرجه|انظر)\s+", "", name).strip(" .")
        if name in surah_names or name in GRADES or not 2 <= len(name) <= 60:
            continue
        works.append(SAME_WORK.get(name, name))
    return works


def build_surahs():
    ws = openpyxl.load_workbook(ROOT / "data/surahs/suar_quran_Mustafa_Yakoub.xlsx", data_only=True).active
    rows = list(ws.iter_rows(values_only=True))[2:116]
    names = {clean(r[2]) for r in rows}
    cards, all_works = [], Counter()
    for r in rows:
        card = {f: clean(r[i]) for i, f in FIELDS.items()}
        if card["virtues_extra"] == card["virtues"]:
            card["virtues_extra"] = None
        works = Counter()
        for f in ("reason_db", "virtues_summary", "virtues", "virtues_extra", "guidance",
                  "virtues_sahih", "virtues_daif", "virtues_sayings", "naskh"):
            works.update(cited_works(card[f] if isinstance(card[f], str) else "", names))
        card["cited"] = [{"work": w, "count": n} for w, n in works.most_common()]
        all_works.update(works)
        cards.append(card)
    assert [c["n"] for c in cards] == list(range(1, 115)), "surah numbers must run 1..114"
    return cards, all_works


def build_tajweed():
    ayat = json.loads((ROOT / "data/ayat/quran_text_with_tajweed_v2.json").read_text(encoding="utf-8"))
    by = {}
    for a in ayat:
        by.setdefault(a["surahNo"], []).append(a["tajweedText"])
    (OUT / "tajweed").mkdir(parents=True, exist_ok=True)
    for s, texts in by.items():
        (OUT / f"tajweed/{s}.json").write_text(json.dumps(texts, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    tags = Counter(t for a in ayat for t in re.findall(r"<(\d+)>", a["tajweedText"]))
    return tags


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    cards, works = build_surahs()
    (OUT / "surahs.json").write_text(json.dumps(cards, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    tags = build_tajweed()
    shutil.copy(ROOT / "tools/tajweed/rules.json", OUT / "tajweed-rules.json")

    sources = json.loads((ROOT / "data/sources.json").read_text(encoding="utf-8"))
    sources["cited_in_cards"] = [{"work": w, "count": n} for w, n in works.most_common()]
    (OUT / "sources.json").write_text(json.dumps(sources, ensure_ascii=False, indent=1), encoding="utf-8")

    words = len(json.loads((ROOT / "data/words/word_statistics.json").read_text(encoding="utf-8")))
    stats = {
        "surahs": len(cards),
        "ayat": sum(c["ayat"] for c in cards),
        "words": words,
        "letters": sum(c["letters"] for c in cards),
        "makki": sum(1 for c in cards if c["type"] == "مكية"),
        "madani": sum(1 for c in cards if c["type"] == "مدنية"),
        "riwayat_available": 6,
        "riwayat_planned": 8,
        "tajweed_rules": len(tags),
        "tajweed_marks": sum(tags.values()),
        "cited_works": len(works),
        "citations": sum(works.values()),
    }
    (OUT / "stats.json").write_text(json.dumps(stats, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps(stats, ensure_ascii=False))


if __name__ == "__main__":
    main()
