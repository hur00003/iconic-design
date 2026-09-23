"""Render a batch of HMM movie cards (batch-NN.json) to an Anki import file and a readable Markdown copy.

Usage: python make_batch.py batch-02.json
Each card: [hanzi, pinyin, tone, meaning, actor, location, area, props, movie (HTML <b> allowed), words]
"""
import json
import re
import sys
from pathlib import Path

TONES = {1: "Tone 1 (flat)", 2: "Tone 2 (rising)", 3: "Tone 3 (dipping)", 4: "Tone 4 (falling)", 5: "Neutral"}

src = Path(sys.argv[1])
batch = src.stem  # e.g. batch-02
cards = json.loads(src.read_text(encoding="utf-8"))

lines = ["#separator:tab", "#html:true", "#tags column:3", "#notetype:Basic", "#deck:HMM Hanzi"]
md = [f"# HMM Movies — {batch.replace('-', ' ').title()} ({len(cards)} characters)", "", "Props follow `props.md`.", ""]
for i, (h, py, t, meaning, actor, loc, area, props, movie, words) in enumerate(cards, 1):
    back = (f"<div style='font-size:28px'>{h} · <b>{py}</b></div>"
            f"<div>{meaning}</div><hr>"
            f"<div><b>{actor}</b> · {loc} · {area} · {TONES[t]}</div>"
            f"<div><i>Props:</i> {props}</div><br>"
            f"<div>{movie}</div><br>"
            f"<div><i>Words:</i> {words}</div>")
    front = f"<div style='font-size:96px'>{h}</div>"
    lines.append("\t".join([front, back, f"HMM {batch.replace('-', '')} tone{t}"]))
    md += [f"### {i}. {h} {py} — {meaning}", f"**{actor}** · {loc} · {area} · {TONES[t]}  ",
           f"*Props:* {props}", "", re.sub(r"</?b>", "**", movie), "", f"*Words:* {words}", ""]

src.with_suffix(".tsv").write_text("\n".join(lines) + "\n", encoding="utf-8")
src.with_suffix(".md").write_text("\n".join(md), encoding="utf-8")
print(f"{batch}: {len(cards)} cards")
