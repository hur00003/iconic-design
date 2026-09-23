"""Cross-reference the HSK vocabulary workbook with the HMM grid.

Adds per-character pinyin, initial, final and tone to the vocabulary workbook,
copies the HMM actors and locations into it, and looks up each character's
actor, location and tone area with formulas, so filling in the grid later
updates every character automatically.
"""
import re
import sys
import unicodedata
from collections import Counter, defaultdict

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from pypinyin import Style, pinyin

VOCAB_IN, HMM_IN, OUT = sys.argv[1:4]

TONE_MARKS = {"āēīōūǖ": 1, "áéíóúǘ": 2, "ǎěǐǒǔǚ": 3, "àèìòùǜ": 4}
INITIALS = ["zh", "ch", "sh", "b", "p", "m", "f", "d", "t", "n", "l", "g", "k",
            "h", "j", "q", "x", "r", "z", "c", "s", "y", "w"]
# Finals whose grid label differs from the plain spelling.
FINAL_LABELS = {"EN": "(E)N", "ENG": "(E)NG", "UAN": "UAN (14th, Jacob-specific)"}
# Readings that 一 and 不 take before other syllables (tone sandhi); cards use the citation tone.
CITATION = {"一": "yī", "不": "bù"}
# Grammar particles whose everyday reading is the neutral tone.
NEUTRAL_BASE = {"的": "de", "了": "le", "得": "de", "着": "zhe", "么": "me", "们": "men",
                "吗": "ma", "呢": "ne", "吧": "ba", "啊": "a"}

HEADER_FONT = Font(name="Aptos", size=11, bold=True, color="FFFFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="FF17365D")
BODY_FONT = Font(name="Aptos", size=11, color="FF172B4D")
INPUT_FILL = PatternFill("solid", fgColor="FFFFF2CC")
NOT_ASSIGNED = "(not assigned)"


def strip_tone(s):
    out = []
    for ch in unicodedata.normalize("NFC", s):
        for marks in TONE_MARKS:
            i = marks.find(ch)
            if i >= 0:
                ch = "aeiouü"[i]
                break
        out.append(ch)
    return "".join(out)


def tone_of(syl):
    for marks, n in TONE_MARKS.items():
        if any(m in syl for m in marks):
            return n
    return 5


def is_hanzi(c):
    return "一" <= c <= "鿿"


def split_pinyin(word, given):
    """Align the list's own pinyin to each character, falling back to pypinyin."""
    hz = [c for c in word if is_hanzi(c)]
    s = re.sub(r"[\s'’·\-]", "", unicodedata.normalize("NFC", given.lower()))
    plain = strip_tone(s)
    out, pos = [], 0
    for c in hz:
        cands = {strip_tone(p) for p in pinyin(c, style=Style.TONE, heteronym=True)[0]}
        if c == "儿" and pos < len(plain) and plain[pos] == "r" and (pos + 1 == len(plain)):
            cands.add("r")
        cands = sorted(cands, key=len, reverse=True)
        hit = next((p for p in cands if plain.startswith(p, pos)), None)
        if hit is None:
            return [p[0] for p in pinyin(hz, style=Style.TONE)], False
        out.append(s[pos:pos + len(hit)])
        pos += len(hit)
    if pos != len(plain):
        return [p[0] for p in pinyin(hz, style=Style.TONE)], False
    return out, True


def parse(syl):
    """Split a syllable into (initial symbol, final label, tone) the way the HMM grid spells them."""
    tone = tone_of(syl)
    base = strip_tone(syl)
    ini = next((i for i in INITIALS if base.startswith(i) and len(base) > len(i)), "")
    fin = base[len(ini):]
    if ini in ("zh", "ch", "sh", "r", "z", "c", "s") and fin == "i":
        label = "Ø (no final)"  # zhi/chi/shi/ri/zi/ci/si: the "empty" final
    else:
        label = fin.upper().replace("Ü", "Ü")
        label = FINAL_LABELS.get(label, label)
    return (ini.upper() or "Ø"), label, tone


# ---------- read the HMM grid ----------
hmm = openpyxl.load_workbook(HMM_IN)
locations = [r for r in hmm["Locations"].iter_rows(min_row=2, values_only=True) if r[0]]
actors = [r for r in hmm["Names"].iter_rows(min_row=2, values_only=True) if r[0]]

# ---------- per-character readings from the vocabulary list ----------
wb = openpyxl.load_workbook(VOCAB_IN)
voc = wb["Vocabulary"]
readings = defaultdict(list)       # hanzi -> readings in list order
standalone = {}                    # hanzi -> reading when the character is a vocabulary entry by itself
examples = defaultdict(list)       # hanzi -> words it appears in
word_rows = []
unaligned = 0
for row in range(2, voc.max_row + 1):
    word, given = voc.cell(row, 2).value, voc.cell(row, 3).value
    if not word or not given:
        continue
    sylls, ok = split_pinyin(word, given)
    if not ok:
        unaligned += 1
        print('pinyin fallback:', word, given, sylls)
    hz = [c for c in word if is_hanzi(c)]
    word_rows.append((row, hz, sylls))
    if len(hz) == 1 and tone_of(sylls[0]) != 5:
        standalone[hz[0]] = sylls[0]
    for c, s in zip(hz, sylls):
        if s != "r":
            readings[c].append(s)
        if word not in examples[c]:
            examples[c].append(word)


def base_reading(c):
    if c in CITATION:
        return CITATION[c]
    if c in NEUTRAL_BASE:
        return NEUTRAL_BASE[c]
    if c in standalone:
        return standalone[c]
    toned = [s for s in readings[c] if tone_of(s) != 5]
    if toned:
        counts = Counter(toned)
        return max(toned, key=lambda s: (counts[s], -toned.index(s)))
    return pinyin(c, style=Style.TONE)[0][0]


# ---------- HMM Actors / HMM Locations sheets (editable grid) ----------
def header(ws, cols, widths):
    for i, (name, w) in enumerate(zip(cols, widths), 1):
        cell = ws.cell(1, i, name)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        ws.column_dimensions[cell.column_letter].width = w
    ws.freeze_panes = "A2"


for name in ("HMM Actors", "HMM Locations"):
    if name in wb.sheetnames:
        del wb[name]

wa = wb.create_sheet("HMM Actors")
header(wa, ["Initial", "Actor", "Notes"], [10, 30, 60])
actor_rows = [(s, a, "") for s, a, _ in actors]
actor_rows.append(("Ø", None, "Actor for syllables with no initial (ài, ān, ér, ō …). Fill in."))
for s, a, note in actor_rows:
    if s == "XI":
        note = "Not used by this cross-reference: xi- syllables go to X (Cher). Remove or repurpose."
    wa.append([s, a, note])
for r in range(2, wa.max_row + 1):
    for c in range(1, 4):
        wa.cell(r, c).font = BODY_FONT
    if not wa.cell(r, 2).value:
        wa.cell(r, 2).fill = INPUT_FILL

wl = wb.create_sheet("HMM Locations")
header(wl, ["Final", "Set (Location)", "Tone 1 area", "Tone 2 area", "Tone 3 area",
            "Tone 4 area", "Neutral tone area"], [28, 30, 32, 32, 32, 32, 26])
known = set()
for r in locations:
    wl.append(list(r[:6]) + [None])
    known.add(r[0])

needed = Counter()
chars = []
uh = wb["Unique Hanzi"]
for row in range(2, uh.max_row + 1):
    c = uh.cell(row, 1).value
    if c:
        chars.append((row, c))
        needed[parse(base_reading(c))[1]] += 1
for fin, _ in sorted(needed.items(), key=lambda kv: -kv[1]):
    if fin not in known:
        wl.append([fin] + [None] * 6)
for r in range(2, wl.max_row + 1):
    for c in range(1, 8):
        cell = wl.cell(r, c)
        cell.font = BODY_FONT
        if c > 1 and not cell.value:
            cell.fill = INPUT_FILL
wl.cell(1, 9, "Yellow cells are empty grid slots. Fill them in and the Unique Hanzi sheet updates automatically.").font = Font(name="Aptos", size=11, italic=True, color="FF172B4D")

# ---------- Unique Hanzi: HMM cross-reference ----------
note_title, note_body = uh["D1"].value, uh["D2"].value
uh["D1"], uh["D2"] = None, None
cols = ["Hanzi", "First HSK level", "Pinyin", "Tone", "Initial", "Final", "Actor", "Location",
        "Tone area", "HMM status", "Other readings in list", "Appears in"]
header(uh, cols, [9, 16, 10, 7, 9, 26, 24, 28, 34, 22, 20, 40])
for i in range(1, 3):
    uh.cell(1, i).alignment = Alignment(horizontal="left")

A, L = "'HMM Actors'", "'HMM Locations'"
for row, c in chars:
    py = base_reading(c)
    ini, fin, tone = parse(py)
    others = sorted({s for s in readings[c] if s != py and tone_of(s) != 5} - {py})
    vals = {
        3: py, 4: tone, 5: ini, 6: fin,
        7: f'=IFERROR(IF(INDEX({A}!$B:$B,MATCH(E{row},{A}!$A:$A,0))="","{NOT_ASSIGNED}",'
           f'INDEX({A}!$B:$B,MATCH(E{row},{A}!$A:$A,0))),"{NOT_ASSIGNED}")',
        8: f'=IFERROR(IF(INDEX({L}!$B:$B,MATCH(F{row},{L}!$A:$A,0))="","{NOT_ASSIGNED}",'
           f'INDEX({L}!$B:$B,MATCH(F{row},{L}!$A:$A,0))),"{NOT_ASSIGNED}")',
        9: f'=IFERROR(IF(INDEX({L}!$C:$G,MATCH(F{row},{L}!$A:$A,0),D{row})="","{NOT_ASSIGNED}",'
           f'INDEX({L}!$C:$G,MATCH(F{row},{L}!$A:$A,0),D{row})),"{NOT_ASSIGNED}")',
        10: f'=IF(COUNTIF(G{row}:I{row},"{NOT_ASSIGNED}")=0,"Ready","Needs grid entry")',
        11: ", ".join(others),
        12: "、".join(examples[c][:4]),
    }
    for col, val in vals.items():
        cell = uh.cell(row, col, val)
        cell.font = BODY_FONT
uh.auto_filter.ref = f"A1:L{uh.max_row}"
uh["N1"] = note_title
uh["N1"].font = Font(name="Aptos", size=15, bold=True, color="FF17365D")
uh["N2"] = note_body
uh["N3"] = ("Pinyin is each character's main reading, taken from the words in the Vocabulary sheet "
            "(一 and 不 use their citation tones). Tone 5 = neutral. Initial/Final follow the HMM grid's "
            "spelling rules: y- and w- count as initials, and zhi/chi/shi/ri/zi/ci/si use the Ø final.")
uh["N4"] = "Actor, Location and Tone area are formulas that read the HMM Actors and HMM Locations sheets."
for c in ("N2", "N3", "N4"):
    uh[c].font = BODY_FONT
    uh[c].alignment = Alignment(wrap_text=True, vertical="top")
uh.column_dimensions["N"].width = 70

# ---------- Vocabulary: fill the Character / Tone columns ----------
voc["E1"], voc["F1"] = "Character", "Tone"
for c in ("E1", "F1"):
    voc[c].font, voc[c].fill = HEADER_FONT, HEADER_FILL
for row, hz, sylls in word_rows:
    keys, tones = [], []
    for c, s in zip(hz, sylls):
        if s == "r":
            continue
        ini, fin, tone = parse(s)
        keys.append(f"{c} {s} ({ini}+{fin})")
        tones.append(str(tone))
    voc.cell(row, 5, " · ".join(keys)).font = BODY_FONT
    voc.cell(row, 6, "-".join(tones)).font = BODY_FONT
voc.column_dimensions["E"].width = 44
voc.column_dimensions["F"].width = 9
voc["E1"].comment = Comment("Each character's reading in this word, with its HMM initial + final.", "Claude")

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print(f"characters: {len(chars)}, words: {len(word_rows)}, pinyin fallback words: {unaligned}")
print("finals needing locations:", {f: n for f, n in needed.items() if f not in known})
