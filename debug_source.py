import sys
from pathlib import Path

from openpyxl.utils import get_column_letter
from python_calamine import CalamineWorkbook

PATH = sys.argv[1] if len(sys.argv) > 1 else input("Full path of the source .xlsm: ").strip().strip('"')
if not Path(PATH).is_file():
    sys.exit(f"File not found: {PATH}")
TARGET = float(input("Expected number of units for this BL (e.g. 49): ").strip())

book = CalamineWorkbook.from_path(PATH)
sheets = {n: book.get_sheet_by_name(n).to_python(skip_empty_area=False) for n in book.sheet_names}


def addr(r, c):
    return f"{get_column_letter(c + 1)}{r + 1}"


def row_preview(row, limit=14):
    return [(get_column_letter(i + 1), v) for i, v in enumerate(row) if v not in ("", None)][:limit]


print("\n=== SHEETS ===")
for n, rows in sheets.items():
    print(f"{n!r}: {len(rows)} rows x {max((len(r) for r in rows), default=0)} cols")

print(f"\n=== CELLS EQUAL TO {TARGET:g} (up to 40) ===")
hits = 0
for n, rows in sheets.items():
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            if isinstance(v, (int, float)) and v == TARGET:
                print(f"  {n}!{addr(r, c)} = {v}")
                hits += 1
                if hits >= 40:
                    break
        if hits >= 40:
            break
if hits == 0:
    print("  none found")

print("\n=== ATCH: first 6 non-empty rows ===")
atch = sheets.get("ATCH", [])
shown = 0
for r, row in enumerate(atch):
    p = row_preview(row)
    if p:
        print(f"  row {r + 1}: {p}")
        shown += 1
        if shown >= 6:
            break

print("\n=== ATCH: last 5 non-empty rows ===")
last = [(r + 1, row_preview(row)) for r, row in enumerate(atch) if row_preview(row)]
for r, p in last[-5:]:
    print(f"  row {r}: {p}")

print("\n=== BASE: header row (row 1) ===")
base = sheets.get("BASE", [])
if base:
    print(" ", row_preview(base[0], limit=120))

bl = sheets.get("SI", [[]] * 4)[3][4] if len(sheets.get("SI", [])) > 3 and len(sheets["SI"][3]) > 4 else ""
print(f"\n=== BASE rows containing the BL no {bl!r} ===")
matches = [r + 1 for r, row in enumerate(base) if bl and any(v == bl for v in row)]
print(f"  {len(matches)} row(s): {matches[:60]}")