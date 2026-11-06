import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from python_calamine import CalamineWorkbook

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
TEMPLATE_NAME = "SO.UT.ATTACH.xlsm"      # lives in the project's "templates" folder
TEMPLATE_SHEET = "BASE"
MAX_COLS = 180
MAX_WORKERS = 8

# Double checks on the combined BASE data (0-based column numbers: D = 3, AM = 38, BE = 56)
COL_CHASSIS = 3          # D   chassis no
COL_IV_NO = 38           # AM  IV no
COL_SHIPPER = 56         # BE  shipper
MAX_FLAG_LINES = 20      # how many individual flags are listed (the rest is summarised)
ASK_BEFORE_SAVING = True # if flags were found, ask "Save the file anyway?" before the Save As window

_EXCEL_FILE = re.compile(r"\.(xlsx|xlsm|xlsb|xls)(\.\d+)?$", re.IGNORECASE)
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")   # characters openpyxl refuses to write


# --------------------------------------------------------------------------
# TEMPLATE LOOKUP (bundled with the project, never the one in the Downloads folder)
# --------------------------------------------------------------------------
def _template_candidates():
    here = Path(__file__).resolve().parent            # e.g. <project>/backend/macros
    roots = [here.parent / "templates",               # <project>/backend/templates
             here.parent.parent / "templates",        # <project>/templates
             here / "templates"]                      # <project>/backend/macros/templates
    if getattr(sys, "frozen", False):                 # packaged app (PyInstaller)
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        roots += [base / "templates", base / "backend" / "templates"]
    return [r / TEMPLATE_NAME for r in roots]


def find_template():
    """-> (path or None, list of paths that were tried)"""
    tried = _template_candidates()
    return next((p for p in tried if p.is_file()), None), tried


# --------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------
def _blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def _to_count(v):
    try:
        n = int(float(v))
    except (TypeError, ValueError, OverflowError):
        return 0
    return n if n > 0 else 0


def _resolve_path(path):
    """Existing path, or the same path without a trailing '.<digits>' suffix."""
    if os.path.exists(path):
        return path
    stripped = re.sub(r"\.\d+$", "", path)
    return stripped if os.path.exists(stripped) else None


def _collect_files(file_paths):
    """A single path (file or folder), or a list of them -> list of file paths, order kept."""
    if not file_paths:
        return []
    items = [file_paths] if isinstance(file_paths, (str, os.PathLike)) else list(file_paths)
    files = []
    for item in items:
        item = os.fspath(item)
        if os.path.isdir(item):
            for name in sorted(os.listdir(item)):
                full = os.path.join(item, name)
                if os.path.isfile(full) and not name.startswith("~$") and _EXCEL_FILE.search(name):
                    files.append(full)
        else:
            files.append(item)
    return list(dict.fromkeys(files))


def _with_xlsm(path):
    """Make sure the chosen file ends with .xlsm (a macro-enabled file needs it)."""
    path = re.sub(r"\.(xlsx|xls)$", "", os.fspath(path), flags=re.IGNORECASE)
    return path if path.lower().endswith(".xlsm") else path + ".xlsm"


def parse_single_file(file_path, target_sheet_name="BASE"):
    """-> (rows, error). rows = one list per unit, up to MAX_COLS values, '' for empty cells."""
    name = os.path.basename(file_path)
    target = _resolve_path(file_path)
    if target is None:
        return None, f"File not found: {name}"

    try:
        book = CalamineWorkbook.from_path(target)
        wanted = target_sheet_name.strip().upper()
        sheet = next((s for s in book.sheet_names if s.strip().upper() == wanted), None)
        if sheet is None:
            return None, f"{name}: sheet '{target_sheet_name}' not found"

        data = book.get_sheet_by_name(sheet).to_python(skip_empty_area=False)
        if len(data) < 2:
            return None, f"No rows found in {name}"

        # unit count: BASE!B1, otherwise the filled cells of column D below the header
        header = data[0]
        count = _to_count(header[1]) if len(header) > 1 else 0
        if count == 0:
            count = sum(1 for r in data[1:] if len(r) > 3 and not _blank(r[3]))
        if count == 0:
            return None, f"No units found in {name}"

        return [list(r[:MAX_COLS]) for r in data[1:count + 1]], None
    except Exception as err:
        return None, f"Error reading {name}: {err}"


# --------------------------------------------------------------------------
# DOUBLE CHECKS ON THE COMBINED DATA
# --------------------------------------------------------------------------
def _norm(v):
    """Comparison key: ignores upper/lower case and extra spaces (12.0 and 12 are the same)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", " ", str(v)).strip().upper()


def _shown(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", " ", str(v)).strip()


def _get(row, idx):
    return row[idx] if idx < len(row) else None


def _rows_text(items, first_row, limit=6):
    nums = [str(first_row + i) for i in items[:limit]]
    more = f" +{len(items) - limit} more" if len(items) > limit else ""
    return ("rows " if len(items) > 1 else "row ") + ", ".join(nums) + more


def _files_text(items, origins):
    names = list(dict.fromkeys(origins[i] for i in items))
    return ", ".join(names[:3]) + (" ..." if len(names) > 3 else "")


def check_duplicates(rows, origins, first_row=2):
    """
    Looks for two kinds of problems in the combined data:
      1. the same chassis no (column D) more than once
      2. the same IV no (column AM) used with different shippers (column BE)
    rows: the combined BASE rows. origins: the source file of every row.
    first_row: the sheet row of the first data row. -> (flags, counts)
    """
    flags = []

    chassis = {}
    for i, row in enumerate(rows):
        key = _norm(_get(row, COL_CHASSIS))
        if key:
            chassis.setdefault(key, []).append(i)
    chassis_dups = {k: v for k, v in chassis.items() if len(v) > 1}
    for items in chassis_dups.values():
        value = _shown(_get(rows[items[0]], COL_CHASSIS))
        flags.append(f"Duplicate chassis no '{value}' appears {len(items)} times "
                     f"({_rows_text(items, first_row)}; {_files_text(items, origins)})")

    ivs = {}
    for i, row in enumerate(rows):
        iv, shipper = _norm(_get(row, COL_IV_NO)), _norm(_get(row, COL_SHIPPER))
        if iv and shipper:
            entry = ivs.setdefault(iv, {"shown": _shown(_get(row, COL_IV_NO)), "by": {}})
            entry["by"].setdefault(shipper, {"shown": _shown(_get(row, COL_SHIPPER)), "rows": []})["rows"].append(i)
    iv_dups = {k: v for k, v in ivs.items() if len(v["by"]) > 1}
    for entry in iv_dups.values():
        parts = [f"'{g['shown']}' ({_rows_text(g['rows'], first_row)}; {_files_text(g['rows'], origins)})"
                 for g in entry["by"].values()]
        flags.append(f"IV no '{entry['shown']}' is used with different shippers: " + " vs ".join(parts))

    return flags, {"chassis": len(chassis_dups), "iv": len(iv_dups)}


def _limit_flags(flags):
    if len(flags) <= MAX_FLAG_LINES:
        return list(flags)
    return flags[:MAX_FLAG_LINES] + [f"...and {len(flags) - MAX_FLAG_LINES} more flag(s)"]


# --------------------------------------------------------------------------
# ENTRY POINT (called by Api.run_macro and the one-click pipeline)
# --------------------------------------------------------------------------
def process_shipping_list(file_paths, sheet_name="BASE", template_folder=None,
                          output_path=None, ask_save_path=None, confirm=None):
    """
    Combines the BASE rows of the selected files into a fresh copy of the bundled
    SO.UT.ATTACH.xlsm template (macros kept) and saves it as a NEW file.
    The template itself is never modified.

    output_path:   full path to save to. If given, no dialog is shown.
    ask_save_path: function(suggested_name, start_dir) -> chosen path, or None if cancelled.
                   It is called after the data is ready, so the user picks the location last.
    confirm:       function(title, text) -> True / False. Used to ask "save anyway?" when the double
                   checks find duplicates. Without it the flags are only returned as warnings.
    template_folder is accepted only so older callers do not break; it is ignored.
    """
    template, tried = find_template()
    if template is None:
        listing = "\n".join(f"  {p}" for p in tried)
        return {"ok": False, "message": f"Template {TEMPLATE_NAME} not found in the project. Looked in:\n{listing}"}

    files = _collect_files(file_paths)
    if not files:
        return {"ok": False, "message": "No files to process."}

    all_rows, origins, errors, processed = [], [], [], 0
    with ThreadPoolExecutor(max_workers=max(1, min(MAX_WORKERS, len(files)))) as pool:
        results = pool.map(lambda p: parse_single_file(p, sheet_name), files)       # keeps input order
        for path, (rows, err) in zip(files, results):
            if err:
                errors.append(err)
            elif rows:
                all_rows.extend(rows)
                origins.extend([os.path.basename(path)] * len(rows))
                processed += 1

    if not all_rows:
        return {"ok": False, "message": f"Processed 0 files. Details: {' | '.join(errors[:2])}",
                "errors": errors or None}

    # double checks: duplicate chassis no, and the same IV no with different shippers
    flags, counts = check_duplicates(all_rows, origins, first_row=2)
    warnings = _limit_flags(flags)
    if flags and ASK_BEFORE_SAVING and confirm is not None:
        text = ("Possible problems were found in the combined shipping list:\n\n"
                + "\n".join("- " + w for w in warnings[:8])
                + (f"\n- ...and {len(flags) - 8} more" if len(flags) > 8 else "")
                + "\n\nSave the file anyway?")
        if not confirm("Check the shipping list", text):
            return {"ok": False, "cancelled": True,
                    "message": f"Stopped: {len(flags)} duplicate flag(s) found, so no file was created. "
                            f"Fix the source files and run again.",
                    "warnings": warnings, "errors": errors or None}

    try:
        wb = load_workbook(template, keep_vba=True)        # keeps the macros inside the template
        ws_name = next((s for s in wb.sheetnames if s.strip().upper() == TEMPLATE_SHEET), None)
        if ws_name is None:
            wb.close()
            return {"ok": False, "message": f"The template has no '{TEMPLATE_SHEET}' sheet."}
        ws = wb[ws_name]

        cell = ws.cell
        for r, row in enumerate(all_rows, start=2):
            for c, v in enumerate(row, start=1):
                if v is None or v == "":
                    continue                               # no cell object for blanks
                if isinstance(v, str):
                    v = _ILLEGAL.sub("", v)
                    if not v:
                        continue
                cell(row=r, column=c, value=v)

        # the data is ready: now ask where to save it
        if not output_path:
            if ask_save_path is None:
                wb.close()
                return {"ok": False, "message": "No save location was given."}
            suggested = f"{template.stem}_{datetime.now():%Y-%m-%d_%H%M%S}{template.suffix}"
            start_dir = os.path.dirname(files[0]) if files else None
            output_path = ask_save_path(suggested, start_dir)
            if not output_path:
                wb.close()
                return {"ok": False, "cancelled": True, "message": "Save cancelled. No file was created.",
                        "warnings": warnings or None}

        out_path = Path(_with_xlsm(output_path))
        out_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(out_path)
        wb.close()
    except PermissionError as exc:
        return {"ok": False, "message": f"Could not save the new file (is it open in Excel?): {exc}"}
    except Exception as exc:
        return {"ok": False, "message": f"Template write error: {exc}"}

    message = f"Shipping list saved: {out_path.name} ({processed} file(s), {len(all_rows)} total rows)."
    if errors:
        message += f" {len(errors)} file(s) skipped."
    if flags:
        message += (f" CHECK: {counts['chassis']} duplicate chassis no, "
                    f"{counts['iv']} IV no with different shippers (see the warnings).")
    return {
        "ok": True,
        "message": message,
        "output_path": str(out_path),
        "output_dir": str(out_path.parent),
        "flags": len(flags),
        "warnings": warnings or None,
        "errors": errors or None,
    }