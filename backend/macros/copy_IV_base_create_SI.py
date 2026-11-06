import io
import os
import re
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed

from openpyxl import load_workbook
from python_calamine import CalamineWorkbook

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
MAX_COLS = 180
MAX_WORKERS = 8
USE_PROCESSES = False   # True = real parallelism (openpyxl is CPU-bound), see notes below

# Dynamic path resolution (works for any Windows user, as long as BLC ALL MACRO is in Downloads)
USER_HOME = os.path.expanduser("~")
BLC_ROOT = os.path.join(USER_HOME, "Downloads", "BLC ALL MACRO")

TEMPLATE_DIR = os.path.join(BLC_ROOT, "BL Creation", "Template")
OUTPUT_MULTI_DIR = os.path.join(BLC_ROOT, "WIP BL", "FOR RPA", "MULTI")
OUTPUT_SINGLE_DIR = os.path.join(BLC_ROOT, "WIP BL", "FOR RPA", "SINGLE")
SINGLE_TEMPLATE_NAME = "SINGLE SI TEMPLATE.xlsm"
MULTI_TEMPLATE_NAME = "MULTI SI TEMPLATE.xlsm"
TEMPLATE_SHEET = "BASE"

_EXCEL_FILE = re.compile(r"\.(xlsx|xlsm|xlsb|xls)(\.\d+)?$", re.IGNORECASE)
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")   # characters openpyxl rejects

_TEMPLATES = {}   # {"single": (bytes, insert_mode), "multi": (bytes, insert_mode)}


# --------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------
def _init_worker(templates):
    global _TEMPLATES
    _TEMPLATES = templates


def _blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def _to_count(v):
    try:
        n = int(float(v))
    except (TypeError, ValueError, OverflowError):
        return 0
    return n if n > 0 else 0


def _match_sheet(names, wanted):
    wanted = wanted.strip().upper()
    return next((s for s in names if s.strip().upper() == wanted), None)


def _resolve_path(path):
    """Existing path, or the same path without a trailing '.<digits>' suffix."""
    if os.path.exists(path):
        return path
    stripped = re.sub(r"\.\d+$", "", path)
    return stripped if os.path.exists(stripped) else None


def _collect_files(folder_or_files):
    items = [folder_or_files] if isinstance(folder_or_files, (str, os.PathLike)) else list(folder_or_files or [])
    files = []
    for item in items:
        item = os.fspath(item)
        if os.path.isdir(item):
            for f in sorted(os.listdir(item)):
                full = os.path.join(item, f)
                if os.path.isfile(full) and not f.startswith("~$") and _EXCEL_FILE.search(f):
                    files.append(full)
        else:
            files.append(item)           # explicitly chosen files are always processed
    return list(dict.fromkeys(files))    # de-duplicate, keep order


def _inside(path, folder):
    """True if path is inside folder."""
    try:
        folder = os.path.abspath(folder)
        return os.path.commonpath([os.path.abspath(path), folder]) == folder
    except ValueError:          # different drives
        return False


def _clear_output_folders(folders):
    """Delete the old Excel files in the output folders. -> (deleted_count, error_message_or_None)"""
    targets, open_names = [], []
    for folder in folders:
        if not os.path.isdir(folder):
            continue
        for name in os.listdir(folder):
            full = os.path.join(folder, name)
            if not os.path.isfile(full):
                continue
            if name.startswith("~$"):
                open_names.append(name[2:])          # Excel lock file: that workbook is open
            elif _EXCEL_FILE.search(name):
                targets.append(full)

    if open_names:
        shown = ", ".join(open_names[:5]) + (" ..." if len(open_names) > 5 else "")
        return 0, f"Close these files in Excel first, nothing was deleted: {shown}"

    deleted, failed = 0, []
    for path in targets:
        try:
            os.remove(path)
            deleted += 1
        except OSError:
            failed.append(os.path.basename(path))
    if failed:
        shown = ", ".join(failed[:5]) + (" ..." if len(failed) > 5 else "")
        return deleted, f"Could not delete {len(failed)} old file(s), are they open in Excel? {shown}"
    return deleted, None


def _template_has_rows_below_header(path):
    """True if the template's BASE sheet has any data below row 1 (the old VBA inserted above it)."""
    try:
        book = CalamineWorkbook.from_path(path)
        name = _match_sheet(book.sheet_names, TEMPLATE_SHEET)
        if name is None:
            return False
        rows = book.get_sheet_by_name(name).to_python(skip_empty_area=False)
        return any(not _blank(v) for row in rows[1:] for v in row)
    except Exception:
        return False


def _read_source(path, sheet_name):
    """-> (rows, unit_count, warning_or_None)"""
    book = CalamineWorkbook.from_path(path)
    name = _match_sheet(book.sheet_names, sheet_name)
    if name is None:
        raise ValueError(f"sheet '{sheet_name}' not found (sheets: {book.sheet_names})")

    data = book.get_sheet_by_name(name).to_python(skip_empty_area=False)
    if not data:
        raise ValueError("sheet is empty")

    b1 = _to_count(data[0][1]) if len(data[0]) > 1 else 0                         # same cell the VBA uses
    d_units = sum(1 for r in data[1:] if len(r) > 3 and not _blank(r[3]))         # column D = chassis no
    count = b1 or d_units
    if count < 1:
        raise ValueError("no units found (B1 is empty and column D has no data)")

    warning = None
    if b1 and d_units and b1 != d_units:
        warning = f"B1 says {b1} unit(s) but column D has {d_units}"

    return data[1:count + 1], count, warning


def _write_rows(ws, rows, start_row=2):
    cell = ws.cell
    for r, row in enumerate(rows, start=start_row):
        for c, v in enumerate(row[:MAX_COLS], start=1):
            if v is None or v == "":
                continue                                  # no cell object for blanks
            if isinstance(v, str):
                v = _ILLEGAL.sub("", v)
                if not v:
                    continue
            cell(row=r, column=c, value=v)


# --------------------------------------------------------------------------
# WORKER: one source file -> one SI file
# --------------------------------------------------------------------------
def _process_one(path, sheet_name):
    name = os.path.basename(path)
    try:
        resolved = _resolve_path(path)
        if resolved is None:
            return {"status": "error", "message": f"File not found: {name}"}

        rows, count, warning = _read_source(resolved, sheet_name)

        kind = "single" if count == 1 else "multi"          # VBA: If Range("B1").Value = 1
        template_bytes, insert_mode = _TEMPLATES[kind]
        out_dir = OUTPUT_SINGLE_DIR if kind == "single" else OUTPUT_MULTI_DIR

        src_name = os.path.basename(resolved)
        out_name = src_name if src_name.lower().endswith(".xlsm") else f"{src_name}.xlsm"

        wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
        ws_name = _match_sheet(wb.sheetnames, TEMPLATE_SHEET)
        if ws_name is None:
            raise ValueError(f"template has no '{TEMPLATE_SHEET}' sheet")
        ws = wb[ws_name]

        if insert_mode:
            ws.insert_rows(2, amount=count)                 # like the VBA: Range("2:2").Insert
        _write_rows(ws, rows, start_row=2)

        wb.save(os.path.join(out_dir, out_name))
        wb.close()
        return {"status": "ok", "kind": kind, "file": out_name,
                "warning": f"{name}: {warning}" if warning else None}
    except Exception as exc:
        return {"status": "error", "message": f"{name}: {exc}"}


# --------------------------------------------------------------------------
# ENTRY POINT (called by Api.run_macro)
# --------------------------------------------------------------------------
def process_si_creation(folder_or_files, sheet_name="BASE"):
    """
    folder_or_files: folder path (e.g. SEPARATED IV), file path, or list of paths.
    Creates one SI .xlsm per source file: SINGLE template if B1 = 1, otherwise MULTI.
    Old Excel files in the SINGLE and MULTI output folders are deleted just before writing.
    """
    files = _collect_files(folder_or_files)
    if not files:
        return {"ok": False, "message": "No valid files found to process."}

    single_path = os.path.join(TEMPLATE_DIR, SINGLE_TEMPLATE_NAME)
    multi_path = os.path.join(TEMPLATE_DIR, MULTI_TEMPLATE_NAME)
    for p in (single_path, multi_path):
        if not os.path.exists(p):
            return {"ok": False, "message": f"Template missing: {p}"}

    os.makedirs(OUTPUT_SINGLE_DIR, exist_ok=True)
    os.makedirs(OUTPUT_MULTI_DIR, exist_ok=True)

    templates = {}
    notes = []
    for kind, p in (("single", single_path), ("multi", multi_path)):
        with open(p, "rb") as fh:
            insert_mode = _template_has_rows_below_header(p)
            templates[kind] = (fh.read(), insert_mode)
            if insert_mode:
                notes.append(f"{os.path.basename(p)}: BASE has data below row 1, rows were inserted above it "
                             f"(formulas in that data are not re-pointed)")

    # the input must not live in the folders we are about to empty
    for f in files:
        if _inside(f, OUTPUT_SINGLE_DIR) or _inside(f, OUTPUT_MULTI_DIR):
            return {"ok": False, "message": f"{os.path.basename(f)} is inside the SINGLE or MULTI output folder. "
                                            f"Select the SEPARATED IV folder instead."}
    if not any(_resolve_path(f) for f in files):
        return {"ok": False, "message": "None of the selected files could be found."}

    # delete the old output so the folders only hold fresh files
    deleted, problem = _clear_output_folders([OUTPUT_SINGLE_DIR, OUTPUT_MULTI_DIR])
    if problem:
        return {"ok": False, "message": problem}

    workers = max(1, min(MAX_WORKERS, os.cpu_count() or 4, len(files)))
    Executor = ProcessPoolExecutor if USE_PROCESSES else ThreadPoolExecutor

    single_count = multi_count = 0
    errors, warnings = [], []

    with Executor(max_workers=workers, initializer=_init_worker, initargs=(templates,)) as pool:
        futures = {pool.submit(_process_one, f, sheet_name): f for f in files}
        for fut in as_completed(futures):
            try:
                res = fut.result()
            except Exception as exc:
                res = {"status": "error", "message": f"{os.path.basename(futures[fut])}: {exc}"}

            if res["status"] == "ok":
                if res["kind"] == "single":
                    single_count += 1
                else:
                    multi_count += 1
                if res.get("warning"):
                    warnings.append(res["warning"])
            else:
                errors.append(res["message"])

    created = single_count + multi_count
    msg = f"Stage 2 Complete: Created {single_count} SINGLE SI file(s) and {multi_count} MULTI SI file(s)."
    if deleted:
        msg += f" {deleted} old file(s) deleted first."
    if warnings:
        msg += f" {len(warnings)} warning(s)."
    if errors:
        msg += f" {len(errors)} error(s)."

    return {
        "ok": created > 0,
        "message": msg,
        "single": single_count,
        "multi": multi_count,
        "deleted_old": deleted,
        "errors": errors or None,
        "warnings": (warnings + notes) or None,
    }