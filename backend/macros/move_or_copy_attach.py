import os
import threading
from concurrent.futures import ThreadPoolExecutor
from copy import copy
from pathlib import Path

from openpyxl import load_workbook

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
# Dynamic path resolution (works for any Windows user, as long as BLC ALL MACRO is in Downloads)
BL_ROOT = Path(os.path.expanduser("~")) / "Downloads" / "BLC ALL MACRO"
TARGET_DIR = BL_ROOT / "WIP BL" / "BL" / "ALL"     # the BL files that receive the attach sheet

SHEET_NAME = "ATCH (2)"      # the sheet that is copied from each source file
BL_CELL = "G2"               # cell of that sheet that holds the BL no (used to find the matching BL file)
IF_EXISTS = "skip"           # when the BL file already has the sheet: "skip" it or "replace" it
MAX_WORKERS = 8

_locks, _locks_guard = {}, threading.Lock()


# --------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------
def _lock_for(path):
    """One lock per BL file, so two sources that point to the same BL file never write at once."""
    key = os.path.normcase(str(path))
    with _locks_guard:
        return _locks.setdefault(key, threading.Lock())


def _find_name(names, wanted):
    wanted = wanted.strip().lower()
    return next((n for n in names if n.strip().lower() == wanted), None)


def _bl_text(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return "" if v is None else str(v).strip()


def find_targets(bl_no, target_dir):
    """BL files in target_dir for this BL no. Names that start with the BL no win over a loose 'contains'."""
    files = sorted(p for p in Path(target_dir).glob("*.xlsm") if not p.name.startswith("~$"))
    low = bl_no.lower()
    exact = [p for p in files if p.name.lower().startswith(low + " ") or p.stem.lower() == low]
    return exact or [p for p in files if low in p.name.lower()]


def copy_sheet(src, dst_wb, title):
    """
    Copy a worksheet into ANOTHER workbook: values, cell formatting, column widths, row heights,
    merged cells and the page / print setup. Returns (new_sheet, warnings).
    """
    dst = dst_wb.create_sheet(title)
    warnings = []

    for row in src.iter_rows():
        for c in row:
            if c.value is None and not c.has_style:
                continue
            d = dst.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                d.font = copy(c.font)
                d.border = copy(c.border)
                d.fill = copy(c.fill)
                d.number_format = c.number_format
                d.protection = copy(c.protection)
                d.alignment = copy(c.alignment)

    for key, dim in src.column_dimensions.items():
        nd = dst.column_dimensions[key]
        nd.min, nd.max = dim.min, dim.max
        nd.width, nd.hidden, nd.outlineLevel, nd.bestFit = dim.width, dim.hidden, dim.outlineLevel, dim.bestFit
    for key, dim in src.row_dimensions.items():
        nd = dst.row_dimensions[key]
        nd.height, nd.hidden, nd.outlineLevel = dim.height, dim.hidden, dim.outlineLevel

    dst.sheet_format = copy(src.sheet_format)
    dst.sheet_properties = copy(src.sheet_properties)
    dst.page_margins = copy(src.page_margins)
    dst.page_setup = copy(src.page_setup)
    dst.print_options = copy(src.print_options)
    dst.sheet_state = src.sheet_state
    try:
        dst.HeaderFooter = copy(src.HeaderFooter)
        dst.print_title_rows = src.print_title_rows
        if src.print_area:
            dst.print_area = str(src.print_area).split("!")[-1].replace("$", "")
        for attr in ("view", "showGridLines", "zoomScale", "zoomScaleNormal", "zoomScaleSheetLayoutView"):
            setattr(dst.sheet_view, attr, getattr(src.sheet_view, attr))
    except Exception:
        warnings.append("some print or view settings could not be copied")

    for rng in list(src.merged_cells.ranges):                  # merge last, so no cell write hits a merged cell
        dst.merge_cells(str(rng))

    if getattr(src, "_images", None):
        warnings.append("the sheet has pictures, which were not copied")
    try:
        if len(list(src.conditional_formatting)):
            warnings.append("the sheet has conditional formatting, which was not copied")
    except Exception:
        pass
    return dst, warnings


# --------------------------------------------------------------------------
# WORKER: one source file -> its BL file
# --------------------------------------------------------------------------
def attach_one(src_path, target_dir):
    """-> (status, message, warnings)  status: 'added' | 'skipped' | 'error'"""
    name = src_path.name
    try:
        swb = load_workbook(src_path, data_only=True)           # saved values, as Excel shows them
    except Exception as exc:
        return "error", f"{name}: could not be read ({exc})", []

    try:
        sname = _find_name(swb.sheetnames, SHEET_NAME)
        if sname is None:
            return "error", f"{name}: no '{SHEET_NAME}' sheet (only multi-unit files have one)", []
        src = swb[sname]

        bl_no = _bl_text(src[BL_CELL].value)
        if not bl_no:
            return "error", f"{name}: {BL_CELL} (BL no) is empty on '{SHEET_NAME}', was Stage 3 run?", []

        matches = find_targets(bl_no, target_dir)
        if not matches:
            return "error", f"{name}: no BL file containing '{bl_no}' in {Path(target_dir).name}", []
        if len(matches) > 1:
            shown = ", ".join(p.name for p in matches[:3])
            return "error", f"{name}: more than one BL file matches '{bl_no}' ({shown}), nothing was changed", []
        target = matches[0]

        with _lock_for(target):
            twb = load_workbook(target, keep_vba=True)           # the BL file keeps its macros
            existing = _find_name(twb.sheetnames, SHEET_NAME)
            if existing is not None:
                if IF_EXISTS != "replace":
                    twb.close()
                    return "skipped", f"{target.name}: already has '{SHEET_NAME}'", []
                del twb[existing]

            _, warnings = copy_sheet(src, twb, SHEET_NAME)

            for i, ws in enumerate(twb.worksheets):              # like the macro: finish on sheet 1, cell A1
                ws.sheet_view.tabSelected = (i == 0)
            twb.active = 0

            tmp = target.with_name(target.name + ".tmp")
            try:
                twb.save(tmp)
                twb.close()
                os.replace(tmp, target)
            except PermissionError:
                return "error", f"{target.name}: could not be saved, is it open in Excel? Close it and run again", []
            finally:
                if tmp.exists():
                    try:
                        tmp.unlink()
                    except OSError:
                        pass
        return "added", f"{name} -> {target.name}", [f"{target.name}: {w}" for w in warnings]
    except Exception as exc:
        return "error", f"{name}: {exc}", []
    finally:
        swb.close()


# --------------------------------------------------------------------------
# ENTRY POINT (called by Api.run_macro)
# --------------------------------------------------------------------------
def _collect_sources(folder_or_files):
    items = [folder_or_files] if isinstance(folder_or_files, (str, os.PathLike)) else list(folder_or_files or [])
    files, problems = [], []
    for item in items:
        p = Path(os.fspath(item))
        if p.is_dir():
            multi = next((s for s in p.iterdir() if s.is_dir() and s.name.lower() == "multi"), None)
            base = multi or p                               # FOR RPA selected -> use its MULTI folder
            files += [f for f in sorted(base.glob("*.xlsm")) if not f.name.startswith("~$")]
        elif p.is_file() and p.suffix.lower() == ".xlsm":
            files.append(p)
        else:
            problems.append(f"{p.name}: not an .xlsm file or folder" if p.exists() else f"Not found: {p}")
    return list(dict.fromkeys(files)), problems


def process_attach(folder_or_files, target_folder=None):
    """
    For every multi-unit file in the selected folder (FOR RPA\\MULTI), copies its "ATCH (2)" sheet to the end
    of the matching BL file in BL\\ALL. The match is made with the BL no found in cell G2 of that sheet.
    """
    try:
        if not BL_ROOT.is_dir():
            return {"ok": False, "message": f"The folder 'BLC ALL MACRO' was not found in Downloads: {BL_ROOT}"}

        target_dir = Path(target_folder) if target_folder else TARGET_DIR
        if not target_dir.is_dir():
            return {"ok": False, "message": f"The BL folder was not found: {target_dir}. Run Stage 4 first."}

        files, problems = _collect_sources(folder_or_files)
        if not files:
            return {"ok": False, "message": "No .xlsm files found. Select the FOR RPA\\MULTI folder.",
                    "errors": problems or None}

        added, skipped, errors, warnings = 0, [], list(problems), []
        with ThreadPoolExecutor(max_workers=max(1, min(MAX_WORKERS, len(files)))) as pool:
            for status, message, warn in pool.map(lambda f: attach_one(f, target_dir), files):
                warnings.extend(warn)
                if status == "added":
                    added += 1
                elif status == "skipped":
                    skipped.append(message)
                else:
                    errors.append(message)

        text = f"Stage 7: '{SHEET_NAME}' added to {added} BL file(s)."
        if skipped:
            text += f" {len(skipped)} already had it."
        if warnings:
            text += f" {len(warnings)} warning(s)."
        if errors:
            text += f" {len(errors)} error(s)."
        return {
            "ok": added + len(skipped) > 0,
            "message": text,
            "added": added,
            "skipped": skipped or None,
            "warnings": warnings or None,
            "errors": errors or None,
            "target_dir": str(target_dir),
        }
    except Exception as exc:
        return {"ok": False, "message": f"Attach sheet failed: {exc!r}"}