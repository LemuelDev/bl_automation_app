import io
import os
import re
import threading
from concurrent.futures import ThreadPoolExecutor
from copy import copy
from functools import lru_cache
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

# Layout
# Column widths are counted in digits of the workbook's default font. When the BL file and the attach sheet's
# file use different default fonts, the same number is a different width in pixels. "auto" measures both fonts
# (Windows) and converts the widths so the sheet keeps its size. Use a number (1 = off, 8 / 7 = 14% wider) to
# set it by hand.
WIDTH_SCALE = "auto"
BLANK_ZEROS = False          # True = also blank every numeric 0 (the old behaviour). False = keep the values and
                             # let the copied "show zeros" setting / number formats / conditional formats hide them

_locks, _locks_guard = {}, threading.Lock()
_notes, _notes_guard = set(), threading.Lock()


# --------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------
def _lock_for(path):
    """One lock per BL file, so two sources that point to the same BL file never write at once."""
    key = os.path.normcase(str(path))
    with _locks_guard:
        return _locks.setdefault(key, threading.Lock())


def _note_once(text):
    """True the first time a message is seen in a run (a layout note is not repeated for every file)."""
    with _notes_guard:
        if text in _notes:
            return False
        _notes.add(text)
        return True


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


# --------------------------------------------------------------------------
# DEFAULT FONT -> COLUMN WIDTH CONVERSION
# --------------------------------------------------------------------------
def _normal_font(wb):
    try:
        f = wb._named_styles["Normal"].font
    except Exception:
        f = wb._fonts[0]
    return str(f.name or "Calibri"), float(f.sz or 11)


@lru_cache(maxsize=32)
def _digit_width(name, size):
    """Width in pixels of the widest digit of a font at 96 dpi: the unit Excel counts column widths in.
    Windows only. None when it cannot be measured (also when the font is not installed)."""
    try:
        import ctypes
        from ctypes import wintypes
        gdi32, user32 = ctypes.windll.gdi32, ctypes.windll.user32
        vp, ci = ctypes.c_void_p, ctypes.c_int
        user32.GetDC.restype, user32.GetDC.argtypes = vp, [vp]
        user32.ReleaseDC.argtypes = [vp, vp]
        gdi32.CreateFontW.restype = vp
        gdi32.CreateFontW.argtypes = [ci] * 13 + [ctypes.c_wchar_p]
        gdi32.SelectObject.restype, gdi32.SelectObject.argtypes = vp, [vp, vp]
        gdi32.DeleteObject.argtypes = [vp]
        gdi32.GetTextFaceW.argtypes = [vp, ci, ctypes.c_wchar_p]
        gdi32.GetTextExtentPoint32W.argtypes = [vp, ctypes.c_wchar_p, ci, ctypes.POINTER(wintypes.SIZE)]
    except Exception:
        return None

    hdc = user32.GetDC(None)
    if not hdc:
        return None
    font = old = None
    try:
        font = gdi32.CreateFontW(-int(round(size * 96 / 72)), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, name)
        if not font:
            return None
        old = gdi32.SelectObject(hdc, font)
        face = ctypes.create_unicode_buffer(64)
        gdi32.GetTextFaceW(hdc, 64, face)
        if face.value.strip().lower() != name.strip().lower():
            return None                                   # Windows substituted another font: do not trust it
        widest, box = 0, wintypes.SIZE()
        for ch in "0123456789":
            gdi32.GetTextExtentPoint32W(hdc, ch, 1, ctypes.byref(box))
            widest = max(widest, box.cx)
        return widest or None
    except Exception:
        return None
    finally:
        try:
            if old:
                gdi32.SelectObject(hdc, old)
            if font:
                gdi32.DeleteObject(font)
        finally:
            user32.ReleaseDC(None, hdc)


def _width_scale(src_wb, dst_wb):
    """-> (factor to multiply the source column widths with, note or None)"""
    if isinstance(WIDTH_SCALE, (int, float)) and not isinstance(WIDTH_SCALE, bool):
        return float(WIDTH_SCALE), (f"column widths scaled by {WIDTH_SCALE:g} (WIDTH_SCALE)" if WIDTH_SCALE != 1 else None)

    (sname, ssize), (dname, dsize) = _normal_font(src_wb), _normal_font(dst_wb)
    if (sname.lower(), ssize) == (dname.lower(), dsize):
        return 1.0, None
    sw, dw = _digit_width(sname, ssize), _digit_width(dname, dsize)
    if sw and dw:
        if sw == dw:
            return 1.0, None
        factor = sw / dw
        return factor, (f"layout: the default fonts differ ('{sname} {ssize:g}' digit {sw}px, BL file '{dname} {dsize:g}' "
                        f"digit {dw}px), column widths were scaled by {factor:.3f}")
    return 1.0, (f"layout: the default fonts differ ('{sname} {ssize:g}' and '{dname} {dsize:g}') and could not be measured. "
                f"If the columns look narrower than in the source, set WIDTH_SCALE by hand (for example 8 / 7)")


# --------------------------------------------------------------------------
# SHEET COPY (one worksheet into ANOTHER workbook, same look)
# --------------------------------------------------------------------------
_VIEW_ATTRS = ("view", "showGridLines", "showRowColHeaders", "showFormulas", "showZeros", "showOutlineSymbols",
               "showRuler", "showWhiteSpace", "rightToLeft", "defaultGridColor", "zoomScale", "zoomScaleNormal",
               "zoomScaleSheetLayoutView", "zoomScalePageLayoutView")


def copy_sheet(src, dst_wb, title, scale=1.0):
    """Copy values, formatting, widths, heights, merged cells, page setup, print settings, view settings and
    pictures. Formulas are copied as their saved values. Returns (new sheet, warnings)."""
    dst = dst_wb.create_sheet(title)
    warnings = []

    # sheet defaults: row height and column width that rows / columns without their own size use
    dst.sheet_format = copy(src.sheet_format)
    if src.sheet_format.defaultRowHeight:
        dst.sheet_format.customHeight = True            # keep that height; do not let Excel re-fit it to the font
    if src.sheet_format.defaultColWidth:
        dst.sheet_format.defaultColWidth = round(src.sheet_format.defaultColWidth * scale, 2)

    # cells
    for row in src.iter_rows():
        for c in row:
            v = c.value
            if BLANK_ZEROS and isinstance(v, (int, float)) and not isinstance(v, bool) and v == 0:
                v = None
            if v is None and not c.has_style:
                continue
            d = dst.cell(row=c.row, column=c.column, value=v)
            if c.has_style:
                f = copy(c.font)
                if f.scheme:
                    f.scheme = None                     # a theme font would change with the BL file's theme
                d.font = f
                d.border = copy(c.border)
                d.fill = copy(c.fill)
                d.number_format = c.number_format
                d.protection = copy(c.protection)
                d.alignment = copy(c.alignment)

    # column widths and row heights
    for key, dim in src.column_dimensions.items():
        nd = dst.column_dimensions[key]
        nd.min, nd.max = dim.min, dim.max
        if dim.width is not None:
            nd.width = round(dim.width * scale, 2)
        nd.hidden, nd.outlineLevel, nd.bestFit = dim.hidden, dim.outlineLevel, dim.bestFit
    for key, dim in src.row_dimensions.items():
        nd = dst.row_dimensions[key]
        if dim.height is not None:
            nd.height = dim.height
        nd.hidden, nd.outlineLevel = dim.hidden, dim.outlineLevel

    # page setup and printing
    dst.sheet_properties = copy(src.sheet_properties)   # includes fit-to-page
    dst.sheet_properties.codeName = None                # a macro code name must be unique in the BL file
    dst.page_margins = copy(src.page_margins)
    dst.page_setup = copy(src.page_setup)
    dst.print_options = copy(src.print_options)
    dst.sheet_state = src.sheet_state
    for attr in ("row_breaks", "col_breaks", "HeaderFooter"):
        try:
            setattr(dst, attr, copy(getattr(src, attr)))
        except Exception:
            warnings.append(f"{attr} could not be copied")
    try:
        dst.print_title_rows = src.print_title_rows
        dst.print_title_cols = src.print_title_cols
        if src.print_area:
            dst.print_area = str(src.print_area).split("!")[-1].replace("$", "")
    except Exception:
        warnings.append("the print area / titles could not be copied")

    # how the sheet is shown (zoom, page break preview, "show zeros", gridlines ...)
    for attr in _VIEW_ATTRS:
        try:
            setattr(dst.sheet_view, attr, getattr(src.sheet_view, attr))
        except Exception:
            pass

    # conditional formatting
    try:
        for cf in src.conditional_formatting:
            for rule in cf.rules:
                dst.conditional_formatting.add(str(cf.sqref), copy(rule))
    except Exception:
        warnings.append("conditional formatting could not be copied")

    for rng in list(src.merged_cells.ranges):           # merge last, so no cell write hits a merged cell
        dst.merge_cells(str(rng))

    images = list(getattr(src, "_images", []) or [])
    if images:
        try:
            from openpyxl.drawing.image import Image
            for img in images:
                new = Image(io.BytesIO(img._data()))
                new.anchor = copy(img.anchor)
                dst.add_image(new)
        except Exception:
            warnings.append("pictures could not be copied")
    return dst, warnings


# --------------------------------------------------------------------------
# WORKER: one source file -> its BL file
# --------------------------------------------------------------------------
def attach_one(src_path, target_dir):
    """-> (status, message, warnings)  status: 'added' | 'skipped' | 'error'"""
    name = src_path.name
    try:
        swb = load_workbook(src_path, data_only=True)       # saved values, as Excel shows them
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
            twb = load_workbook(target, keep_vba=True)       # the BL file keeps its macros
            existing = _find_name(twb.sheetnames, SHEET_NAME)
            if existing is not None:
                if IF_EXISTS != "replace":
                    twb.close()
                    return "skipped", f"{target.name}: already has '{SHEET_NAME}'", []
                del twb[existing]

            scale, note = _width_scale(swb, twb)
            _, warnings = copy_sheet(src, twb, SHEET_NAME, scale)
            warnings = [f"{target.name}: {w}" for w in warnings]
            if note and _note_once(note):
                warnings.insert(0, note)

            for i, ws in enumerate(twb.worksheets):          # like the macro: finish on sheet 1, cell A1
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
        return "added", f"{name} -> {target.name}", warnings
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

        with _notes_guard:
            _notes.clear()

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


# import os
# import threading
# from concurrent.futures import ThreadPoolExecutor
# from copy import copy
# from pathlib import Path
# from openpyxl import load_workbook
# import re
# from python_calamine import CalamineWorkbook

# # --------------------------------------------------------------------------
# # CONFIG
# # --------------------------------------------------------------------------
# # Dynamic path resolution (works for any Windows user, as long as BLC ALL MACRO is in Downloads)
# BL_ROOT = Path(os.path.expanduser("~")) / "Downloads" / "BLC ALL MACRO"
# TARGET_DIR = BL_ROOT / "WIP BL" / "BL" / "ALL"     # the BL files that receive the attach sheet

# SHEET_NAME = "ATCH (2)"      # the sheet that is copied from each source file
# BL_CELL = "G2"               # cell of that sheet that holds the BL no (used to find the matching BL file)
# IF_EXISTS = "skip"           # when the BL file already has the sheet: "skip" it or "replace" it
# MAX_WORKERS = 8

# _locks, _locks_guard = {}, threading.Lock()


# # --------------------------------------------------------------------------
# # HELPERS
# # --------------------------------------------------------------------------
# def _lock_for(path):
#     """One lock per BL file, so two sources that point to the same BL file never write at once."""
#     key = os.path.normcase(str(path))
#     with _locks_guard:
#         return _locks.setdefault(key, threading.Lock())


# def _find_name(names, wanted):
#     wanted = wanted.strip().lower()
#     return next((n for n in names if n.strip().lower() == wanted), None)


# def _bl_text(v):
#     if isinstance(v, float) and v.is_integer():
#         v = int(v)
#     return "" if v is None else str(v).strip()


# def find_targets(bl_no, target_dir):
#     """BL files in target_dir for this BL no. Names that start with the BL no win over a loose 'contains'."""
#     files = sorted(p for p in Path(target_dir).glob("*.xlsm") if not p.name.startswith("~$"))
#     low = bl_no.lower()
#     exact = [p for p in files if p.name.lower().startswith(low + " ") or p.stem.lower() == low]
#     return exact or [p for p in files if low in p.name.lower()]


# def copy_sheet(src_path, src_sheet_name, dst_wb, title):
#     """
#     1:1 Worksheet transfer into target workbook.
#     - Preserves exact row heights, column widths, print-fit properties, and page scaling.
#     - Strips formula #REF! / 0 artifacts in Column D using python_calamine.
#     """
#     # 1. Load source workbook with openpyxl (for styles and dimensions)
#     swb_temp = load_workbook(src_path, data_only=False)
#     src = swb_temp[src_sheet_name]

#     # 2. Extract pre-calculated static values via Calamine
#     cal_book = CalamineWorkbook.from_path(str(src_path))
#     cal_sheet = cal_book.get_sheet_by_name(src_sheet_name)
#     evaluated_data = cal_sheet.to_python(skip_empty_area=False)

#     dst = dst_wb.create_sheet(title)
#     warnings = []

#     # 3. Copy sheet-level default row/col dimensions
#     if src.sheet_format.defaultColWidth:
#         dst.sheet_format.defaultColWidth = src.sheet_format.defaultColWidth
#     if src.sheet_format.defaultRowHeight:
#         dst.sheet_format.defaultRowHeight = src.sheet_format.defaultRowHeight

#     # 4. Copy cells, styles, and static evaluated values (suppressing stray 0 values)
#     for row in src.iter_rows():
#         for c in row:
#             if c.value is None and not c.has_style:
#                 continue

#             val = c.value
#             try:
#                 eval_val = evaluated_data[c.row - 1][c.column - 1]
#                 if eval_val is not None and eval_val != "" and eval_val != 0:
#                     val = eval_val
#                 else:
#                     val = None
#             except (IndexError, TypeError):
#                 pass

#             d = dst.cell(row=c.row, column=c.column, value=val)
#             if c.has_style:
#                 try:
#                     d.font = copy(c.font)
#                     d.border = copy(c.border)
#                     d.fill = copy(c.fill)
#                     d.number_format = c.number_format
#                     d.protection = copy(c.protection)
#                     d.alignment = copy(c.alignment)
#                 except Exception:
#                     pass

#     # 5. Copy Column Widths explicitly
#     for key, dim in src.column_dimensions.items():
#         nd = dst.column_dimensions[key]
#         nd.min = dim.min
#         nd.max = dim.max
#         if dim.width is not None:
#             nd.width = dim.width
#         nd.hidden = dim.hidden
#         nd.outlineLevel = dim.outlineLevel
#         nd.bestFit = dim.bestFit

#     # 6. Copy Row Heights explicitly
#     for key, dim in src.row_dimensions.items():
#         nd = dst.row_dimensions[key]
#         if dim.height is not None:
#             nd.height = dim.height
#         nd.hidden = dim.hidden
#         nd.outlineLevel = dim.outlineLevel

#     # 7. Copy Page Setup, Scaling, Fit-to-Page, and Print Options for exact 1:1 layout
#     dst.sheet_properties = copy(src.sheet_properties)
#     dst.page_margins = copy(src.page_margins)
#     dst.page_setup = copy(src.page_setup)
#     dst.print_options = copy(src.print_options)
#     dst.sheet_state = src.sheet_state

#     if hasattr(src.page_setup, "orientation"):
#         dst.page_setup.orientation = src.page_setup.orientation
#     if hasattr(src.page_setup, "paperSize"):
#         dst.page_setup.paperSize = src.page_setup.paperSize
#     if hasattr(src.page_setup, "fitToWidth"):
#         dst.page_setup.fitToWidth = src.page_setup.fitToWidth
#     if hasattr(src.page_setup, "fitToHeight"):
#         dst.page_setup.fitToHeight = src.page_setup.fitToHeight
#     if hasattr(src.page_setup, "fitToPage"):
#         dst.page_setup.fitToPage = src.page_setup.fitToPage
#     if hasattr(src.page_setup, "scale"):
#         dst.page_setup.scale = src.page_setup.scale

#     try:
#         dst.HeaderFooter = copy(src.HeaderFooter)
#         dst.print_title_rows = src.print_title_rows
#         if src.print_area:
#             dst.print_area = str(src.print_area).split("!")[-1].replace("$", "")
        
#         for attr in ("view", "showGridLines", "zoomScale", "zoomScaleNormal", "zoomScalePageLayoutView", "zoomScalePageBreakView"):
#             if hasattr(src.sheet_view, attr):
#                 setattr(dst.sheet_view, attr, getattr(src.sheet_view, attr))
#     except Exception:
#         warnings.append("some print or view settings could not be copied")

#     # 8. Merge Cells last
#     for rng in list(src.merged_cells.ranges):
#         try:
#             dst.merge_cells(str(rng))
#         except Exception:
#             pass

#     swb_temp.close()
#     return dst, warnings


# # --------------------------------------------------------------------------
# # WORKER: one source file -> its BL file
# # --------------------------------------------------------------------------
# def attach_one(src_path, target_dir):
#     """-> (status, message, warnings) status: 'added' | 'skipped' | 'error'"""
#     name = src_path.name
#     try:
#         # 1. Read cell G2 value using python_calamine (evaluates formulas like =INDEX(...) safely)
#         cal_book = CalamineWorkbook.from_path(str(src_path))
#         sheet_names = cal_book.sheet_names
#         sname = _find_name(sheet_names, SHEET_NAME)
#         if sname is None:
#             return "error", f"{name}: no '{SHEET_NAME}' sheet (only multi-unit files have one)", []

#         sheet_data = cal_book.get_sheet_by_name(sname).to_python(skip_empty_area=False)

#         # Parse BL_CELL coordinates (G2 -> row 1, col 6)
#         m = re.match(r"^([A-Za-z]+)(\d+)$", BL_CELL)
#         col_idx = 0
#         for ch in m.group(1).upper():
#             col_idx = col_idx * 26 + (ord(ch) - 64)
#         row_idx = int(m.group(2)) - 1
#         col_idx -= 1

#         raw_bl = ""
#         try:
#             raw_bl = sheet_data[row_idx][col_idx]
#         except IndexError:
#             pass

#         bl_no = _bl_text(raw_bl)
#         if not bl_no:
#             return "error", f"{name}: {BL_CELL} (BL no) is empty on '{SHEET_NAME}', was Stage 3 run?", []

#         matches = find_targets(bl_no, target_dir)
#         if not matches:
#             return "error", f"{name}: no BL file containing '{bl_no}' in {Path(target_dir).name}", []
#         if len(matches) > 1:
#             shown = ", ".join(p.name for p in matches[:3])
#             return "error", f"{name}: more than one BL file matches '{bl_no}' ({shown}), nothing was changed", []
#         target = matches[0]

#         # 2. Perform in-memory sheet transfer
#         with _lock_for(target):
#             twb = load_workbook(target, keep_vba=True)
#             existing = _find_name(twb.sheetnames, SHEET_NAME)
#             if existing is not None:
#                 if IF_EXISTS != "replace":
#                     twb.close()
#                     return "skipped", f"{target.name}: already has '{SHEET_NAME}'", []
#                 del twb[existing]

#             _, warnings = copy_sheet(src_path, sname, twb, SHEET_NAME)

#             for i, ws in enumerate(twb.worksheets):
#                 ws.sheet_view.tabSelected = (i == 0)
#             twb.active = 0

#             tmp = target.with_name(target.name + ".tmp")
#             try:
#                 twb.save(tmp)
#                 twb.close()
#                 os.replace(tmp, target)
#             except PermissionError:
#                 return "error", f"{target.name}: could not be saved, is it open in Excel? Close it and run again", []
#             finally:
#                 if tmp.exists():
#                     try:
#                         tmp.unlink()
#                     except OSError:
#                         pass
#         return "added", f"{name} -> {target.name}", [f"{target.name}: {w}" for w in warnings]
#     except Exception as exc:
#         return "error", f"{name}: {exc}", []


# # --------------------------------------------------------------------------
# # ENTRY POINT (called by Api.run_macro)
# # --------------------------------------------------------------------------
# def _collect_sources(folder_or_files):
#     items = [folder_or_files] if isinstance(folder_or_files, (str, os.PathLike)) else list(folder_or_files or [])
#     files, problems = [], []
#     for item in items:
#         p = Path(os.fspath(item))
#         if p.is_dir():
#             multi = next((s for s in p.iterdir() if s.is_dir() and s.name.lower() == "multi"), None)
#             base = multi or p                               # FOR RPA selected -> use its MULTI folder
#             files += [f for f in sorted(base.glob("*.xlsm")) if not f.name.startswith("~$")]
#         elif p.is_file() and p.suffix.lower() == ".xlsm":
#             files.append(p)
#         else:
#             problems.append(f"{p.name}: not an .xlsm file or folder" if p.exists() else f"Not found: {p}")
#     return list(dict.fromkeys(files)), problems


# def process_attach(folder_or_files, target_folder=None):
#     """
#     For every multi-unit file in the selected folder (FOR RPA\\MULTI), copies its "ATCH (2)" sheet to the end
#     of the matching BL file in BL\\ALL. The match is made with the BL no found in cell G2 of that sheet.
#     """
#     try:
#         if not BL_ROOT.is_dir():
#             return {"ok": False, "message": f"The folder 'BLC ALL MACRO' was not found in Downloads: {BL_ROOT}"}

#         target_dir = Path(target_folder) if target_folder else TARGET_DIR
#         if not target_dir.is_dir():
#             return {"ok": False, "message": f"The BL folder was not found: {target_dir}. Run Stage 4 first."}

#         files, problems = _collect_sources(folder_or_files)
#         if not files:
#             return {"ok": False, "message": "No .xlsm files found. Select the FOR RPA\\MULTI folder.",
#                     "errors": problems or None}

#         added, skipped, errors, warnings = 0, [], list(problems), []
#         with ThreadPoolExecutor(max_workers=max(1, min(MAX_WORKERS, len(files)))) as pool:
#             for status, message, warn in pool.map(lambda f: attach_one(f, target_dir), files):
#                 warnings.extend(warn)
#                 if status == "added":
#                     added += 1
#                 elif status == "skipped":
#                     skipped.append(message)
#                 else:
#                     errors.append(message)

#         text = f"Stage 7: '{SHEET_NAME}' added to {added} BL file(s)."
#         if skipped:
#             text += f" {len(skipped)} already had it."
#         if warnings:
#             text += f" {len(warnings)} warning(s)."
#         if errors:
#             text += f" {len(errors)} error(s)."
#         return {
#             "ok": added + len(skipped) > 0,
#             "message": text,
#             "added": added,
#             "skipped": skipped or None,
#             "warnings": warnings or None,
#             "errors": errors or None,
#             "target_dir": str(target_dir),
#         }
#     except Exception as exc:
#         return {"ok": False, "message": f"Attach sheet failed: {exc!r}"}