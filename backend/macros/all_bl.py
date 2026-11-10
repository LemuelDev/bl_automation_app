# import io
# import os
# import re
# from copy import copy
# from datetime import datetime
# from pathlib import Path

# from openpyxl import Workbook, load_workbook

# # --------------------------------------------------------------------------
# # CONFIG
# # --------------------------------------------------------------------------
# SUFFIXES = (".xlsx", ".xlsm")            # the formats that can be combined (.xls / .csv / .xlsb are reported)
# OUTPUT_PREFIX = "BL OUTPUT"              # suggested name of the new file
# _PLACEHOLDER = "__start__"               # the blank first sheet, removed at the end (like the macro)


# # --------------------------------------------------------------------------
# # HELPERS
# # --------------------------------------------------------------------------
# def _collect_files(folder_or_files):
#     """Folder (e.g. BL\\ALL), file, or a list of either -> (files sorted by name, problems)."""
#     if not folder_or_files:
#         return [], []
#     items = [folder_or_files] if isinstance(folder_or_files, (str, os.PathLike)) else list(folder_or_files)

#     files, problems = [], []
#     for item in items:
#         p = Path(os.fspath(item))
#         if p.is_dir():                                    # only the files directly inside the folder
#             for f in p.iterdir():
#                 if not f.is_file() or f.name.startswith("~$"):
#                     continue
#                 if f.suffix.lower() in SUFFIXES:
#                     files.append(f)
#                 elif f.suffix.lower() in (".xls", ".xlsb", ".csv"):
#                     problems.append(f"{f.name}: .xls, .xlsb and .csv files are not supported, skipped")
#         elif p.is_file():
#             if p.suffix.lower() in SUFFIXES:
#                 files.append(p)
#             else:
#                 problems.append(f"{p.name}: only .xlsx and .xlsm files are supported")
#         else:
#             problems.append(f"Not found: {p}")
#     files = sorted(dict.fromkeys(files), key=lambda f: f.name.lower())
#     return files, problems


# def _unique_title(name, used):
#     """Sheet name that is not used yet, the way Excel renames copies: 'Table 1' -> 'Table 1 (2)'."""
#     name = re.sub(r"[\[\]:*?/\\]", "_", name).strip() or "Sheet"
#     name = name[:31]
#     if name.lower() not in used:
#         return name
#     m = re.match(r"^(.*) \((\d+)\)$", name)
#     base, n = (m.group(1), int(m.group(2))) if m else (name, 1)
#     while True:
#         n += 1
#         suffix = f" ({n})"
#         candidate = base[:31 - len(suffix)] + suffix
#         if candidate.lower() not in used:
#             return candidate


# def _with_xlsm(path):
#     """Ensures file path ends with .xlsm extension."""
#     path = re.sub(r"\.(xlsx|xls)$", "", os.fspath(path), flags=re.IGNORECASE)
#     return path if path.lower().endswith(".xlsm") else path + ".xlsm"


# class _Cached:
#     """Saved (calculated) values of a workbook, loaded only if a formula cell is met."""

#     def __init__(self, path):
#         self.path, self.wb = path, None

#     def get(self, sheet, coord):
#         if self.wb is None:
#             self.wb = load_workbook(self.path, data_only=True)
#         return self.wb[sheet][coord].value

#     def close(self):
#         if self.wb is not None:
#             self.wb.close()


# def copy_sheet(src, cached, dst_wb, title):
#     """
#     Copy one worksheet into ANOTHER workbook: values, formatting, widths and heights, merged cells,
#     page / print setup and pictures. A formula is replaced by its saved value when the file has one,
#     otherwise the formula itself is kept (Excel calculates it when the new file is opened).
#     Returns the warnings.
#     """
#     dst = dst_wb.create_sheet(title)
#     warnings = []

#     for row in src.iter_rows():
#         for c in row:
#             v = c.value
#             if c.data_type == "f":
#                 saved = cached.get(src.title, c.coordinate)
#                 if saved is not None:
#                     v = saved
#             if v is None and not c.has_style:
#                 continue
#             d = dst.cell(row=c.row, column=c.column, value=v)
#             if c.has_style:
#                 d.font = copy(c.font)
#                 d.border = copy(c.border)
#                 d.fill = copy(c.fill)
#                 d.number_format = c.number_format
#                 d.protection = copy(c.protection)
#                 d.alignment = copy(c.alignment)

#     for key, dim in src.column_dimensions.items():
#         nd = dst.column_dimensions[key]
#         nd.min, nd.max = dim.min, dim.max
#         nd.width, nd.hidden, nd.outlineLevel, nd.bestFit = dim.width, dim.hidden, dim.outlineLevel, dim.bestFit
#     for key, dim in src.row_dimensions.items():
#         nd = dst.row_dimensions[key]
#         nd.height, nd.hidden, nd.outlineLevel = dim.height, dim.hidden, dim.outlineLevel

#     dst.sheet_format = copy(src.sheet_format)
#     dst.sheet_properties = copy(src.sheet_properties)
#     dst.page_margins = copy(src.page_margins)
#     dst.page_setup = copy(src.page_setup)
#     dst.print_options = copy(src.print_options)
#     dst.sheet_state = src.sheet_state
#     try:
#         dst.HeaderFooter = copy(src.HeaderFooter)
#         dst.print_title_rows = src.print_title_rows
#         if src.print_area:
#             dst.print_area = str(src.print_area).split("!")[-1].replace("$", "")
#         for attr in ("view", "showGridLines", "zoomScale", "zoomScaleNormal", "zoomScaleSheetLayoutView"):
#             setattr(dst.sheet_view, attr, getattr(src.sheet_view, attr))
#     except Exception:
#         warnings.append("some print or view settings could not be copied")

#     for rng in list(src.merged_cells.ranges):                  # merge last, so no cell write hits a merged cell
#         dst.merge_cells(str(rng))

#     images = list(getattr(src, "_images", []) or [])
#     if images:
#         try:
#             from openpyxl.drawing.image import Image
#             for img in images:
#                 new = Image(io.BytesIO(img._data()))
#                 new.anchor = copy(img.anchor)
#                 dst.add_image(new)
#         except Exception:
#             warnings.append("pictures could not be copied")
#     if getattr(src, "_charts", None):
#         warnings.append("charts were not copied")
#     try:
#         if len(list(src.conditional_formatting)):
#             warnings.append("conditional formatting was not copied")
#     except Exception:
#         pass
#     return warnings


# # --------------------------------------------------------------------------
# # ENTRY POINT (called by Api.run_macro)
# # --------------------------------------------------------------------------
# def process_all_bl(folder_or_files, output_path=None, ask_save_path=None):
#     """
#     Combines every sheet of every BL file in the selected folder (BL\\ALL) into ONE new workbook, in file
#     name order, and saves it as a new file. The source files are never changed.

#     output_path:   full path to save to. If given, no dialog is shown.
#     ask_save_path: function(suggested_name, start_dir) -> chosen path, or None if cancelled.
#     """
#     try:
#         files, problems = _collect_files(folder_or_files)
#         if not files:
#             return {"ok": False, "message": "No .xlsx or .xlsm files found. Select the BL\\ALL folder.",
#                     "errors": problems or None}

#         out_wb = Workbook()
#         out_wb.active.title = _PLACEHOLDER                    # the blank first sheet of the macro
#         used, errors, warnings = set(), list(problems), []
#         combined, sheets, renamed = 0, 0, 0

#         for path in files:
#             cached = _Cached(path)
#             created = []
#             try:
#                 src_wb = load_workbook(path)
#                 for ws in src_wb.worksheets:
#                     title = _unique_title(ws.title, used)
#                     created.append(title)
#                     used.add(title.lower())
#                     warnings += [f"{path.name} [{ws.title}]: {w}" for w in copy_sheet(ws, cached, out_wb, title)]
#                     renamed += title != ws.title
#                 src_wb.close()
#                 combined += 1
#                 sheets += len(created)
#             except Exception as exc:
#                 for title in created:                          # do not leave half of a file behind
#                     used.discard(title.lower())
#                     if title in out_wb.sheetnames:
#                         del out_wb[title]
#                 errors.append(f"{path.name}: could not be combined ({exc})")
#             finally:
#                 cached.close()

#         if combined == 0:
#             return {"ok": False, "message": f"Nothing was combined. {errors[0] if errors else ''}".strip(),
#                     "errors": errors or None}

#         del out_wb[_PLACEHOLDER]
#         out_wb.active = 0

#         if not output_path:
#             if ask_save_path is None:
#                 return {"ok": False, "message": "No save location was given."}
#             suggested = f"{OUTPUT_PREFIX}_{datetime.now():%Y-%m-%d_%H%M%S}.xlsm"
#             output_path = ask_save_path(suggested, str(files[0].parent))
#             if not output_path:
#                 return {"ok": False, "cancelled": True, "message": "Save cancelled. No file was created."}

#         out_path = Path(_with_xlsm(output_path))
#         try:
#             out_path.parent.mkdir(parents=True, exist_ok=True)
#             out_wb.code_name = 'ThisWorkbook'
#             out_wb.save(out_path)
#         except PermissionError as exc:
#             return {"ok": False, "message": f"Could not save the file (is it open in Excel?): {exc}"}

#         text = f"BL output saved: {out_path.name} ({combined} file(s), {sheets} sheet(s))."
#         if renamed:
#             text += f" {renamed} sheet(s) renamed to keep the names unique."
#         if warnings:
#             text += f" {len(warnings)} warning(s)."
#         if errors:
#             text += f" {len(errors)} error(s)."
#         return {
#             "ok": True,
#             "message": text,
#             "output_path": str(out_path),
#             "output_dir": str(out_path.parent),
#             "files": combined,
#             "sheets": sheets,
#             "warnings": warnings or None,
#             "errors": errors or None,
#         }
#     except Exception as exc:
#         return {"ok": False, "message": f"All BL failed: {exc!r}"}



import io
import os
import re
import sys
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
SUFFIXES = (".xlsx", ".xlsm")            # the formats that can be combined (.xls / .csv / .xlsb are reported)
OUTPUT_PREFIX = "BL_OUTPUT"              # suggested name of the new file
MODULE_TEMPLATE = "BL_OUTPUT.xlsm"  # in the project's "templates" folder: one blank sheet + the SpellNumber module

_EXTERNAL = re.compile(r"\[\d+\]")       # a formula that points to another workbook, e.g. =[1]Sheet1!A1

WIDTH_SCALE = "auto"
BLANK_ZEROS = False 

# --------------------------------------------------------------------------
# TEMPLATE LOOKUP (bundled with the project)
# --------------------------------------------------------------------------
def _template_candidates():
    here = Path(__file__).resolve().parent            # e.g. <project>/backend/macros
    roots = [here.parent / "templates",               # <project>/backend/templates
             here.parent.parent / "templates",        # <project>/templates
             here / "templates"]                      # <project>/backend/macros/templates
    if getattr(sys, "frozen", False):                 # packaged app (PyInstaller)
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        roots += [base / "templates", base / "backend" / "templates"]
    return [r / MODULE_TEMPLATE for r in roots]


def find_template():
    """-> (path or None, list of paths that were tried)"""
    tried = _template_candidates()
    return next((p for p in tried if p.is_file()), None), tried


# --------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------
def _collect_files(folder_or_files):
    """Folder (e.g. BL\\ALL), file, or a list of either -> (files sorted by name, problems)."""
    if not folder_or_files:
        return [], []
    items = [folder_or_files] if isinstance(folder_or_files, (str, os.PathLike)) else list(folder_or_files)

    files, problems = [], []
    for item in items:
        p = Path(os.fspath(item))
        if p.is_dir():                                    # only the files directly inside the folder
            for f in p.iterdir():
                if not f.is_file() or f.name.startswith("~$"):
                    continue
                if f.suffix.lower() in SUFFIXES:
                    files.append(f)
                elif f.suffix.lower() in (".xls", ".xlsb", ".csv"):
                    problems.append(f"{f.name}: .xls, .xlsb and .csv files are not supported, skipped")
        elif p.is_file():
            if p.suffix.lower() in SUFFIXES:
                files.append(p)
            else:
                problems.append(f"{p.name}: only .xlsx and .xlsm files are supported")
        else:
            problems.append(f"Not found: {p}")
    files = sorted(dict.fromkeys(files), key=lambda f: f.name.lower())
    return files, problems


def _unique_title(name, used):
    """Sheet name that is not used yet, the way Excel renames copies: 'Table 1' -> 'Table 1 (2)'."""
    name = re.sub(r"[\[\]:*?/\\]", "_", name).strip() or "Sheet"
    name = name[:31]
    if name.lower() not in used:
        return name
    m = re.match(r"^(.*) \((\d+)\)$", name)
    base, n = (m.group(1), int(m.group(2))) if m else (name, 1)
    while True:
        n += 1
        suffix = f" ({n})"
        candidate = base[:31 - len(suffix)] + suffix
        if candidate.lower() not in used:
            return candidate


def _with_xlsm(path):
    path = re.sub(r"\.(xlsx|xls)$", "", os.fspath(path), flags=re.IGNORECASE)
    return path if path.lower().endswith(".xlsm") else path + ".xlsm"


def _retarget(text, mapping):
    """Formulas that name a sheet of the same file must follow that sheet when it is renamed
    ('Table 1'!G30 -> 'Table 1 (2)'!G30), like Excel does when it copies sheets."""
    if not mapping:
        return text
    names = "|".join(sorted((re.escape(o) for o in mapping), key=len, reverse=True))
    rx = re.compile(r"(?:'(" + names + r")'|(?<![\w.'])(" + names + r"))!")
    return rx.sub(lambda m: "'" + mapping[m.group(1) or m.group(2)].replace("'", "''") + "'!", text)


class _Cached:
    """Saved (calculated) values of a workbook, loaded only if they are needed."""

    def __init__(self, path):
        self.path, self.wb = path, None

    def get(self, sheet, coord):
        if self.wb is None:
            self.wb = load_workbook(self.path, data_only=True)
        return self.wb[sheet][coord].value

    def close(self):
        if self.wb is not None:
            self.wb.close()


def _formula_value(cell, value, cached, sheet_title, mapping, warnings):
    text = value.text if isinstance(value, ArrayFormula) else value
    if not isinstance(text, str):                         # other kinds of formula: keep them as they are
        return value
    if _EXTERNAL.search(text):                            # a link to another file cannot be carried over
        warnings.append(f"{cell.coordinate}: a formula with a link to another file was replaced by its saved value")
        return cached.get(sheet_title, cell.coordinate)
    text = _retarget(text, mapping)
    return ArrayFormula(ref=value.ref, text=text) if isinstance(value, ArrayFormula) else text


def copy_sheet(src, cached, dst_wb, title, mapping, dst=None):
    """
    Copy one worksheet into ANOTHER workbook: values, formulas, formatting, widths and heights, merged cells,
    page / print setup and pictures. Formulas are kept as formulas (Excel calculates them when the new file is
    opened). If dst is given, that existing sheet receives the copy. Returns the warnings.
    """
    if dst is None:
        dst, keep_code_name = dst_wb.create_sheet(title), None
    else:
        keep_code_name = dst.sheet_properties.codeName     # the sheet keeps the name its macros are tied to
        dst.title = title
    warnings = []

    for row in src.iter_rows():
        for c in row:
            v = c.value

            # Evaluate formulas and cached values
            if c.data_type == "f":
                v = _formula_value(c, v, cached, src.title, mapping, warnings)

            # Apply BLANK_ZEROS rule if explicitly enabled
            if BLANK_ZEROS and isinstance(v, (int, float)) and not isinstance(v, bool) and v == 0:
                v = None

            # Skip empty cells with no styling
            if v is None and not c.has_style:
                continue

            d = dst.cell(row=c.row, column=c.column, value=v)

            if c.has_style:
                f = copy(c.font)
                if f.scheme:
                    f.scheme = None                        # prevent theme shift across workbooks
                d.font = f
                d.border = copy(c.border)
                d.fill = copy(c.fill)
                d.number_format = c.number_format
                d.protection = copy(c.protection)
                d.alignment = copy(c.alignment)

    # Column widths and row heights
    for key, dim in src.column_dimensions.items():
        nd = dst.column_dimensions[key]
        nd.min, nd.max = dim.min, dim.max
        nd.width, nd.hidden, nd.outlineLevel, nd.bestFit = dim.width, dim.hidden, dim.outlineLevel, dim.bestFit
    for key, dim in src.row_dimensions.items():
        nd = dst.row_dimensions[key]
        nd.height, nd.hidden, nd.outlineLevel = dim.height, dim.hidden, dim.outlineLevel

    # # Page and view settings
    # dst.sheet_format = copy(src.sheet_format)
    # dst.sheet_properties = copy(src.sheet_properties)
    # dst.sheet_properties.codeName = keep_code_name         # code names must remain unique
    # dst.page_margins = copy(src.page_margins)
    # dst.page_setup = copy(src.page_setup)
    # dst.print_options = copy(src.print_options)
    # dst.sheet_state = src.sheet_state
    
    # Sheet format, properties, and page settings
    dst.sheet_format = copy(src.sheet_format)
    dst.sheet_properties = copy(src.sheet_properties)
    dst.sheet_properties.codeName = keep_code_name         # code names must remain unique
    dst.page_margins = copy(src.page_margins)
    dst.page_setup = copy(src.page_setup)
    dst.page_setup.orientation = src.page_setup.orientation or "portrait"
    dst.page_setup.paperSize = src.page_setup.paperSize or 9
    
    # Enable Auto-Fit to 1 Page Wide across different PC resolutions/printers
    dst.page_setup.fitToWidth = 1
    dst.page_setup.fitToHeight = 0                          # Allows rows to flow naturally across multiple pages
    dst.sheet_properties.pageSetUpPr.fitToPage = True

    dst.print_options = copy(src.print_options)
    dst.sheet_state = src.sheet_state

    try:
        dst.HeaderFooter = copy(src.HeaderFooter)
        dst.print_title_rows = src.print_title_rows
        if src.print_area:
            dst.print_area = str(src.print_area).split("!")[-1].replace("$", "")
        for attr in ("view", "showGridLines", "showZeros", "zoomScale", "zoomScaleNormal", "zoomScaleSheetLayoutView"):
            if hasattr(src.sheet_view, attr):
                setattr(dst.sheet_view, attr, getattr(src.sheet_view, attr))
    except Exception:
        warnings.append("some print or view settings could not be copied")

    for rng in list(src.merged_cells.ranges):                  # merge last to avoid collision
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

    if getattr(src, "_charts", None):
        warnings.append("charts were not copied")
    try:
        if len(list(src.conditional_formatting)):
            warnings.append("conditional formatting was not copied")
    except Exception:
        pass

    return warnings


def _open_module_template():
    """-> (workbook, None) or (None, error message). The template holds the SpellNumber module."""
    template, tried = find_template()
    if template is None:
        listing = "\n".join(f"  {p}" for p in tried)
        return None, (f"The template '{MODULE_TEMPLATE}' was not found in the project's templates folder "
                    f"(a blank workbook that contains the SpellNumber module, saved as .xlsm). Looked in:\n{listing}")
    wb = load_workbook(template, keep_vba=True)
    problem = None
    if not wb.vba_archive or "xl/vbaProject.bin" not in wb.vba_archive.namelist():
        problem = f"'{MODULE_TEMPLATE}' has no macro project. Add the SpellNumber module in Excel and save it as .xlsm."
    else:
        ws = wb.worksheets[0]
        if len(wb.worksheets) != 1 or ws.max_row > 1 or ws.max_column > 1 or ws["A1"].value is not None:
            problem = f"'{MODULE_TEMPLATE}' must contain exactly one blank sheet."
    if problem:
        if wb.vba_archive:
            wb.vba_archive.close()
        return None, problem
    return wb, None


# --------------------------------------------------------------------------
# ENTRY POINT (called by Api.run_macro)
# --------------------------------------------------------------------------
def process_all_bl(folder_or_files, output_path=None, ask_save_path=None):
    """
    Combines every sheet of every BL file in the selected folder (BL\\ALL) into ONE new .xlsm, in file name
    order. The new file starts from the module template, so it already contains the SpellNumber module.
    The source files are never changed.

    output_path:   full path to save to. If given, no dialog is shown.
    ask_save_path: function(suggested_name, start_dir) -> chosen path, or None if cancelled.
    """
    try:
        files, problems = _collect_files(folder_or_files)
        if not files:
            return {"ok": False, "message": "No .xlsx or .xlsm files found. Select the BL\\ALL folder.",
                    "errors": problems or None}

        out_wb, problem = _open_module_template()
        if problem:
            return {"ok": False, "message": problem}
        receiver = out_wb.worksheets[0]                       # the template's blank sheet takes the first copy

        used, errors, warnings = set(), list(problems), []
        combined, sheets, renamed = 0, 0, 0
        receiver_used = False

        for path in files:
            cached = _Cached(path)
            try:
                try:
                    src_wb = load_workbook(path)
                except Exception as exc:
                    errors.append(f"{path.name}: could not be read ({exc})")
                    continue

                src_sheets = src_wb.worksheets
                titles = []
                for ws in src_sheets:                          # decide every new name first, so formulas can follow
                    t = _unique_title(ws.title, used)
                    used.add(t.lower())
                    titles.append(t)
                mapping = {ws.title: t for ws, t in zip(src_sheets, titles) if ws.title != t}

                took_receiver = not receiver_used
                created = []
                try:
                    for ws, title in zip(src_sheets, titles):
                        target = None
                        if not receiver_used:
                            target, receiver_used = receiver, True
                        warnings += [f"{path.name} [{ws.title}]: {w}"
                                for w in copy_sheet(ws, cached, out_wb, title, mapping, dst=target)]
                        created.append(title)
                    combined += 1
                    sheets += len(created)
                    renamed += len(mapping)
                except Exception as exc:
                    if took_receiver:
                        raise                                  # the template's own sheet was already touched
                    for t in titles:
                        used.discard(t.lower())
                        if t in out_wb.sheetnames:
                            del out_wb[t]
                    errors.append(f"{path.name}: could not be combined ({exc})")
                finally:
                    src_wb.close()
            finally:
                cached.close()

        if combined == 0:
            return {"ok": False, "message": f"Nothing was combined. {errors[0] if errors else ''}".strip(),
                    "errors": errors or None}

        # open on the first visible sheet, and make Excel calculate every formula when the file is opened
        first = next((i for i, ws in enumerate(out_wb.worksheets) if ws.sheet_state == "visible"), 0)
        for i, ws in enumerate(out_wb.worksheets):
            ws.sheet_view.tabSelected = (i == first)
        out_wb.active = first
        out_wb.calculation.fullCalcOnLoad = True

        if not output_path:
            if ask_save_path is None:
                return {"ok": False, "message": "No save location was given."}
            suggested = f"{OUTPUT_PREFIX}_{datetime.now():%Y-%m-%d_%H%M%S}.xlsm"
            output_path = ask_save_path(suggested, str(files[0].parent))
            if not output_path:
                return {"ok": False, "cancelled": True, "message": "Save cancelled. No file was created."}

        out_path = Path(_with_xlsm(output_path))
        try:
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_wb.save(out_path)
        except PermissionError as exc:
            return {"ok": False, "message": f"Could not save the file (is it open in Excel?): {exc}"}

        text = f"BL output saved: {out_path.name} ({combined} file(s), {sheets} sheet(s), SpellNumber module included)."
        if renamed:
            text += f" {renamed} sheet(s) renamed to keep the names unique."
        if warnings:
            text += f" {len(warnings)} warning(s)."
        if errors:
            text += f" {len(errors)} error(s)."
        return {
            "ok": True,
            "message": text,
            "output_path": str(out_path),
            "output_dir": str(out_path.parent),
            "files": combined,
            "sheets": sheets,
            "warnings": warnings or None,
            "errors": errors or None,
        }
    except Exception as exc:
        return {"ok": False, "message": f"All BL failed: {exc!r}"}