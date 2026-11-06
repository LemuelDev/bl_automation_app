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
TEMPLATE_CHECKER = "Checker_result.xlsx"
TEMPLATE_IV_KANRI = "IV番号管理表.xlsx"

MAX_WORKERS = 8
INPUT_SUBFOLDERS = ("single", "multi")
HEADER_ROWS = 1

# Cell in BL file -> Column in Checker_result.xlsx
COLUMNS = [
    ("A3", 1),     # Shipper
    ("H2", 2),     # BL no
    ("B28", 3),    # Units
    ("G51", 4),    # Weight
    ("I30", 5),    # m3
]
BL_NO_ADDRESS = "H2"
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


# --------------------------------------------------------------------------
# TEMPLATE LOOKUP
# --------------------------------------------------------------------------
def _template_candidates(template_name):
    here = Path(__file__).resolve().parent
    roots = [
        here.parent / "templates",
        here.parent.parent / "templates",
        here / "templates",
    ]
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        roots += [base / "templates", base / "backend" / "templates"]
    return [r / template_name for r in roots]

# ADD "=TEMPLATE_CHECKER" AS A DEFAULT ARGUMENT HERE:
def find_template(template_name=TEMPLATE_CHECKER):
    """-> (path or None, list of paths that were tried)"""
    tried = _template_candidates(template_name)
    return next((p for p in tried if p.is_file()), None), tried


# --------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------
def _blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def get_cell(rows, address):
    m = re.match(r"^([A-Za-z]+)(\d+)$", address)
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + (ord(ch) - 64)
    try:
        return rows[int(m.group(2)) - 1][col - 1]
    except IndexError:
        return ""


def _clean_val(v):
    if v is None or v == "":
        return None
    if isinstance(v, str):
        v = _ILLEGAL.sub("", v)
        return v if v else None
    return v


def _with_xlsx(path):
    path = re.sub(r"\.(xlsm|xls)$", "", os.fspath(path), flags=re.IGNORECASE)
    return path if path.lower().endswith(".xlsx") else path + ".xlsx"


def _collect_files(folder_or_files):
    if not folder_or_files:
        return []
    items = [folder_or_files] if isinstance(folder_or_files, (str, os.PathLike)) else list(folder_or_files)

    found = {}

    def add(f):
        if f.suffix.lower() == ".xlsm" and not f.name.startswith("~$"):
            found[os.path.normcase(str(f.resolve()))] = f

    for item in items:
        p = Path(os.fspath(item))
        if p.is_dir():
            for f in p.glob("*.xlsm"):
                add(f)
            for sub in p.iterdir():
                if sub.is_dir() and sub.name.lower() in INPUT_SUBFOLDERS:
                    for f in sub.glob("*.xlsm"):
                        add(f)
        elif p.is_file():
            add(p)
    return sorted(found.values(), key=lambda x: x.name.lower())


def _last_used_row(template, ncols):
    book = CalamineWorkbook.from_path(str(template))
    rows = book.get_sheet_by_name(book.sheet_names[0]).to_python(skip_empty_area=False)
    last = 0
    for i, row in enumerate(rows, start=1):
        if any(not _blank(v) for v in row[:ncols]):
            last = i
    return last


def read_bl_file(path):
    try:
        book = CalamineWorkbook.from_path(str(path))
        rows = book.get_sheet_by_name(book.sheet_names[0]).to_python(skip_empty_area=False)
    except Exception as exc:
        return "error", f"{path.name}: {exc}"

    if _blank(get_cell(rows, BL_NO_ADDRESS)):
        return "skip", f"{path.name}: {BL_NO_ADDRESS} (BL no) is empty, not a BL form"
    return "ok", [get_cell(rows, addr) for addr, _ in COLUMNS]


# --------------------------------------------------------------------------
# ENTRY POINT
# --------------------------------------------------------------------------
def process_checker_result(folder_or_files, template_folder=None, output_path=None, ask_save_path=None):
    # 1. Locate both templates
    checker_tpl, tried_checker = find_template(TEMPLATE_CHECKER)
    iv_tpl, tried_iv = find_template(TEMPLATE_IV_KANRI)

    if checker_tpl is None:
        listing = "\n".join(f"  {p}" for p in tried_checker)
        return {"ok": False, "message": f"Template {TEMPLATE_CHECKER} not found in project. Looked in:\n{listing}"}

    if iv_tpl is None:
        listing = "\n".join(f"  {p}" for p in tried_iv)
        return {"ok": False, "message": f"Template {TEMPLATE_IV_KANRI} not found in project. Looked in:\n{listing}"}

    files = _collect_files(folder_or_files)
    if not files:
        return {"ok": False, "message": "No .xlsm files found. Select the BL folder."}

    rows_out, errors, skipped = [], [], []
    with ThreadPoolExecutor(max_workers=max(1, min(MAX_WORKERS, len(files)))) as pool:
        for status, payload in pool.map(read_bl_file, files):
            if status == "ok":
                rows_out.append(payload)
            elif status == "skip":
                skipped.append(payload)
            else:
                errors.append(payload)

    if not rows_out:
        detail = (errors + skipped)[:2]
        return {
            "ok": False,
            "message": f"Nothing to write. {' | '.join(detail)}".strip(),
            "errors": errors or None,
            "skipped": skipped or None,
        }

    try:
        timestamp_str = datetime.now().strftime("%Y-%m-%d_%H%M%S")

        # Determine target directory
        if output_path:
            out_dir = Path(output_path).parent
            checker_filename = Path(output_path).name
        else:
            if ask_save_path is None:
                return {"ok": False, "message": "No save location was given."}
            suggested = f"{checker_tpl.stem}_{timestamp_str}{checker_tpl.suffix}"
            chosen = ask_save_path(suggested, str(files[0].parent))
            if not chosen:
                return {"ok": False, "cancelled": True, "message": "Save cancelled. No file was created."}
            out_path_chk = Path(_with_xlsx(chosen))
            out_dir = out_path_chk.parent
            checker_filename = out_path_chk.name

        out_path_checker = out_dir / checker_filename
        out_path_iv = out_dir / f"{iv_tpl.stem}_{timestamp_str}.xlsx"
        out_dir.mkdir(parents=True, exist_ok=True)

        # ------------------------------------------------------------------
        # FILE 1: Populate Checker_result.xlsx
        # ------------------------------------------------------------------
        ncols = max(c for _, c in COLUMNS)
        start_row_chk = max(_last_used_row(checker_tpl, ncols), HEADER_ROWS) + 1

        wb_chk = load_workbook(checker_tpl)
        ws_chk = wb_chk.worksheets[0]
        for r, values in enumerate(rows_out, start=start_row_chk):
            for (_, col), v in zip(COLUMNS, values):
                val = _clean_val(v)
                if val is not None:
                    ws_chk.cell(row=r, column=col, value=val)

        wb_chk.save(out_path_checker)
        wb_chk.close()

        # ------------------------------------------------------------------
        # FILE 2: Populate IV番号管理表.xlsx
        # Mapping:
        #   Shipper Name (values[0]) -> Col B (2)
        #   BL Number    (values[1]) -> Col D (4)
        #   Units        (values[2]) -> Col E (5)
        # ------------------------------------------------------------------
        start_row_iv = HEADER_ROWS + 1

        wb_iv = load_workbook(iv_tpl)
        ws_iv = wb_iv.worksheets[0]
        for r, values in enumerate(rows_out, start=start_row_iv):
            shipper = _clean_val(values[0])
            bl_no = _clean_val(values[1])
            units = _clean_val(values[2])

            if shipper is not None:
                ws_iv.cell(row=r, column=2, value=shipper)  # Col B
            if bl_no is not None:
                ws_iv.cell(row=r, column=4, value=bl_no)    # Col D
            if units is not None:
                ws_iv.cell(row=r, column=5, value=units)    # Col E

        wb_iv.save(out_path_iv)
        wb_iv.close()

    except PermissionError as exc:
        return {"ok": False, "message": f"Could not save output file (is it open in Excel?): {exc}"}
    except Exception as exc:
        return {"ok": False, "message": f"Checker result error: {exc}"}

    message = f"Checker & IV番号管理表 saved: {out_path_checker.name} & {out_path_iv.name} ({len(rows_out)} BL file(s))."
    if skipped:
        message += f" {len(skipped)} skipped (no BL no)."
    if errors:
        message += f" {len(errors)} error(s)."

    return {
        "ok": True,
        "message": message,
        "output_path": str(out_path_checker),
        "iv_kanri_path": str(out_path_iv),
        "output_dir": str(out_dir),
        "skipped": skipped or None,
        "errors": errors or None,
    }








# import os
# import re
# import sys
# from concurrent.futures import ThreadPoolExecutor
# from datetime import datetime
# from pathlib import Path

# from openpyxl import load_workbook
# from python_calamine import CalamineWorkbook

# # --------------------------------------------------------------------------
# # CONFIG
# # --------------------------------------------------------------------------
# TEMPLATE_NAME = "Checker_result.xlsx"     # lives in the project's "templates" folder
# MAX_WORKERS = 8
# INPUT_SUBFOLDERS = ("single", "multi")    # looked at inside the folder you select (e.g. the BL folder)
# HEADER_ROWS = 1                           # rows at the top of the template that are headers

# # Cell in the BL file's first sheet -> column in the Checker result
# COLUMNS = [
#     ("A3", 1),     # shipper
#     ("H2", 2),     # BL no
#     ("B28", 3),    # units
#     ("G51", 4),    # weight
#     ("I30", 5),    # m3
# ]
# BL_NO_ADDRESS = "H2"
# _ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")   # characters openpyxl refuses to write


# # --------------------------------------------------------------------------
# # TEMPLATE LOOKUP (bundled with the project)
# # --------------------------------------------------------------------------
# def _template_candidates():
#     here = Path(__file__).resolve().parent            # e.g. <project>/backend/macros
#     roots = [here.parent / "templates",               # <project>/backend/templates
#              here.parent.parent / "templates",        # <project>/templates
#              here / "templates"]                      # <project>/backend/macros/templates
#     if getattr(sys, "frozen", False):                 # packaged app (PyInstaller)
#         base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
#         roots += [base / "templates", base / "backend" / "templates"]
#     return [r / TEMPLATE_NAME for r in roots]


# def find_template():
#     """-> (path or None, list of paths that were tried)"""
#     tried = _template_candidates()
#     return next((p for p in tried if p.is_file()), None), tried


# # --------------------------------------------------------------------------
# # HELPERS
# # --------------------------------------------------------------------------
# def _blank(v):
#     return v is None or (isinstance(v, str) and v.strip() == "")


# def get_cell(rows, address):
#     """Value of an A1 address from calamine rows (row 0 == Excel row 1)."""
#     m = re.match(r"^([A-Za-z]+)(\d+)$", address)
#     col = 0
#     for ch in m.group(1).upper():
#         col = col * 26 + (ord(ch) - 64)
#     try:
#         return rows[int(m.group(2)) - 1][col - 1]
#     except IndexError:
#         return ""


# def _with_xlsx(path):
#     """Make sure the chosen file ends with .xlsx."""
#     path = re.sub(r"\.(xlsm|xls)$", "", os.fspath(path), flags=re.IGNORECASE)
#     return path if path.lower().endswith(".xlsx") else path + ".xlsx"


# def _collect_files(folder_or_files):
#     """
#     Folder (e.g. the BL folder), file, or a list of either -> list of .xlsm files sorted by name.
#     A folder contributes its own .xlsm files plus the ones inside its SINGLE and MULTI subfolders.
#     """
#     if not folder_or_files:
#         return []
#     items = [folder_or_files] if isinstance(folder_or_files, (str, os.PathLike)) else list(folder_or_files)

#     found = {}

#     def add(f):
#         if f.suffix.lower() == ".xlsm" and not f.name.startswith("~$"):
#             found[os.path.normcase(str(f.resolve()))] = f

#     for item in items:
#         p = Path(os.fspath(item))
#         if p.is_dir():
#             for f in p.glob("*.xlsm"):
#                 add(f)
#             for sub in p.iterdir():
#                 if sub.is_dir() and sub.name.lower() in INPUT_SUBFOLDERS:
#                     for f in sub.glob("*.xlsm"):
#                         add(f)
#         elif p.is_file():
#             add(p)
#     return sorted(found.values(), key=lambda x: x.name.lower())


# def _last_used_row(template, ncols):
#     """Last row of the template that has anything in its first ncols columns (0 if none)."""
#     book = CalamineWorkbook.from_path(str(template))
#     rows = book.get_sheet_by_name(book.sheet_names[0]).to_python(skip_empty_area=False)
#     last = 0
#     for i, row in enumerate(rows, start=1):
#         if any(not _blank(v) for v in row[:ncols]):
#             last = i
#     return last


# def read_bl_file(path):
#     """-> (status, payload): ("ok", [values]) / ("skip", message) / ("error", message)"""
#     try:
#         book = CalamineWorkbook.from_path(str(path))
#         rows = book.get_sheet_by_name(book.sheet_names[0]).to_python(skip_empty_area=False)   # Sheets(1)
#     except Exception as exc:
#         return "error", f"{path.name}: {exc}"

#     if _blank(get_cell(rows, BL_NO_ADDRESS)):
#         return "skip", f"{path.name}: {BL_NO_ADDRESS} (BL no) is empty, not a BL form"
#     return "ok", [get_cell(rows, addr) for addr, _ in COLUMNS]


# # --------------------------------------------------------------------------
# # ENTRY POINT (called by Api.run_macro)
# # --------------------------------------------------------------------------
# def process_checker_result(folder_or_files, template_folder=None, output_path=None, ask_save_path=None):
#     """
#     Collects shipper, BL no, units, weight and m3 from the finished BL files and writes one row per
#     file into a fresh copy of the bundled "Checker result.xlsx" template, saved as a NEW file.
#     The template itself is never modified.

#     output_path:   full path to save to. If given, no dialog is shown.
#     ask_save_path: function(suggested_name, start_dir) -> chosen path, or None if cancelled.
#     template_folder is accepted only so callers do not break; it is ignored.
#     """
#     template, tried = find_template()
#     if template is None:
#         listing = "\n".join(f"  {p}" for p in tried)
#         return {"ok": False, "message": f"Template {TEMPLATE_NAME} not found in the project. Looked in:\n{listing}"}

#     files = _collect_files(folder_or_files)
#     if not files:
#         return {"ok": False, "message": "No .xlsm files found. Select the BL folder "
#                                         "(or the SINGLE / MULTI folder inside it)."}

#     rows_out, errors, skipped = [], [], []
#     with ThreadPoolExecutor(max_workers=max(1, min(MAX_WORKERS, len(files)))) as pool:
#         for status, payload in pool.map(read_bl_file, files):          # keeps the sorted order
#             if status == "ok":
#                 rows_out.append(payload)
#             elif status == "skip":
#                 skipped.append(payload)
#             else:
#                 errors.append(payload)

#     if not rows_out:
#         detail = (errors + skipped)[:2]
#         return {"ok": False, "message": f"Nothing to write. {' | '.join(detail)}".strip(),
#                 "errors": errors or None, "skipped": skipped or None}

#     try:
#         ncols = max(c for _, c in COLUMNS)
#         start_row = max(_last_used_row(template, ncols), HEADER_ROWS) + 1   # first empty row, like End(xlUp)+1

#         wb = load_workbook(template)
#         ws = wb.worksheets[0]                                               # Sheets(1)
#         for r, values in enumerate(rows_out, start=start_row):
#             for (addr, col), v in zip(COLUMNS, values):
#                 if v is None or v == "":
#                     continue
#                 if isinstance(v, str):
#                     v = _ILLEGAL.sub("", v)
#                     if not v:
#                         continue
#                 ws.cell(row=r, column=col, value=v)

#         # the data is ready: now ask where to save it
#         if not output_path:
#             if ask_save_path is None:
#                 wb.close()
#                 return {"ok": False, "message": "No save location was given."}
#             suggested = f"{template.stem}_{datetime.now():%Y-%m-%d_%H%M%S}{template.suffix}"
#             output_path = ask_save_path(suggested, str(files[0].parent))
#             if not output_path:
#                 wb.close()
#                 return {"ok": False, "cancelled": True, "message": "Save cancelled. No file was created."}

#         out_path = Path(_with_xlsx(output_path))
#         out_path.parent.mkdir(parents=True, exist_ok=True)
#         wb.save(out_path)
#         wb.close()
#     except PermissionError as exc:
#         return {"ok": False, "message": f"Could not save the file (is it open in Excel?): {exc}"}
#     except Exception as exc:
#         return {"ok": False, "message": f"Checker result error: {exc}"}

#     message = f"Checker result saved: {out_path.name} ({len(rows_out)} BL file(s))."
#     if skipped:
#         message += f" {len(skipped)} skipped (no BL no)."
#     if errors:
#         message += f" {len(errors)} error(s)."
#     return {
#         "ok": True,
#         "message": message,
#         "output_path": str(out_path),
#         "output_dir": str(out_path.parent),
#         "skipped": skipped or None,
#         "errors": errors or None,
#     }
