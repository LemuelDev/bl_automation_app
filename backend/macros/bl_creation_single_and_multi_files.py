import io
import os
import re
import shutil
import sys
import zipfile
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from python_calamine import CalamineWorkbook

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
SINGLE, MULTI = "single", "multi"

# Dynamic path resolution (the WIP BL folders; the BL templates now live inside the project)
BL_ROOT = Path(os.path.expanduser("~")) / "Downloads" / "BLC ALL MACRO"

# BL templates (with the signature) in the project's "templates" folder. (kind, navis) -> file name
TEMPLATE_FILES = {
    (SINGLE, False): "SINGLE_TEMPLATE.xlsm",
    (MULTI, False): "MULTI_TEMPLATE.xlsm",
    (SINGLE, True): "NAVIS_SINGLE.xlsm",
    (MULTI, True): "NAVIS_MULTI.xlsm",
}

# Vessel details the user enters: where each one is written in every BL file
VESSEL_CELLS = {
    "name": ("B23",),            # vessel name
    "voyage": ("E23",),          # voy no.
    "pol": ("B24", "B56"),       # port of loading
    "pod": ("E24",),             # port of discharge
    "etd": ("B57",),             # ETD (written as a real date)
}
VESSEL_LABELS = {"name": "Vessel name", "voyage": "Voy no.", "pol": "POL", "pod": "POD"}
ETD_NUMBER_FORMAT = "mm/dd/yyyy"  # how the ETD date shows in the BL
UPPERCASE_DETAILS = True          # BL documents are in capitals; spaces are tidied too

BL_OUT = BL_ROOT / "WIP BL" / "BL"
OUTPUT_DIRS = {
    SINGLE: BL_OUT / "SINGLE",
    MULTI: BL_OUT / "MULTI",
}
ALL_DIR = BL_OUT / "ALL"
CLEAR_DIRS = [BL_OUT / "SINGLE", BL_OUT / "MULTI", ALL_DIR]

INPUT_SUBFOLDERS = ("single", "multi")
TEMPLATE_SHEET = "Table 1"
MAX_WORKERS = 8
USE_PROCESSES = False

# Cells that must hold a real, non-zero value, otherwise Stage 3 has not been run
REQUIRED = {
    MULTI: (("C58", "units"), ("M58", "weight"), ("O58", "m3")),
    SINGLE: (("M58", "weight"), ("O58", "m3")),
}

SHIPPER_MAP = {"D6": "A3", "D7": "A4", "D8": "A5"}
COPY_OR_CLEAR_MAP = {
    "D9": "A6",
    "D14": "A11", "D15": "A12", "D16": "A13",
    "D22": "A17", "D23": "A18", "D24": "A19",
}

_EXCEL_FILE = re.compile(r"\.(xlsx|xlsm|xlsb|xls)(\.\d+)?$", re.IGNORECASE)
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_PICTURE = (".png", ".jpg", ".jpeg", ".gif", ".bmp")


# --------------------------------------------------------------------------
# TEMPLATES (bundled with the project, one set with and one without Navis)
# --------------------------------------------------------------------------
def _template_roots():
    here = Path(__file__).resolve().parent            # e.g. <project>/backend/macros
    roots = [here.parent / "templates",               # <project>/backend/templates
             here.parent.parent / "templates",        # <project>/templates
             here / "templates"]                      # <project>/backend/macros/templates
    if getattr(sys, "frozen", False):                 # packaged app (PyInstaller)
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        roots += [base / "templates", base / "backend" / "templates"]
    return roots


def find_template(kind, navis):
    """-> (path or None, list of paths that were tried)"""
    tried = [r / TEMPLATE_FILES[(kind, bool(navis))] for r in _template_roots()]
    return next((p for p in tried if p.is_file()), None), tried


def missing_templates(navis):
    """Messages for the templates that cannot be found (empty list = all there)."""
    problems = []
    for kind in (SINGLE, MULTI):
        path, tried = find_template(kind, navis)
        if path is None:
            problems.append(f"Template not found: {TEMPLATE_FILES[(kind, bool(navis))]} "
                            f"(looked in {tried[0].parent} and {len(tried) - 1} other place(s))")
    return problems


def _loaded_images(wb):
    return sum(len(getattr(ws, "_images", []) or []) for ws in wb.worksheets)


def _template_problem(template_bytes, label):
    """The signature must survive: openpyxl needs Pillow to keep pictures, and cannot keep header / footer pictures."""
    try:
        with zipfile.ZipFile(io.BytesIO(template_bytes)) as z:
            pictures = [n for n in z.namelist() if n.startswith("xl/media/") and n.lower().endswith(_PICTURE)]
    except zipfile.BadZipFile:
        return f"{label} is not a valid Excel file"
    if not pictures:
        return None
    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    loaded = _loaded_images(wb)
    wb.close()
    if loaded == 0:
        return (f"{label} contains a picture (the signature) that cannot be carried into the BL files. "
                f"Install Pillow (pip install pillow), and keep the signature as a normal picture, not a header or footer picture.")
    return None


# --------------------------------------------------------------------------
# VESSEL DETAILS
# --------------------------------------------------------------------------
def _parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return datetime.strptime(value.strip()[:10], "%Y-%m-%d").date()     # what the date picker sends
        except ValueError:
            return None
    return None


def validate_details(details):
    """-> (clean details, None) or (None, message). Every detail is required, nothing is guessed."""
    if not isinstance(details, dict):
        return None, "Enter the vessel details first (vessel name, voy no., POL, POD, ETD and Navis)."

    clean, missing = {}, []
    for key, label in VESSEL_LABELS.items():
        text = re.sub(r"\s+", " ", _ILLEGAL.sub("", str(details.get(key) or ""))).strip()
        if not text:
            missing.append(label)
        clean[key] = text.upper() if UPPERCASE_DETAILS else text

    etd = _parse_date(details.get("etd"))
    if etd is None:
        missing.append("ETD (a valid date)")
    clean["etd"] = etd

    navis = details.get("navis")
    if isinstance(navis, str):
        navis = {"yes": True, "no": False}.get(navis.strip().lower())
    if not isinstance(navis, bool):
        missing.append("Navis (yes or no)")
    clean["navis"] = navis

    if missing:
        return None, "The vessel details are incomplete: " + ", ".join(missing) + "."
    return clean, None


def _vessel_cells(details):
    cells = {}
    for key, addresses in VESSEL_CELLS.items():
        for addr in addresses:
            cells[addr] = details[key]
    return cells


def _describe(details):
    return (f"{details['name']} V.{details['voyage']}, {details['pol']} to {details['pod']}, "
            f"ETD {details['etd']:%d-%b-%Y}, {'NAVIS' if details['navis'] else 'standard'} templates")


# --------------------------------------------------------------------------
# CELL HELPERS
# --------------------------------------------------------------------------
def get_cell(rows, address):
    m = re.match(r"^([A-Za-z]+)(\d+)$", address)
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + (ord(ch) - 64)
    try:
        return rows[int(m.group(2)) - 1][col - 1]
    except IndexError:
        return ""


def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def is_numeric(v):
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        try:
            float(v.strip().replace(",", ""))
            return True
        except ValueError:
            return False
    return False


def to_float(v):
    return float(v.strip().replace(",", "")) if isinstance(v, str) else float(v)


def is_empty_or_zero(v):
    if is_blank(v):
        return True
    return is_numeric(v) and to_float(v) == 0


def has_content(v):
    return not is_empty_or_zero(v) if is_numeric(v) or is_blank(v) else True


def to_write(v):
    return None if v == "" else v


def vba_str(v):
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:.15g}"
    return str(v)


def for_filename(v):
    return re.sub(r'[\\/:*?"<>|]', "_", vba_str(v)).strip()


def _set_cell(ws, addr, value):
    """Write a value; a cell inside a merged area is written through the area's top-left cell."""
    cell = ws[addr]
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if addr in rng:
                cell = ws.cell(row=rng.min_row, column=rng.min_col)
                break
    cell.value = value
    if isinstance(value, (date, datetime)):
        cell.number_format = ETD_NUMBER_FORMAT


# --------------------------------------------------------------------------
# FILE HELPERS
# --------------------------------------------------------------------------
def collect_source_files(folder_or_files):
    if isinstance(folder_or_files, (str, os.PathLike)):
        folder_or_files = [folder_or_files]

    found = {}

    def add(f):
        if f.suffix.lower() == ".xlsm" and not f.name.startswith("~$"):
            found[os.path.normcase(str(f.resolve()))] = f

    for item in folder_or_files:
        p = Path(item)
        if p.is_dir():
            for f in p.glob("*.xlsm"):
                add(f)
            for sub in p.iterdir():
                if sub.is_dir() and sub.name.lower() in INPUT_SUBFOLDERS:
                    for f in sub.glob("*.xlsm"):
                        add(f)
        elif p.is_file():
            add(p)
    return sorted(found.values(), key=lambda x: (x.parent.name.lower(), x.name.lower()))


def _inside(path, folder):
    try:
        folder = os.path.abspath(folder)
        return os.path.commonpath([os.path.abspath(path), folder]) == folder
    except ValueError:
        return False


def _clear_output_folders(folders):
    targets, open_names = [], []
    for folder in folders:
        if not os.path.isdir(folder):
            continue
        for name in os.listdir(folder):
            full = os.path.join(folder, name)
            if not os.path.isfile(full):
                continue
            if name.startswith("~$"):
                open_names.append(name[2:])
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


def _copy_to_all(made):
    errors, warnings, copied = [], [], 0
    try:
        ALL_DIR.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return 0, [f"Could not create the ALL folder: {exc}"], warnings

    for src in sorted(made, key=lambda p: p.name.lower()):
        dest, n = ALL_DIR / src.name, 2
        while dest.exists():
            dest = ALL_DIR / f"{src.stem} ({n}){src.suffix}"
            n += 1
        try:
            shutil.copy2(src, dest)
            copied += 1
            if dest.name != src.name:
                warnings.append(f"{src.name}: name already exists in ALL, copied as {dest.name}")
        except OSError as exc:
            errors.append(f"{src.name}: could not copy to ALL: {exc}")
    return copied, errors, warnings


# --------------------------------------------------------------------------
# CELL MAPS
# --------------------------------------------------------------------------
def _common_cells(g):
    cells = {"A28": to_write(g("A61"))}
    for src, dst in SHIPPER_MAP.items():
        v = g(src)
        if v != "":
            cells[dst] = v
    for src, dst in COPY_OR_CLEAR_MAP.items():
        v = g(src)
        cells[dst] = v if has_content(v) else None
    cells["H2"] = to_write(g("E4"))
    cells["G30"] = cells["G51"] = to_write(g("M58"))
    cells["I30"] = to_write(g("O58"))
    return cells


def _multi_cells(g):
    cells = _common_cells(g)
    cells["B28"] = to_write(g("C58"))
    return cells


def _single_cells(g):
    cells = _common_cells(g)

    for src, dst, label in (("E40", "D30", "MAKER: "), ("E41", "D31", "MODEL: "),
                            ("E42", "D32", "CHASSIS: "), ("E43", "D33", "YEAR: ")):
        v = g(src)
        if v != "":
            cells[dst] = label + vba_str(v)

    fuel, cc = g("E44"), g("H44")
    if fuel != "" and cc != "" and is_numeric(cc) and to_float(cc) != 0:
        cells["D34"] = f"FUEL: {vba_str(fuel)} {vba_str(cc)} CC"

    d45 = g("D45")
    cells["D35"] = f"{vba_str(d45)} {vba_str(g('E45'))}" if d45 != "" else None

    d47 = g("D47")
    if d47 != "":
        cells["D37"] = "INNER CARGO" if d47 == "INNER CARGO" else " "

    for src_row, dst_row in zip(range(48, 53), range(38, 43)):
        m = g(f"M{src_row}")
        if m == "":
            continue
        if is_numeric(m) and to_float(m) != 0:
            cells[f"D{dst_row}"] = (f"{vba_str(g(f'D{src_row}'))} {vba_str(g(f'K{src_row}'))} "
                                    f"PC(S) {vba_str(m)} KG(S)")
        else:
            cells[f"D{dst_row}"] = " "
    return cells


# --------------------------------------------------------------------------
# PHASE 1: read + validate (STRICT ERROR REPORTING, NO SKIPPING)
# --------------------------------------------------------------------------
def _plan_file(path):
    path = Path(path)
    items = []
    book = CalamineWorkbook.from_path(str(path))
    names = book.sheet_names

    ivsi = [n for n in names if "ivsi" in n.lower()]
    if ivsi:
        kind, targets = SINGLE, ivsi
    else:
        kind, targets = MULTI, [n for n in names if "si" in n.lower()]
    if not targets:
        return [("ERROR", f"{path.name}: missing SI / IVSI sheet")]

    hint = path.parent.name.lower()

    for sheet in targets:
        label = f"{path.name} [{sheet}]"
        try:
            rows = book.get_sheet_by_name(sheet).to_python(skip_empty_area=False)
            g = lambda addr, rows=rows: get_cell(rows, addr)

            bl_no = g("E4")
            if is_blank(bl_no):
                items.append(("ERROR", f"{label}: cell E4 (BL no) is empty. Stage 3 may not have populated this file properly."))
                continue

            missing = [f"{lbl} ({addr})" for addr, lbl in REQUIRED[kind] if is_empty_or_zero(g(addr))]
            if missing:
                items.append(("ERROR", f"{label}: {', '.join(missing)} empty or 0, run Stage 3 first"))
                continue

            if kind == MULTI:
                cells = _multi_cells(g)
                name = f"{for_filename(bl_no)} {for_filename(g('C58'))}UT.xlsm"
            else:
                chassis = g("E42")
                if is_blank(chassis):
                    items.append(("ERROR", f"{label}: E42 (chassis) is empty, cannot name the file"))
                    continue
                cells = _single_cells(g)
                name = f"{for_filename(bl_no)} {for_filename(chassis)}.xlsm"

            job = {"kind": kind, "label": label, "name": name, "cells": cells, "warning": None}
            if hint in (SINGLE, MULTI) and hint != kind:
                job["warning"] = (f"{label}: sits in the {path.parent.name} folder but has the {kind}-unit layout, "
                                f"it was treated as {kind}")
            items.append(("JOB", job))
        except Exception as exc:
            items.append(("ERROR", f"{label}: {exc!r}"))
    return items


# --------------------------------------------------------------------------
# PHASE 2: write one output file
# --------------------------------------------------------------------------
def _write_output(template_bytes, cells, out_path):
    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    ws = wb[TEMPLATE_SHEET]
    for addr, v in cells.items():
        if isinstance(v, str):
            v = _ILLEGAL.sub("", v)
            if v == "":
                v = None
        _set_cell(ws, addr, v)
    wb.save(out_path)
    wb.close()
    return out_path


# --------------------------------------------------------------------------
# ENTRY POINT
# --------------------------------------------------------------------------
def process_file(folder_or_files, vessel=None):
    """
    folder_or_files: the FOR RPA folder (with SINGLE and MULTI inside), a single folder, or files.
    vessel: {"name", "voyage", "pol", "pod", "etd" (YYYY-MM-DD), "navis" (True / False)}. Required.
            The details are written into every BL file, and "navis" picks the NAVIS templates.
    """
    try:
        details, problem = validate_details(vessel)
        if problem:
            return {"ok": False, "message": problem}

        if not BL_ROOT.is_dir():
            return {"ok": False, "message": f"The folder 'BLC ALL MACRO' was not found in Downloads: {BL_ROOT}"}

        files = collect_source_files(folder_or_files)
        if not files:
            return {"ok": False, "message": "No .xlsm files found. Select the FOR RPA folder."}

        for f in files:
            if any(_inside(f, d) for d in CLEAR_DIRS):
                return {"ok": False, "message": f"{f.name} is inside the BL output folders. Select FOR RPA instead."}

        jobs, errors, warnings = [], [], []

        # phase 1: read and validate
        with ThreadPoolExecutor(max_workers=max(1, min(MAX_WORKERS, len(files)))) as pool:
            futures = {pool.submit(_plan_file, f): f for f in files}
            for fut in as_completed(futures):
                f = futures[fut]
                try:
                    items = fut.result()
                except Exception as exc:
                    errors.append(f"{f.name}: {exc!r}")
                    continue
                for status, payload in items:
                    if status == "JOB":
                        jobs.append(payload)
                        if payload["warning"]:
                            warnings.append(payload["warning"])
                    else:
                        errors.append(payload)

        # STRICT CHECK: if any error occurred in phase 1, stop completely
        if errors:
            return {
                "ok": False,
                "message": f"Stage 4 stopped: {len(errors)} error(s) found during file validation. {errors[0]}",
                "errors": errors,
            }

        if not jobs:
            return {"ok": False, "message": "No valid files to process in Stage 4."}

        # the templates for this vessel (NAVIS or standard), checked before anything is deleted
        templates = {}
        for kind in {j["kind"] for j in jobs}:
            path, tried = find_template(kind, details["navis"])
            if path is None:
                errors.append(f"Template not found: {TEMPLATE_FILES[(kind, details['navis'])]} "
                            f"(looked in {tried[0].parent})")
                continue
            data = path.read_bytes()
            problem = _template_problem(data, path.name)
            if problem:
                errors.append(problem)
            else:
                templates[kind] = data
        if errors:
            return {"ok": False, "message": f"Stage 4 stopped: {errors[0]}", "errors": errors}

        # the vessel details go into every BL file
        vessel_cells = _vessel_cells(details)
        for j in jobs:
            j["cells"].update(vessel_cells)

        # delete the old output
        deleted, problem = _clear_output_folders(CLEAR_DIRS)
        if problem:
            return {"ok": False, "message": problem}

        # assign output names
        jobs.sort(key=lambda j: j["label"].lower())
        used = set()
        for j in jobs:
            out_dir = OUTPUT_DIRS[j["kind"]]
            out_dir.mkdir(parents=True, exist_ok=True)
            stem, ext = os.path.splitext(j["name"])
            path, n = out_dir / j["name"], 2
            while str(path).lower() in used:
                path = out_dir / f"{stem} ({n}){ext}"
                n += 1
            if path.name != j["name"]:
                warnings.append(f"{j['label']}: name already used in this run, saved as {path.name}")
            used.add(str(path).lower())
            j["out_path"] = str(path)

        # phase 2: write
        created = {SINGLE: [], MULTI: []}
        made = []
        Executor = ProcessPoolExecutor if USE_PROCESSES else ThreadPoolExecutor
        with Executor(max_workers=max(1, min(MAX_WORKERS, len(jobs)))) as pool:
            futures = {pool.submit(_write_output, templates[j["kind"]], j["cells"], j["out_path"]): j
                for j in jobs}
            for fut in as_completed(futures):
                j = futures[fut]
                try:
                    fut.result()
                    created[j["kind"]].append(Path(j["out_path"]).name)
                    made.append(Path(j["out_path"]))
                except Exception as exc:
                    errors.append(f"{j['label']}: could not write {Path(j['out_path']).name}: {exc!r}")

        if errors:
            return {
                "ok": False,
                "message": f"Stage 4 stopped: {len(errors)} error(s) occurred while writing outputs. {errors[0]}",
                "errors": errors,
            }

        # copy created files to ALL
        copied, copy_errors, copy_warnings = _copy_to_all(made)
        if copy_errors:
            return {
                "ok": False,
                "message": f"Stage 4 stopped: Failed copying to ALL folder. {copy_errors[0]}",
                "errors": copy_errors,
            }
        warnings.extend(copy_warnings)

        total = len(created[SINGLE]) + len(created[MULTI])
        message = (f"Stage 4: {total} BL file(s) created "
                f"({len(created[SINGLE])} single in SINGLE, {len(created[MULTI])} multi in MULTI), "
                f"{copied} copied to ALL. Vessel: {_describe(details)}.")
        if deleted:
            message += f" {deleted} old file(s) deleted first."
        if warnings:
            message += f" {len(warnings)} warning(s)."

        return {
            "ok": True,
            "message": message,
            "single": len(created[SINGLE]),
            "multi": len(created[MULTI]),
            "copied_to_all": copied,
            "created": created[SINGLE] + created[MULTI],
            "deleted_old": deleted,
            "details": {**{k: details[k] for k in ("name", "voyage", "pol", "pod", "navis")},
                        "etd": details["etd"].isoformat()},
            "warnings": warnings or None,
            "output_dirs": {**{k: str(v) for k, v in OUTPUT_DIRS.items()}, "all": str(ALL_DIR)},
        }
    except Exception as exc:
        return {"ok": False, "message": f"BL creation failed: {exc!r}"}










# import io
# import os
# import re
# import shutil
# from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
# from pathlib import Path

# from openpyxl import load_workbook
# from python_calamine import CalamineWorkbook

# # --------------------------------------------------------------------------
# # CONFIG
# # --------------------------------------------------------------------------
# SINGLE, MULTI = "single", "multi"

# # Dynamic path resolution
# BL_ROOT = Path(os.path.expanduser("~")) / "Downloads" / "BLC ALL MACRO"
# TEMPLATES = {
#     SINGLE: BL_ROOT / "SO TEMPLATES" / "DINAS FORM single.xlsm",
#     MULTI: BL_ROOT / "SO TEMPLATES" / "DINAS FORM multi.xlsm",
# }

# BL_OUT = BL_ROOT / "WIP BL" / "BL"
# OUTPUT_DIRS = {
#     SINGLE: BL_OUT / "SINGLE",
#     MULTI: BL_OUT / "MULTI",
# }
# ALL_DIR = BL_OUT / "ALL"
# CLEAR_DIRS = [BL_OUT / "SINGLE", BL_OUT / "MULTI", ALL_DIR]

# INPUT_SUBFOLDERS = ("single", "multi")
# TEMPLATE_SHEET = "Table 1"
# MAX_WORKERS = 8
# USE_PROCESSES = False

# # Cells that must hold a real, non-zero value, otherwise Stage 3 has not been run
# REQUIRED = {
#     MULTI: (("C58", "units"), ("M58", "weight"), ("O58", "m3")),
#     SINGLE: (("M58", "weight"), ("O58", "m3")),
# }

# SHIPPER_MAP = {"D6": "A3", "D7": "A4", "D8": "A5"}
# COPY_OR_CLEAR_MAP = {
#     "D9": "A6",
#     "D14": "A11", "D15": "A12", "D16": "A13",
#     "D22": "A17", "D23": "A18", "D24": "A19",
# }

# _EXCEL_FILE = re.compile(r"\.(xlsx|xlsm|xlsb|xls)(\.\d+)?$", re.IGNORECASE)
# _ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


# # --------------------------------------------------------------------------
# # CELL HELPERS
# # --------------------------------------------------------------------------
# def get_cell(rows, address):
#     m = re.match(r"^([A-Za-z]+)(\d+)$", address)
#     col = 0
#     for ch in m.group(1).upper():
#         col = col * 26 + (ord(ch) - 64)
#     try:
#         return rows[int(m.group(2)) - 1][col - 1]
#     except IndexError:
#         return ""


# def is_blank(v):
#     return v is None or (isinstance(v, str) and v.strip() == "")


# def is_numeric(v):
#     if isinstance(v, bool):
#         return False
#     if isinstance(v, (int, float)):
#         return True
#     if isinstance(v, str):
#         try:
#             float(v.strip().replace(",", ""))
#             return True
#         except ValueError:
#             return False
#     return False


# def to_float(v):
#     return float(v.strip().replace(",", "")) if isinstance(v, str) else float(v)


# def is_empty_or_zero(v):
#     if is_blank(v):
#         return True
#     return is_numeric(v) and to_float(v) == 0


# def has_content(v):
#     return not is_empty_or_zero(v) if is_numeric(v) or is_blank(v) else True


# def to_write(v):
#     return None if v == "" else v


# def vba_str(v):
#     if v is None or v == "":
#         return ""
#     if isinstance(v, bool):
#         return "True" if v else "False"
#     if isinstance(v, int):
#         return str(v)
#     if isinstance(v, float):
#         return str(int(v)) if v.is_integer() else f"{v:.15g}"
#     return str(v)


# def for_filename(v):
#     return re.sub(r'[\\/:*?"<>|]', "_", vba_str(v)).strip()


# # --------------------------------------------------------------------------
# # FILE HELPERS
# # --------------------------------------------------------------------------
# def collect_source_files(folder_or_files):
#     if isinstance(folder_or_files, (str, os.PathLike)):
#         folder_or_files = [folder_or_files]

#     found = {}

#     def add(f):
#         if f.suffix.lower() == ".xlsm" and not f.name.startswith("~$"):
#             found[os.path.normcase(str(f.resolve()))] = f

#     for item in folder_or_files:
#         p = Path(item)
#         if p.is_dir():
#             for f in p.glob("*.xlsm"):
#                 add(f)
#             for sub in p.iterdir():
#                 if sub.is_dir() and sub.name.lower() in INPUT_SUBFOLDERS:
#                     for f in sub.glob("*.xlsm"):
#                         add(f)
#         elif p.is_file():
#             add(p)
#     return sorted(found.values(), key=lambda x: (x.parent.name.lower(), x.name.lower()))


# def _inside(path, folder):
#     try:
#         folder = os.path.abspath(folder)
#         return os.path.commonpath([os.path.abspath(path), folder]) == folder
#     except ValueError:
#         return False


# def _clear_output_folders(folders):
#     targets, open_names = [], []
#     for folder in folders:
#         if not os.path.isdir(folder):
#             continue
#         for name in os.listdir(folder):
#             full = os.path.join(folder, name)
#             if not os.path.isfile(full):
#                 continue
#             if name.startswith("~$"):
#                 open_names.append(name[2:])
#             elif _EXCEL_FILE.search(name):
#                 targets.append(full)

#     if open_names:
#         shown = ", ".join(open_names[:5]) + (" ..." if len(open_names) > 5 else "")
#         return 0, f"Close these files in Excel first, nothing was deleted: {shown}"

#     deleted, failed = 0, []
#     for path in targets:
#         try:
#             os.remove(path)
#             deleted += 1
#         except OSError:
#             failed.append(os.path.basename(path))
#     if failed:
#         shown = ", ".join(failed[:5]) + (" ..." if len(failed) > 5 else "")
#         return deleted, f"Could not delete {len(failed)} old file(s), are they open in Excel? {shown}"
#     return deleted, None


# def _copy_to_all(made):
#     errors, warnings, copied = [], [], 0
#     try:
#         ALL_DIR.mkdir(parents=True, exist_ok=True)
#     except OSError as exc:
#         return 0, [f"Could not create the ALL folder: {exc}"], warnings

#     for src in sorted(made, key=lambda p: p.name.lower()):
#         dest, n = ALL_DIR / src.name, 2
#         while dest.exists():
#             dest = ALL_DIR / f"{src.stem} ({n}){src.suffix}"
#             n += 1
#         try:
#             shutil.copy2(src, dest)
#             copied += 1
#             if dest.name != src.name:
#                 warnings.append(f"{src.name}: name already exists in ALL, copied as {dest.name}")
#         except OSError as exc:
#             errors.append(f"{src.name}: could not copy to ALL: {exc}")
#     return copied, errors, warnings


# # --------------------------------------------------------------------------
# # CELL MAPS
# # --------------------------------------------------------------------------
# def _common_cells(g):
#     cells = {"A28": to_write(g("A61"))}
#     for src, dst in SHIPPER_MAP.items():
#         v = g(src)
#         if v != "":
#             cells[dst] = v
#     for src, dst in COPY_OR_CLEAR_MAP.items():
#         v = g(src)
#         cells[dst] = v if has_content(v) else None
#     cells["H2"] = to_write(g("E4"))
#     cells["G30"] = cells["G51"] = to_write(g("M58"))
#     cells["I30"] = to_write(g("O58"))
#     return cells


# def _multi_cells(g):
#     cells = _common_cells(g)
#     cells["B28"] = to_write(g("C58"))
#     return cells


# def _single_cells(g):
#     cells = _common_cells(g)

#     for src, dst, label in (("E40", "D30", "MAKER: "), ("E41", "D31", "MODEL: "),
#                             ("E42", "D32", "CHASSIS: "), ("E43", "D33", "YEAR: ")):
#         v = g(src)
#         if v != "":
#             cells[dst] = label + vba_str(v)

#     fuel, cc = g("E44"), g("H44")
#     if fuel != "" and cc != "" and is_numeric(cc) and to_float(cc) != 0:
#         cells["D34"] = f"FUEL: {vba_str(fuel)} {vba_str(cc)} CC"

#     d45 = g("D45")
#     cells["D35"] = f"{vba_str(d45)} {vba_str(g('E45'))}" if d45 != "" else None

#     d47 = g("D47")
#     if d47 != "":
#         cells["D37"] = "INNER CARGO" if d47 == "INNER CARGO" else " "

#     for src_row, dst_row in zip(range(48, 53), range(38, 43)):
#         m = g(f"M{src_row}")
#         if m == "":
#             continue
#         if is_numeric(m) and to_float(m) != 0:
#             cells[f"D{dst_row}"] = (f"{vba_str(g(f'D{src_row}'))} {vba_str(g(f'K{src_row}'))} "
#                                     f"PC(S) {vba_str(m)} KG(S)")
#         else:
#             cells[f"D{dst_row}"] = " "
#     return cells


# # --------------------------------------------------------------------------
# # PHASE 1: read + validate (STRICT ERROR REPORTING, NO SKIPPING)
# # --------------------------------------------------------------------------
# def _plan_file(path):
#     path = Path(path)
#     items = []
#     book = CalamineWorkbook.from_path(str(path))
#     names = book.sheet_names

#     ivsi = [n for n in names if "ivsi" in n.lower()]
#     if ivsi:
#         kind, targets = SINGLE, ivsi
#     else:
#         kind, targets = MULTI, [n for n in names if "si" in n.lower()]
#     if not targets:
#         return [("ERROR", f"{path.name}: missing SI / IVSI sheet")]

#     hint = path.parent.name.lower()

#     for sheet in targets:
#         label = f"{path.name} [{sheet}]"
#         try:
#             rows = book.get_sheet_by_name(sheet).to_python(skip_empty_area=False)
#             g = lambda addr, rows=rows: get_cell(rows, addr)

#             bl_no = g("E4")
#             # CHANGED: Treat missing E4 (BL no) as a hard ERROR instead of a SKIP
#             if is_blank(bl_no):
#                 items.append(("ERROR", f"{label}: cell E4 (BL no) is empty. Stage 3 may not have populated this file properly."))
#                 continue

#             missing = [f"{lbl} ({addr})" for addr, lbl in REQUIRED[kind] if is_empty_or_zero(g(addr))]
#             if missing:
#                 items.append(("ERROR", f"{label}: {', '.join(missing)} empty or 0, run Stage 3 first"))
#                 continue

#             if kind == MULTI:
#                 cells = _multi_cells(g)
#                 name = f"{for_filename(bl_no)} {for_filename(g('C58'))}UT.xlsm"
#             else:
#                 chassis = g("E42")
#                 if is_blank(chassis):
#                     items.append(("ERROR", f"{label}: E42 (chassis) is empty, cannot name the file"))
#                     continue
#                 cells = _single_cells(g)
#                 name = f"{for_filename(bl_no)} {for_filename(chassis)}.xlsm"

#             job = {"kind": kind, "label": label, "name": name, "cells": cells, "warning": None}
#             if hint in (SINGLE, MULTI) and hint != kind:
#                 job["warning"] = (f"{label}: sits in the {path.parent.name} folder but has the {kind}-unit layout, "
#                                 f"it was treated as {kind}")
#             items.append(("JOB", job))
#         except Exception as exc:
#             items.append(("ERROR", f"{label}: {exc!r}"))
#     return items


# # --------------------------------------------------------------------------
# # PHASE 2: write one output file
# # --------------------------------------------------------------------------
# def _write_output(template_bytes, cells, out_path):
#     wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
#     ws = wb[TEMPLATE_SHEET]
#     for addr, v in cells.items():
#         if isinstance(v, str):
#             v = _ILLEGAL.sub("", v)
#             if v == "":
#                 v = None
#         ws[addr].value = v
#     wb.save(out_path)
#     wb.close()
#     return out_path


# # --------------------------------------------------------------------------
# # ENTRY POINT
# # --------------------------------------------------------------------------
# def process_file(folder_or_files):
#     try:
#         if not BL_ROOT.is_dir():
#             return {"ok": False, "message": f"The folder 'BLC ALL MACRO' was not found in Downloads: {BL_ROOT}"}

#         files = collect_source_files(folder_or_files)
#         if not files:
#             return {"ok": False, "message": "No .xlsm files found. Select the FOR RPA folder."}

#         for f in files:
#             if any(_inside(f, d) for d in CLEAR_DIRS):
#                 return {"ok": False, "message": f"{f.name} is inside the BL output folders. Select FOR RPA instead."}

#         jobs, errors, warnings = [], [], []

#         # phase 1: read and validate
#         with ThreadPoolExecutor(max_workers=max(1, min(MAX_WORKERS, len(files)))) as pool:
#             futures = {pool.submit(_plan_file, f): f for f in files}
#             for fut in as_completed(futures):
#                 f = futures[fut]
#                 try:
#                     items = fut.result()
#                 except Exception as exc:
#                     errors.append(f"{f.name}: {exc!r}")
#                     continue
#                 for status, payload in items:
#                     if status == "JOB":
#                         jobs.append(payload)
#                         if payload["warning"]:
#                             warnings.append(payload["warning"])
#                     else:
#                         errors.append(payload)

#         # STRICT CHECK: If any error occurred in Phase 1, stop execution completely!
#         if errors:
#             return {
#                 "ok": False,
#                 "message": f"Stage 4 stopped: {len(errors)} error(s) found during file validation. {errors[0]}",
#                 "errors": errors,
#             }

#         templates = {}
#         for kind in {j["kind"] for j in jobs}:
#             tpl = TEMPLATES[kind]
#             if tpl.is_file():
#                 templates[kind] = tpl.read_bytes()
#             else:
#                 errors.append(f"Template not found: {tpl}")
        
#         if errors:
#             return {"ok": False, "message": f"Stage 4 stopped: {errors[0]}", "errors": errors}

#         if not jobs:
#             return {"ok": False, "message": "No valid files to process in Stage 4."}

#         # delete the old output
#         deleted, problem = _clear_output_folders(CLEAR_DIRS)
#         if problem:
#             return {"ok": False, "message": problem}

#         # assign output names
#         jobs.sort(key=lambda j: j["label"].lower())
#         used = set()
#         for j in jobs:
#             out_dir = OUTPUT_DIRS[j["kind"]]
#             out_dir.mkdir(parents=True, exist_ok=True)
#             stem, ext = os.path.splitext(j["name"])
#             path, n = out_dir / j["name"], 2
#             while str(path).lower() in used:
#                 path = out_dir / f"{stem} ({n}){ext}"
#                 n += 1
#             if path.name != j["name"]:
#                 warnings.append(f"{j['label']}: name already used in this run, saved as {path.name}")
#             used.add(str(path).lower())
#             j["out_path"] = str(path)

#         # phase 2: write
#         created = {SINGLE: [], MULTI: []}
#         made = []
#         Executor = ProcessPoolExecutor if USE_PROCESSES else ThreadPoolExecutor
#         with Executor(max_workers=max(1, min(MAX_WORKERS, len(jobs)))) as pool:
#             futures = {pool.submit(_write_output, templates[j["kind"]], j["cells"], j["out_path"]): j
#                     for j in jobs}
#             for fut in as_completed(futures):
#                 j = futures[fut]
#                 try:
#                     fut.result()
#                     created[j["kind"]].append(Path(j["out_path"]).name)
#                     made.append(Path(j["out_path"]))
#                 except Exception as exc:
#                     errors.append(f"{j['label']}: could not write {Path(j['out_path']).name}: {exc!r}")

#         if errors:
#             return {
#                 "ok": False,
#                 "message": f"Stage 4 stopped: {len(errors)} error(s) occurred while writing outputs. {errors[0]}",
#                 "errors": errors,
#             }

#         # copy created files to ALL
#         copied, copy_errors, copy_warnings = _copy_to_all(made)
#         if copy_errors:
#             return {
#                 "ok": False,
#                 "message": f"Stage 4 stopped: Failed copying to ALL folder. {copy_errors[0]}",
#                 "errors": copy_errors,
#             }
#         warnings.extend(copy_warnings)

#         total = len(created[SINGLE]) + len(created[MULTI])
#         message = (f"Stage 4: {total} BL file(s) created "
#                 f"({len(created[SINGLE])} single in SINGLE, {len(created[MULTI])} multi in MULTI), "
#                 f"{copied} copied to ALL.")
#         if deleted:
#             message += f" {deleted} old file(s) deleted first."
#         if warnings:
#             message += f" {len(warnings)} warning(s)."

#         return {
#             "ok": True,
#             "message": message,
#             "single": len(created[SINGLE]),
#             "multi": len(created[MULTI]),
#             "copied_to_all": copied,
#             "created": created[SINGLE] + created[MULTI],
#             "deleted_old": deleted,
#             "warnings": warnings or None,
#             "output_dirs": {**{k: str(v) for k, v in OUTPUT_DIRS.items()}, "all": str(ALL_DIR)},
#         }
#     except Exception as exc:
#         return {"ok": False, "message": f"BL creation failed: {exc!r}"}


