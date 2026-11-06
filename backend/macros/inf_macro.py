import os
import posixpath
import re
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from xml.sax.saxutils import escape, unescape

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
TARGET_SHEET = "Table 1"

# Cell -> value written into it. A text value is written as is (" " is one blank space),
# a number is written as a number, and None or "" empties the cell.
# Add or change lines here the way you used to comment / uncomment lines in the VBA, for example:
#   "B23": "AMUR",   "E23": 38,   "C21": None
CELL_VALUES = {
    "A18": " ",
    "A19": " ",
}

MAX_WORKERS = 8
_EXCEL_ZIP = (".xlsx", ".xlsm")          # the formats that can be patched directly (.xls / .xlsb cannot)
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


# --------------------------------------------------------------------------
# SHEET XML EDITING (only the one worksheet part is touched, everything else is copied byte for byte)
# --------------------------------------------------------------------------
_REF = re.compile(r"^([A-Za-z]+)(\d+)$")
_ROW_TAG = re.compile(r"<row\b[^>]*>")
_ROW_NUM = re.compile(r'\sr="(\d+)"')
_CELL_COL = re.compile(r'<c\b[^>]*?\sr="([A-Za-z]+)\d+"')


def _col_index(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def _cell_xml(ref, style, value):
    s = f' s="{style}"' if style else ""
    if value is None or value == "":
        return f'<c r="{ref}"{s}/>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"{s}><v>{value!r}</v></c>'
    text = escape(_ILLEGAL.sub("", str(value)))
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def _set_cell(xml, ref, value):
    m = _REF.match(ref)
    if not m:
        raise ValueError(f"bad cell address {ref!r}")
    letters, row_num = m.group(1).upper(), int(m.group(2))
    col = _col_index(letters)
    ref = f"{letters}{row_num}"

    sd = re.search(r"<sheetData\b[^>]*>", xml)
    if sd is None:
        raise ValueError("the sheet has no data section")
    if sd.group(0).endswith("/>"):                              # empty sheet: <sheetData/>
        row = f'<row r="{row_num}">{_cell_xml(ref, None, value)}</row>'
        return xml[:sd.start()] + sd.group(0)[:-2] + ">" + row + "</sheetData>" + xml[sd.end():]

    # find the row, or the place where it has to be inserted
    target, insert_at = None, None
    for rm in _ROW_TAG.finditer(xml, sd.end()):
        n = _ROW_NUM.search(rm.group(0))
        if n is None:
            raise ValueError("rows without a number are not supported")
        r = int(n.group(1))
        if r == row_num:
            target = rm
            break
        if r > row_num:
            insert_at = rm.start()
            break

    if target is None:                                          # the row does not exist yet
        row = f'<row r="{row_num}">{_cell_xml(ref, None, value)}</row>'
        if insert_at is None:
            insert_at = xml.index("</sheetData>", sd.end())
        return xml[:insert_at] + row + xml[insert_at:]

    tag = target.group(0)
    if tag.endswith("/>"):                                      # empty row: <row r="18" .../>
        new_tag = re.sub(r'\sspans="[^"]*"', "", tag[:-2]) + ">"
        return xml[:target.start()] + new_tag + _cell_xml(ref, None, value) + "</row>" + xml[target.end():]

    body_start = target.end()
    body_end = xml.index("</row>", body_start)
    body = xml[body_start:body_end]

    cm = re.search(r'<c\b[^>]*?\sr="%s"[^>]*?(?:/>|>.*?</c>)' % ref, body, re.S)
    if cm:                                                      # the cell exists: replace it, keep its style
        old = cm.group(0)
        if "<f" in old:
            raise ValueError(f"{ref} contains a formula, so it was not changed")
        style = re.search(r'\ss="(\d+)"', re.match(r"<c\b[^>]*>", old).group(0))
        new_body = body[:cm.start()] + _cell_xml(ref, style.group(1) if style else None, value) + body[cm.end():]
        return xml[:body_start] + new_body + xml[body_end:]

    pos = len(body)                                             # the cell is missing: insert it in column order
    for c in _CELL_COL.finditer(body):
        if _col_index(c.group(1)) > col:
            pos = c.start()
            break
    new_body = body[:pos] + _cell_xml(ref, None, value) + body[pos:]
    new_tag = re.sub(r'\sspans="[^"]*"', "", tag)
    return xml[:target.start()] + new_tag + new_body + xml[body_end:]


# --------------------------------------------------------------------------
# WORKBOOK LOOKUP + FILE PATCHING
# --------------------------------------------------------------------------
def _sheet_part(zf, sheet_name):
    """Name of the zip entry that holds the worksheet called sheet_name."""
    wanted = sheet_name.strip().lower()
    wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    rid = None
    for tag in re.findall(r"<sheet\b[^>]*>", wb_xml):
        name = re.search(r'\bname="([^"]*)"', tag)
        if name and unescape(name.group(1), {"&quot;": '"', "&apos;": "'"}).strip().lower() == wanted:
            m = re.search(r'\s\w+:id="([^"]*)"', tag)
            rid = m.group(1) if m else None
            break
    if rid is None:
        raise ValueError(f"sheet '{sheet_name}' not found")

    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    for tag in re.findall(r"<Relationship\b[^>]*>", rels):
        if re.search(r'\bId="%s"' % re.escape(rid), tag):
            target = re.search(r'\bTarget="([^"]*)"', tag).group(1)
            return target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
    raise ValueError(f"sheet '{sheet_name}' could not be located inside the file")


def patch_file(path):
    """Apply CELL_VALUES to TARGET_SHEET of one workbook, in place. -> 'updated' or 'unchanged'."""
    path = Path(path)
    if path.suffix.lower() not in _EXCEL_ZIP:
        raise ValueError("only .xlsx and .xlsm files are supported")

    tmp = path.with_name(path.name + ".tmp")
    try:
        with zipfile.ZipFile(path) as zin:
            part = _sheet_part(zin, TARGET_SHEET)
            original = zin.read(part).decode("utf-8")
            xml = original
            for ref, value in CELL_VALUES.items():
                xml = _set_cell(xml, ref, value)
            if xml == original:
                return "unchanged"

            with zipfile.ZipFile(tmp, "w") as zout:
                for info in zin.infolist():
                    data = xml.encode("utf-8") if info.filename == part else zin.read(info.filename)
                    new = zipfile.ZipInfo(info.filename, info.date_time)
                    new.compress_type = info.compress_type
                    new.external_attr = info.external_attr
                    new.create_system = info.create_system
                    zout.writestr(new, data)

        os.replace(tmp, path)                       # the zip is closed here, so the swap also works on Windows
        return "updated"
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass


def _run_one(path):
    try:
        return "ok", path.name, patch_file(path)
    except PermissionError:
        return "error", path.name, "could not be saved, is it open in Excel? Close it and run again"
    except Exception as exc:
        return "error", path.name, str(exc)


# --------------------------------------------------------------------------
# ENTRY POINT (called by Api.run_macro)
# --------------------------------------------------------------------------
def _collect_files(folder_or_files):
    if not folder_or_files:
        return [], []
    items = [folder_or_files] if isinstance(folder_or_files, (str, os.PathLike)) else list(folder_or_files)

    files, problems = [], []
    for item in items:
        p = Path(os.fspath(item))
        if p.is_dir():                              # only the files directly inside the folder
            files += [f for f in sorted(p.iterdir())
                      if f.is_file() and f.suffix.lower() in _EXCEL_ZIP and not f.name.startswith("~$")]
        elif p.is_file():
            if p.suffix.lower() in _EXCEL_ZIP:
                files.append(p)
            else:
                problems.append(f"{p.name}: only .xlsx and .xlsm files are supported")
        else:
            problems.append(f"Not found: {p}")
    return list(dict.fromkeys(files)), problems


def process_remove_zero(folder_or_files):
    """
    Writes CELL_VALUES (A18 and A19 = one blank space) into the "Table 1" sheet of every selected file
    and saves each file in place. Only the worksheet part inside the file is rewritten, so macros,
    formatting, images and everything else stay exactly as they were.
    """
    files, problems = _collect_files(folder_or_files)
    if not files:
        return {"ok": False, "message": "No .xlsm or .xlsx files found in the selection.",
                "errors": problems or None}

    updated, unchanged, errors = 0, 0, list(problems)
    with ThreadPoolExecutor(max_workers=max(1, min(MAX_WORKERS, len(files)))) as pool:
        for status, name, detail in pool.map(_run_one, files):
            if status == "ok":
                if detail == "updated":
                    updated += 1
                else:
                    unchanged += 1
            else:
                errors.append(f"{name}: {detail}")

    message = f"Stage 6: {updated} file(s) updated"
    if unchanged:
        message += f", {unchanged} already clean"
    message += "."
    if errors:
        message += f" {len(errors)} error(s)."
    return {
        "ok": (updated + unchanged) > 0,
        "message": message,
        "updated": updated,
        "unchanged": unchanged,
        "errors": errors or None,
    }