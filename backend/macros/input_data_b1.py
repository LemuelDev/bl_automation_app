import gc
import os
import queue
import re
import shutil
import subprocess
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from python_calamine import CalamineWorkbook

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
MAX_EXCEL_INSTANCES = 3        # separate hidden Excel processes working at the same time
RESTART_EVERY = 20             # start a fresh Excel after this many files (keeps Excel stable on big batches)
JOB_TIMEOUT = 180              # seconds one file may take; after that its Excel is killed (a hidden dialog would block forever)
OPEN_TRIES = 4                 # attempts to open a file
SAVE_TRIES = 8                 # attempts to save / put a finished file back (antivirus, indexer or sync locks)
RETRY_ROUNDS = 2               # extra passes over the files that failed, one Excel at a time
RETRY_PAUSE = 4                # seconds to wait before a retry pass (multiplied by the round number)

# The MULTI template's BASE event macro shows a MsgBox above 50 units, which would freeze a hidden Excel.
MAX_UNITS = {"multi": 50}

SINGLE, MULTI = "single", "multi"
SI_SHEET = "SI"
ATCH_SHEET = "ATCH"
IVSI_SHEET = "IVSI"
FLAG_SHEET = "Sheet2"

XL_SHEET_VISIBLE = -1
XL_SHEET_HIDDEN = 0
XL_CALC_AUTOMATIC = -4105
MSO_AUTOMATION_SECURITY_LOW = 1

_ATCH2 = re.compile(r"^ATCH\s*\(2\)$", re.IGNORECASE)
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


# --------------------------------------------------------------------------
# SMALL HELPERS
# --------------------------------------------------------------------------
def _blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def _to_int(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def _is_zero(v):
    if _blank(v):
        return True
    try:
        return float(v) == 0
    except (TypeError, ValueError):
        return False


def _find(names, wanted):
    wanted = wanted.strip().upper()
    return next((n for n in names if n.strip().upper() == wanted), None)


def _as_list(value):
    if not value:
        return []
    if isinstance(value, (str, os.PathLike)):
        return [os.fspath(value)]
    return [os.fspath(v) for v in value]


def _backoff(attempt):
    return min(0.5 * 2 ** attempt, 8)


def _retry(fn, tries):
    """Run fn, and try again a few times when it fails (Excel busy, file briefly locked)."""
    last = None
    for attempt in range(tries):
        try:
            return fn()
        except Exception as exc:
            last = exc
            if attempt < tries - 1:
                time.sleep(_backoff(attempt))
    raise last


def _clean_error(exc):
    """The readable part of a COM error (they are long tuples otherwise)."""
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) > 2 and args[2][2]:
        text = str(args[2][2])
    else:
        text = str(exc)
    return re.sub(r"\s+", " ", text).strip()[:300]


# --------------------------------------------------------------------------
# STEP 1: fast pre-check with calamine (no Excel needed)
# --------------------------------------------------------------------------
def _kind_for(count):
    return SINGLE if count == 1 else MULTI


def _structure_problem(kind, names, count):
    has_ivsi = _find(names, IVSI_SHEET) is not None
    has_flag = _find(names, FLAG_SHEET) is not None
    has_si = _find(names, SI_SHEET) is not None
    has_atch = _find(names, ATCH_SHEET) is not None or any(_ATCH2.match(n.strip()) for n in names)

    if kind == SINGLE:
        if not has_ivsi:
            return "1 unit, but there is no IVSI sheet"
        if not has_flag:
            return f"1 unit, but sheet '{FLAG_SHEET}' is missing"
    else:
        if has_ivsi:
            return f"{count} units, but the file has an IVSI sheet"
        if not (has_si and has_atch):
            return f"{count} units, but the SI / ATCH sheets are missing"
    return None


def _precheck(path, sheet_name):
    """-> (status, base_sheet_name, unit_count, kind, message)"""
    book = CalamineWorkbook.from_path(path)
    names = book.sheet_names

    base = _find(names, sheet_name)
    if base is None:
        return "error", None, 0, None, f"sheet '{sheet_name}' not found"

    rows = book.get_sheet_by_name(base).to_python(skip_empty_area=False)
    count = sum(1 for r in rows if len(r) > 3 and not _blank(r[3])) - 1
    if count < 1:
        return "error", base, count, None, "no units found in column D"

    kind = _kind_for(count)
    limit = MAX_UNITS.get(kind)
    if limit and count > limit:
        return "error", base, count, kind, f"{count} units is more than the template allows ({limit})"

    problem = _structure_problem(kind, names, count)
    if problem:
        return "error", base, count, kind, problem
    return "ready", base, count, kind, None


# --------------------------------------------------------------------------
# EXCEL PROCESSES: only the ones we start are ever touched
# --------------------------------------------------------------------------
_PID_FILE = Path(tempfile.gettempdir()) / "bl_stage3_excel.pids"
_pid_lock = threading.Lock()


def _is_excel(pid):
    try:
        out = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                             capture_output=True, text=True, creationflags=_NO_WINDOW).stdout
        return "EXCEL.EXE" in out.upper()
    except Exception:
        return False


def _kill_pid(pid):
    try:
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True, creationflags=_NO_WINDOW)
    except Exception:
        pass


def _read_pids():
    try:
        return {int(x) for x in _PID_FILE.read_text().split()}
    except Exception:
        return set()


def _write_pids(pids):
    try:
        _PID_FILE.write_text(" ".join(str(p) for p in sorted(pids)))
    except Exception:
        pass


def _register_pid(pid):
    if pid:
        with _pid_lock:
            _write_pids(_read_pids() | {pid})


def _unregister_pid(pid):
    if pid:
        with _pid_lock:
            _write_pids(_read_pids() - {pid})


def _kill_leftovers():
    """An earlier run that crashed may have left its hidden Excel behind: end exactly those, nothing else."""
    with _pid_lock:
        for pid in _read_pids():
            if _is_excel(pid):
                _kill_pid(pid)
        _write_pids(set())


def _com_init():
    import pythoncom
    pythoncom.CoInitialize()


def _com_done():
    import pythoncom
    pythoncom.CoUninitialize()


class _Excel:
    """One hidden Excel process owned by this stage."""

    def __init__(self):
        import win32com.client
        self.killed = False
        self.pid = None
        self.app = win32com.client.DispatchEx("Excel.Application")
        try:
            try:
                import win32process
                self.pid = win32process.GetWindowThreadProcessId(self.app.Hwnd)[1]
            except Exception:
                self.pid = None
            _register_pid(self.pid)
            a = self.app
            a.Visible = False
            a.DisplayAlerts = False
            a.ScreenUpdating = False
            a.AskToUpdateLinks = False
            a.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        except Exception:
            self.quit()
            raise

    def arm(self, seconds):
        """Watchdog: kill this Excel if one file takes longer than `seconds`."""
        timer = threading.Timer(seconds, self.kill)
        timer.daemon = True
        timer.start()
        return timer

    def kill(self):
        self.killed = True
        if self.pid:
            _kill_pid(self.pid)

    def quit(self):
        app, self.app = self.app, None
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        del app
        gc.collect()
        if self.pid:
            deadline = time.time() + 5
            while time.time() < deadline and _is_excel(self.pid):
                time.sleep(0.5)
            if _is_excel(self.pid):
                _kill_pid(self.pid)
            _unregister_pid(self.pid)


# --------------------------------------------------------------------------
# STEP 2: one file through Excel
# --------------------------------------------------------------------------
def _sheet_names(wb):
    return [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]


def _trigger(wb, ws, count, kind):
    """Make the BASE event macro run for real, whatever state the file is in."""
    ws.Range("B1").Value = 0                         # 1. reset: both macros go back to their "not processed" state
    if kind == SINGLE:
        flags = wb.Worksheets(FLAG_SHEET)
        flags.Range("BB1").Value = -999              #    the single macro exits when B1 equals Sheet2!BB1
        flags.Range("BA1").Value = ""                #    and skips its work when BA1 already holds "a"
    ws.Range("B1").Value = count                     # 2. the real change: the event macro runs (synchronously)


def _verify(wb, count, kind):
    """Proof that the macro did its work. Returns warnings. (SI!C58 is 0 in the template, so only the macro can fill it.)"""
    names = _sheet_names(wb)
    warnings = []
    if kind == SINGLE:
        flag_name = _find(names, FLAG_SHEET)
        if flag_name is None:
            raise RuntimeError(f"sheet '{FLAG_SHEET}' not found after processing")
        flags = wb.Worksheets(flag_name)
        bb1 = _to_int(flags.Range("BB1").Value2)
        ba1 = str(flags.Range("BA1").Value2 or "").strip().lower()
        if bb1 != count or ba1 != "a":
            raise RuntimeError(f"the BASE event macro did not run ({FLAG_SHEET}!BB1={flags.Range('BB1').Value2!r}, "
                               f"BA1={flags.Range('BA1').Value2!r})")
    else:
        if not any(_ATCH2.match(n.strip()) for n in names):
            raise RuntimeError("the BASE event macro did not run (sheet 'ATCH (2)' was not created)")
        si_name = _find(names, SI_SHEET)
        if si_name is None:
            raise RuntimeError(f"sheet '{SI_SHEET}' not found after processing")
        si = wb.Worksheets(si_name)
        c58 = si.Range("C58").Value2
        if _to_int(c58) != count:
            raise RuntimeError(f"the BASE event macro did not finish (SI!C58 holds {c58!r}, expected {count})")
        for addr, label in (("M58", "weight"), ("O58", "m3")):
            if _is_zero(si.Range(addr).Value2):
                warnings.append(f"the {label} total ({addr}) is 0, check the BASE data")
    return warnings


def _replace_with_retry(src, dst):
    """Put the finished file in place of the original. The original is never half-written."""
    last = None
    for attempt in range(SAVE_TRIES):
        try:
            os.replace(src, dst)
            return
        except PermissionError as exc:               # another program holds the file for a moment
            last = exc
        except OSError as exc:                       # e.g. the temp folder is on another drive: copy instead
            try:
                shutil.copyfile(src, dst)
                os.remove(src)
                return
            except OSError as exc2:
                last = exc2
        if attempt < SAVE_TRIES - 1:
            time.sleep(_backoff(attempt))
    raise RuntimeError(f"the finished file could not be put back after {SAVE_TRIES} tries "
                       f"(another program keeps it locked, for example antivirus): {_clean_error(last)}")


def _close_all(app):
    try:
        for i in range(app.Workbooks.Count, 0, -1):
            app.Workbooks(i).Close(SaveChanges=False)
    except Exception:
        pass


def _run_job(excel, path, base_name, count, kind):
    name = os.path.basename(path)
    app = excel.app
    wb = ws = None
    tmp_dir = None
    phase = "open the file"
    timer = excel.arm(JOB_TIMEOUT)
    try:
        wb = _retry(lambda: app.Workbooks.Open(os.path.abspath(path), UpdateLinks=0), OPEN_TRIES)
        app.Calculation = XL_CALC_AUTOMATIC
        app.EnableEvents = True
        ws = wb.Worksheets(base_name)
        ws.Visible = XL_SHEET_VISIBLE                # the event macro calls BASE.Activate

        phase = "run the BASE event macro"
        _trigger(wb, ws, count, kind)
        app.CalculateFull()                          # one full calculation, so every saved value is fresh
        warnings = _verify(wb, count, kind)
        ws.Visible = XL_SHEET_HIDDEN

        phase = "save"
        tmp_dir = Path(tempfile.mkdtemp(prefix="bl_stage3_"))
        tmp = tmp_dir / name
        _retry(lambda: wb.SaveCopyAs(str(tmp)), SAVE_TRIES)   # saved outside the watched folder, no in-place swap
        wb.Close(SaveChanges=False)
        wb = ws = None

        phase = "put the file back"
        _replace_with_retry(tmp, path)
        return {"path": path, "file": name, "status": "ok", "kind": kind, "warnings": warnings}
    except Exception as exc:
        if excel.killed:
            msg = f"{name} ({kind}): stopped after {JOB_TIMEOUT} s, Excel was probably waiting on a hidden dialog; Excel was restarted"
            return {"path": path, "file": name, "status": "error", "kind": kind, "message": msg, "timeout": True}
        return {"path": path, "file": name, "status": "error", "kind": kind,
                "message": f"{name} ({kind}): could not {phase}: {_clean_error(exc)}"}
    finally:
        timer.cancel()
        ws = None
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        wb = None
        if not excel.killed:
            _close_all(app)                          # no workbook may stay open (same-name conflicts)
        if tmp_dir is not None:
            shutil.rmtree(tmp_dir, ignore_errors=True)


# --------------------------------------------------------------------------
# WORKERS: a shared queue, so a slow file never holds the others back
# --------------------------------------------------------------------------
def _new_excel():
    return _Excel()


def _worker(jobs, results):
    _com_init()
    excel, used = None, 0
    try:
        while True:
            try:
                job = jobs.get_nowait()
            except queue.Empty:
                break
            path, base, count, kind = job
            try:
                if excel is None or excel.killed or used >= RESTART_EVERY:
                    if excel is not None:
                        excel.quit()
                        excel = None
                    excel, used = _new_excel(), 0
                results[path] = _run_job(excel, path, base, count, kind)
                used += 1
            except Exception as exc:                 # Excel itself could not start
                excel = None
                results[path] = {"path": path, "file": os.path.basename(path), "status": "error", "kind": kind,
                                 "message": f"{os.path.basename(path)} ({kind}): could not start Excel: {_clean_error(exc)}"}
    finally:
        if excel is not None:
            excel.quit()
        _com_done()


def _run_pass(jobs, results, workers):
    q = queue.Queue()
    for job in sorted(jobs, key=lambda j: -j[2]):    # the biggest files first: the batch ends sooner
        q.put(job)
    with ThreadPoolExecutor(max_workers=max(1, min(workers, len(jobs)))) as pool:
        for fut in [pool.submit(_worker, q, results) for _ in range(max(1, min(workers, len(jobs))))]:
            fut.result()


def _run_all(jobs):
    """-> (results by path, number of files that only worked on a retry)"""
    results = {}
    _run_pass(jobs, results, MAX_EXCEL_INSTANCES)

    recovered = 0
    for rnd in range(1, RETRY_ROUNDS + 1):
        failed = [j for j in jobs if results[j[0]]["status"] == "error" and not results[j[0]].get("timeout")]
        if not failed:
            break
        time.sleep(RETRY_PAUSE * rnd)                # let antivirus / indexer finish with the new files
        retry_results = {}
        _run_pass(failed, retry_results, 1)          # one Excel, no contention
        for path, res in retry_results.items():
            if res["status"] == "ok":
                recovered += 1
            results[path] = res
    return results, recovered


# --------------------------------------------------------------------------
# ENTRY POINT (called by Api.run_macro and the one-click pipeline)
# --------------------------------------------------------------------------
def _com_problem():
    try:
        import pythoncom          # noqa: F401
        import win32com.client    # noqa: F401
        import win32process       # noqa: F401
    except ImportError:
        return "pywin32 is not installed (pip install pywin32). Stage 3 needs it to drive Excel."
    return None


def process_count_rows_b1(file_paths, sheet_name="BASE"):
    try:
        paths = _as_list(file_paths)
        if not paths:
            return {"ok": False, "message": "No files provided."}
        problem = _com_problem()
        if problem:
            return {"ok": False, "message": problem}

        seen, files = set(), []
        for p in paths:
            key = os.path.normcase(os.path.abspath(p))
            if key not in seen and not os.path.basename(p).startswith("~$"):
                seen.add(key)
                files.append(p)
        if not files:
            return {"ok": False, "message": "No valid files provided."}

        _kill_leftovers()

        errors, jobs = [], []

        def check(path):
            try:
                return path, _precheck(path, sheet_name)
            except Exception as exc:
                return path, ("error", None, 0, None, str(exc))

        with ThreadPoolExecutor(max_workers=min(8, len(files))) as pool:
            for path, (status, base, count, kind, msg) in pool.map(check, files):
                if status == "ready":
                    jobs.append((path, base, count, kind))
                else:
                    errors.append(f"{os.path.basename(path)}: {msg}")

        done = {SINGLE: 0, MULTI: 0}
        warnings, recovered = [], 0
        if jobs:
            results, recovered = _run_all(jobs)
            for res in results.values():
                if res["status"] == "ok":
                    done[res["kind"]] += 1
                    warnings += [f"{res['file']}: {w}" for w in res.get("warnings", [])]
                else:
                    errors.append(res["message"])

        processed = done[SINGLE] + done[MULTI]
        message = f"Stage 3: updated B1 in {processed} file(s) ({done[SINGLE]} single, {done[MULTI]} multi)."
        if recovered:
            message += f" {recovered} file(s) needed a retry."
        if errors:
            message += f" {len(errors)} error(s)."
        return {
            "ok": processed > 0 and not errors,
            "message": message,
            "processed_count": processed,
            "single": done[SINGLE],
            "multi": done[MULTI],
            "retried": recovered,
            "warnings": warnings or None,
            "errors": errors or None,
        }
    except Exception as exc:
        return {"ok": False, "message": f"Stage 3 failed: {exc!r}"}









# import os
# import re
# import subprocess
# import time
# from concurrent.futures import ThreadPoolExecutor

# from python_calamine import CalamineWorkbook

# # --------------------------------------------------------------------------
# # CONFIG
# # --------------------------------------------------------------------------
# MAX_EXCEL_INSTANCES = 3        # separate hidden Excel processes working in parallel
# MAX_UNITS = 1000               # unit limit safeguard

# SINGLE, MULTI = "single", "multi"

# SI_SHEET = "SI"
# ATCH_SHEET = "ATCH"
# IVSI_SHEET = "IVSI"
# FLAG_SHEET = "Sheet2"

# XL_SHEET_VISIBLE = -1
# XL_SHEET_HIDDEN = 0
# XL_CALC_AUTOMATIC = -4105
# MSO_AUTOMATION_SECURITY_LOW = 1

# _ATCH2 = re.compile(r"^ATCH\s*\(2\)$", re.IGNORECASE)


# # --------------------------------------------------------------------------
# # HELPERS
# # --------------------------------------------------------------------------
# def _blank(v):
#     return v is None or (isinstance(v, str) and v.strip() == "")


# def _to_int(v):
#     try:
#         return int(float(v))
#     except (TypeError, ValueError):
#         return None


# def _find(names, wanted):
#     wanted = wanted.strip().upper()
#     return next((n for n in names if n.strip().upper() == wanted), None)


# def _kind_for(count):
#     return SINGLE if count == 1 else MULTI


# def _structure_problem(kind, names, count):
#     has_ivsi = _find(names, IVSI_SHEET) is not None
#     has_flag = _find(names, FLAG_SHEET) is not None
#     has_si = _find(names, SI_SHEET) is not None
#     has_atch = _find(names, ATCH_SHEET) is not None or any(_ATCH2.match(n.strip()) for n in names)

#     if kind == SINGLE:
#         if not has_ivsi:
#             return "1 unit, but there is no IVSI sheet"
#         if not has_flag:
#             return f"1 unit, but sheet '{FLAG_SHEET}' is missing"
#     else:
#         if has_ivsi:
#             return f"{count} units, but the file has an IVSI sheet"
#         if not (has_si and has_atch):
#             return f"{count} units, but the SI / ATCH sheets are missing"
#     return None


# # --------------------------------------------------------------------------
# # STEP 1: Calamine precheck
# # --------------------------------------------------------------------------
# def _precheck(path, sheet_name):
#     book = CalamineWorkbook.from_path(path)
#     names = book.sheet_names

#     base = _find(names, sheet_name)
#     if base is None:
#         return "error", None, 0, None, f"sheet '{sheet_name}' not found"

#     rows = book.get_sheet_by_name(base).to_python(skip_empty_area=False)
#     count = sum(1 for r in rows if len(r) > 3 and not _blank(r[3])) - 1

#     if count < 1:
#         return "error", base, count, None, "no units found in column D"
#     if count > MAX_UNITS:
#         return "error", base, count, None, f"{count} units exceeds limit of {MAX_UNITS}"

#     kind = _kind_for(count)
#     problem = _structure_problem(kind, names, count)
#     if problem:
#         return "error", base, count, kind, problem

#     return "ready", base, count, kind, None


# # --------------------------------------------------------------------------
# # STEP 2: Excel COM work (Forced Event Trigger & Strict Verification)
# # --------------------------------------------------------------------------
# def _sheet_names(wb):
#     return [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]


# def _verify_single(wb):
#     names = _sheet_names(wb)
#     s2_name = _find(names, FLAG_SHEET)
#     if s2_name is None:
#         raise RuntimeError(f"sheet '{FLAG_SHEET}' not found after processing")


# def _verify_multi(wb, count):
#     names = _sheet_names(wb)
#     si_name = _find(names, SI_SHEET)
#     if si_name is None:
#         raise RuntimeError(f"sheet '{SI_SHEET}' not found")


# def _run_job(excel, path, base_name, count, kind):
#     name = os.path.basename(path)
#     wb = None
#     try:
#         wb = excel.Workbooks.Open(os.path.abspath(path), UpdateLinks=0)
#         excel.Calculation = XL_CALC_AUTOMATIC
#         excel.EnableEvents = True

#         ws = wb.Worksheets(base_name)
#         ws.Visible = XL_SHEET_VISIBLE

#         # 1. FORCE EVENT TRIGGER: Reset B1 and Sheet2!BB1 first so Excel detects a true change
#         ws.Range("B1").Value = 0
#         try:
#             sheet2 = wb.Worksheets(FLAG_SHEET)
#             sheet2.Range("BB1").Value = -999
#             sheet2.Range("BA1").Value = ""
#         except Exception:
#             pass
        
#         excel.CalculateFull()

#         # 2. NOW SET THE ACTUAL COUNT (Guarantees Worksheet_Change macro fires)
#         ws.Range("B1").Value = count

#         # 3. Force calculation and wait/verify until the macro populates output cells
#         max_attempts = 15
#         bl_filled = False
        
#         for attempt in range(max_attempts):
#             excel.CalculateFull()
#             time.sleep(0.3)
            
#             if kind == SINGLE:
#                 ivsi_name = _find(_sheet_names(wb), IVSI_SHEET)
#                 if ivsi_name:
#                     ivsi_ws = wb.Worksheets(ivsi_name)
#                     val = ivsi_ws.Range("E4").Value2
#                     if val is not None and str(val).strip() != "":
#                         bl_filled = True
#                         break
#             else:
#                 si_name = _find(_sheet_names(wb), SI_SHEET)
#                 if si_name:
#                     si_ws = wb.Worksheets(si_name)
#                     val = si_ws.Range("C58").Value2
#                     if val is not None and str(val).strip() != "":
#                         bl_filled = True
#                         break

#         # 4. STRICT CHECKS
#         actual_b1 = _to_int(ws.Range("B1").Value2)
#         if actual_b1 != count:
#             raise RuntimeError(f"Cell B1 failed to update (holds {ws.Range('B1').Value2!r}, expected {count})")
        
#         if not bl_filled:
#             raise RuntimeError("Event macro timed out: Output cells (E4/C58) remained empty after calculation.")

#         if kind == SINGLE:
#             _verify_single(wb)
#         else:
#             _verify_multi(wb, count)

#         ws.Visible = XL_SHEET_HIDDEN
        
#         # 5. Explicit save and close to prevent modal popup dialogs
#         wb.Save()
#         wb.Close(SaveChanges=False)
#         wb = None
        
#         return {"file": name, "status": "ok", "kind": kind}
#     except Exception as exc:
#         if wb is not None:
#             try:
#                 wb.Close(SaveChanges=False)
#             except Exception:
#                 pass
#             wb = None
#         return {"file": name, "status": "error", "message": f"{name} ({kind}): {exc}"}


# def _excel_worker(jobs):
#     import pythoncom
#     import win32com.client

#     results = []
#     pythoncom.CoInitialize()
#     excel = None
#     try:
#         excel = win32com.client.DispatchEx("Excel.Application")
#         excel.Visible = False
#         excel.DisplayAlerts = False
#         excel.ScreenUpdating = False
#         excel.AskToUpdateLinks = False
#         excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW

#         for path, base_name, count, kind in jobs:
#             results.append(_run_job(excel, path, base_name, count, kind))
#     except Exception as exc:
#         done = {r["file"] for r in results}
#         for path, _, _, _ in jobs:
#             n = os.path.basename(path)
#             if n not in done:
#                 results.append({"file": n, "status": "error", "message": f"{n}: Excel failed: {exc}"})
#     finally:
#         if excel is not None:
#             try:
#                 excel.Quit()
#             except Exception:
#                 pass
#         pythoncom.CoUninitialize()
#     return results


# # --------------------------------------------------------------------------
# # ENTRY POINT
# # --------------------------------------------------------------------------
# def process_count_rows_b1(file_paths: list[str], sheet_name: str = "BASE") -> dict:
#     if not file_paths:
#         return {"ok": False, "message": "No files provided."}

#     # Safety sweep: Kill any stuck background Excel COM processes before starting
#     try:
#         subprocess.run(["taskkill", "/f", "/im", "EXCEL.EXE"], capture_output=True, text=True)
#     except Exception:
#         pass

#     try:
#         import pythoncom
#         import win32com.client
#     except ImportError:
#         return {"ok": False, "message": "pywin32 is not installed."}

#     files = [p for p in dict.fromkeys(file_paths) if not os.path.basename(p).startswith("~$")]
#     if not files:
#         return {"ok": False, "message": "No valid files provided."}

#     errors, jobs = [], []

#     def check(path):
#         try:
#             return path, _precheck(path, sheet_name)
#         except Exception as exc:
#             return path, ("error", None, 0, None, str(exc))

#     with ThreadPoolExecutor(max_workers=min(8, len(files))) as pool:
#         for path, (status, base, count, kind, msg) in pool.map(check, files):
#             n = os.path.basename(path)
#             if status == "ready":
#                 jobs.append((path, base, count, kind))
#             else:
#                 errors.append(f"{n}: {msg}")

#     done = {SINGLE: 0, MULTI: 0}
#     if jobs:
#         workers = max(1, min(MAX_EXCEL_INSTANCES, len(jobs)))
#         chunks = [jobs[i::workers] for i in range(workers)]

#         with ThreadPoolExecutor(max_workers=workers) as pool:
#             for results in pool.map(_excel_worker, chunks):
#                 for r in results:
#                     if r["status"] == "ok":
#                         done[r["kind"]] += 1
#                     else:
#                         errors.append(r["message"])

#     processed = done[SINGLE] + done[MULTI]
#     message = f"Stage 3: updated B1 in {processed} file(s) ({done[SINGLE]} single, {done[MULTI]} multi)."
#     if errors:
#         message += f" {len(errors)} error(s)."

#     return {
#         "ok": processed > 0 and not errors,
#         "message": message,
#         "processed_count": processed,
#         "single": done[SINGLE],
#         "multi": done[MULTI],
#         "errors": errors or None,
#     }






