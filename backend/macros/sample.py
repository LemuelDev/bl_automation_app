import os
import re
import threading
import time
from datetime import datetime
from pathlib import Path

from backend.macros import shipping_list as s1
from backend.macros import copy_IV_base_create_SI as s2
from backend.macros import input_data_b1 as s3
from backend.macros import bl_creation_single_and_multi_files as s4
from backend.macros import checker_result as s5
from backend.macros import inf_macro as s6
from backend.macros import move_or_copy_attach as s7
from backend.macros import all_bl as s8

# --------------------------------------------------------------------------
# CONFIG (every default path in one place, built from the user's Downloads folder)
# --------------------------------------------------------------------------
BL_ROOT = Path(os.path.expanduser("~")) / "Downloads" / "BLC ALL MACRO"
WIP = BL_ROOT / "WIP BL"

SEPARATED_IV = WIP / "SEPARATED IV"
FOR_RPA = WIP / "FOR RPA"
FOR_RPA_SINGLE = FOR_RPA / "SINGLE"
FOR_RPA_MULTI = FOR_RPA / "MULTI"
BL_SINGLE = WIP / "BL" / "SINGLE"
BL_MULTI = WIP / "BL" / "MULTI"
BL_ALL = WIP / "BL" / "ALL"

# Where the finished files (shipping list, checker result, BL output) are saved in a one-click run.
OUTPUT_DIR = WIP / "BL OUTPUT"

STAGES = [
    (1, "Shipping List"),
    (2, "Copy IV Base Create SI"),
    (3, "Input Data (Base B1)"),
    (4, "BL Creation"),
    (5, "Checker Result"),
    (6, "Removing Zero"),
    (7, "Move or Copy Attach"),
    (8, "All BL"),
]

_EXCEL_FILE = re.compile(r"\.(xlsx|xlsm|xlsb|xls)$", re.IGNORECASE)


# --------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------
def _as_list(value):
    if not value:
        return []
    if isinstance(value, (str, os.PathLike)):
        return [os.fspath(value)]
    return [os.fspath(v) for v in value]


def _files_in(folder, pattern=_EXCEL_FILE):
    """The Excel files directly inside a folder (no ~$ lock files)."""
    folder = Path(folder)
    if not folder.is_dir():
        return []
    return sorted(p for p in folder.iterdir()
                if p.is_file() and not p.name.startswith("~$") and pattern.search(p.name))


def _xlsm_in(folder):
    return _files_in(folder, re.compile(r"\.xlsm$", re.IGNORECASE))


def _lock_files(folder):
    folder = Path(folder)
    return [p for p in folder.iterdir() if p.is_file() and p.name.startswith("~$")] if folder.is_dir() else []


def _fmt_time(seconds):
    return f"{seconds:.1f} s" if seconds < 60 else f"{int(seconds // 60)} min {int(seconds % 60)} s"


# --------------------------------------------------------------------------
# RUNNER
# --------------------------------------------------------------------------
class PipelineRunner:
    """Runs stages 1 to 8 in order in a background thread, so the window stays alive and can show progress."""

    def __init__(self):
        self._lock = threading.Lock()
        self._state = self._fresh_state()
        self.current_vessel = None  # Holds vessel details from frontend

    # ---- state shown in the UI ---------------------------------------------
    @staticmethod
    def _fresh_state():
        return {
            "running": False, "done": False, "ok": None, "message": "Idle",
            "current_stage": 0, "total_stages": len(STAGES),
            "stages": [{"id": i, "name": n, "status": "pending", "message": ""} for i, n in STAGES],
            "log": [], "errors": [], "warnings": [], "archive_dir": None, "outputs": {},
        }

    def status(self):
        with self._lock:
            s = dict(self._state)
            s["stages"] = [dict(x) for x in self._state["stages"]]
            for key in ("log", "errors", "warnings"):
                s[key] = list(self._state[key])
            s["outputs"] = dict(self._state["outputs"])
            return s

    def _log(self, msg):
        with self._lock:
            self._state["log"].append(f"{datetime.now():%H:%M:%S}  {msg}")
            self._state["log"] = self._state["log"][-300:]

    def _set_stage(self, n, status, message=""):
        with self._lock:
            if status == "running":
                self._state["current_stage"] = n
                self._state["message"] = f"Stage {n} of {len(STAGES)}: {STAGES[n - 1][1]}"
            st = self._state["stages"][n - 1]
            st["status"], st["message"] = status, message

    def _issues(self, n, result):
        with self._lock:
            for e in _as_list(result.get("errors")):
                self._state["errors"].append(f"[Stage {n}] {e}")
            for w in _as_list(result.get("warnings")):
                self._state["warnings"].append(f"[Stage {n}] {w}")

    def _finish(self, ok, message):
        with self._lock:
            self._state.update(running=False, done=True, ok=ok, message=message)

    def _stop(self, n, message):
        """The stage failed: skip every stage after it and end the run."""
        with self._lock:
            self._state["stages"][n - 1].update(status="failed", message=message)
            for s in self._state["stages"]:
                if s["status"] == "pending":
                    s["status"] = "skipped"
        self._log(f"STOPPED at stage {n}: {message}")
        self._finish(False, f"Stopped at stage {n} ({STAGES[n - 1][1]}): {message}")

    # ---- checks before anything is touched -------------------------------------
    @staticmethod
    def _excel_problem():
        try:
            import pythoncom         # noqa: F401
            import win32com.client    # noqa: F401
            import winreg
            winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
        except ImportError:
            return "pywin32 is not installed (pip install pywin32). Stage 3 needs it."
        except OSError:
            return "Microsoft Excel does not seem to be installed. Stage 3 needs it."
        return None

    def preview(self):
        """How many old files this run will delete (shown in the confirmation pop-up)."""
        return {
            "stage1_clears": 0,
            "stage2_clears": len(_files_in(FOR_RPA_SINGLE)) + len(_files_in(FOR_RPA_MULTI)),
            "stage4_clears": sum(len(_files_in(f)) for f in (BL_SINGLE, BL_MULTI, BL_ALL)),
            "mode": "delete",
        }

    def preflight(self, selection, vessel=None):
        """Everything that could make a later stage fail, checked BEFORE the first file is deleted."""
        if not BL_ROOT.is_dir():
            return [f"The folder 'BLC ALL MACRO' was not found in Downloads: {BL_ROOT}"]

        problems = []
        
        # Validate vessel inputs if provided
        if vessel:
            _, v_problem = s4.validate_details(vessel)
            if v_problem:
                problems.append(v_problem)
        else:
            problems.append("Vessel details are missing. Please fill out the vessel configuration.")

        for item in selection:
            if not os.path.exists(item):
                problems.append(f"Selected file or folder not found: {item}")
        if selection == [str(SEPARATED_IV)] and not _files_in(SEPARATED_IV):
            problems.append(f"No IV files were selected and the SEPARATED IV folder is empty: {SEPARATED_IV}")

        for folder in (FOR_RPA_SINGLE, FOR_RPA_MULTI, BL_SINGLE, BL_MULTI, BL_ALL):
            for lock in _lock_files(folder):
                problems.append(f"A file is open in Excel ({folder.name}): {lock.name[2:]}. Close it first.")

        # Check Stage 1 template
        if s1.find_template()[0] is None:
            problems.append(f"Template not found for Stage 1: {s1.TEMPLATE_NAME} (project templates folder)")

        # Check Stage 2 templates
        for name in (s2.SINGLE_TEMPLATE_NAME, s2.MULTI_TEMPLATE_NAME):
            if not (Path(s2.TEMPLATE_DIR) / name).is_file():
                problems.append(f"Template not found for Stage 2: {Path(s2.TEMPLATE_DIR) / name}")
        
        # Check Stage 4 internal project templates (Standard & Navis)
        for (kind, navis), filename in s4.TEMPLATE_FILES.items():
            path, tried = s4.find_template(kind, navis)
            if path is None:
                tried_str = ", ".join(str(p.parent) for p in tried)
                problems.append(f"Template not found for Stage 4: {filename} (looked in: {tried_str})")

        # Check Stage 5 template
        if s5.find_template()[0] is None:
            problems.append(f"Template not found for Stage 5: {s5.TEMPLATE_NAME} (project templates folder)")
            
        # Check Stage 8 template
        if s8.find_template()[0] is None:
            problems.append(f"Template not found for Stage 8: {s8.MODULE_TEMPLATE} (project templates folder)")

        excel = self._excel_problem()
        if excel:
            problems.append(excel)
        return problems

    # ---- start -----------------------------------------------------------------
    def start(self, files, sheet_name="BASE", vessel=None, confirm=None, on_finish=None, output_dir=None):
        """
        files: the IV files (or folder) the user selected. Without a selection the SEPARATED IV folder is used.
        vessel: dictionary of vessel metadata (name, voyage, pol, pod, etd, navis).
        confirm: function(title, text) -> bool, used by Stage 1 when it finds duplicates.
        on_finish: called when the run ends (releases the app-wide lock).
        output_dir: where the shipping list, checker result and BL output go (default: OUTPUT_DIR).
        """
        selection = _as_list(files) or [str(SEPARATED_IV)]
        problems = self.preflight(selection, vessel=vessel)
        if problems:
            return {"ok": False, "message": f"Cannot start: {problems[0]}", "problems": problems}

        with self._lock:
            if self._state["running"]:
                return {"ok": False, "message": "A run is already in progress."}
            self.current_vessel = vessel
            self._state = self._fresh_state()
            self._state["running"] = True
            self._state["message"] = "Starting..."

        out_dir = Path(output_dir) if output_dir else OUTPUT_DIR
        threading.Thread(target=self._run, args=(selection, sheet_name, out_dir, confirm, on_finish),
                         daemon=True).start()
        return {"ok": True, "message": "Create BL started.", "output_dir": str(out_dir)}

    # ---- the run -----------------------------------------------------------------
    def _stage(self, n, func, **kwargs):
        name = STAGES[n - 1][1]
        self._set_stage(n, "running")
        self._log(f"Stage {n} ({name}) started")
        t0 = time.perf_counter()
        try:
            result = func(**kwargs) or {}
        except Exception as exc:
            result = {"ok": False, "message": f"{name} crashed: {exc!r}"}
        took = time.perf_counter() - t0
        ok = bool(result.get("ok"))
        message = f"{result.get('message', '')} ({_fmt_time(took)})".strip()
        self._issues(n, result)
        self._set_stage(n, "done" if ok else "failed", message)
        self._log(f"Stage {n}: {message}")
        return ok, result

    def _run(self, selection, sheet_name, out_dir, confirm, on_finish):
        try:
            self._execute(selection, sheet_name, out_dir, confirm)
        except Exception as exc:
            self._log(f"Unexpected error: {exc!r}")
            self._stop(max(1, self.status()["current_stage"]), f"unexpected error: {exc!r}")
        finally:
            if on_finish:
                try:
                    on_finish()
                except Exception:
                    pass

    def _execute(self, selection, sheet_name, out_dir, confirm):
        started = time.perf_counter()
        stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        out_dir.mkdir(parents=True, exist_ok=True)
        outputs = {}

        # 1 - shipping list (the file is saved straight into the output folder, no Save As window)
        path = out_dir / f"{Path(s1.TEMPLATE_NAME).stem}_{stamp}.xlsm"
        ok, r = self._stage(1, s1.process_shipping_list, file_paths=selection, sheet_name=sheet_name,
                            output_path=str(path), confirm=confirm)
        if not ok:
            return self._stop(1, r.get("message", "failed"))
        outputs["Shipping list"] = r.get("output_path", str(path))

        # 2 - SI files for exactly the IV files that were selected
        ok, r = self._stage(2, s2.process_si_creation, folder_or_files=selection, sheet_name=sheet_name)
        if not ok:
            return self._stop(2, r.get("message", "failed"))

        # 3 - B1 through Excel, for every file Stage 2 just created
        rpa_files = [str(p) for p in _xlsm_in(FOR_RPA_SINGLE) + _xlsm_in(FOR_RPA_MULTI)]
        if not rpa_files:
            return self._stop(3, "Stage 2 finished but the FOR RPA folders are empty")
        ok, r = self._stage(3, s3.process_count_rows_b1, file_paths=rpa_files, sheet_name=sheet_name)
        if not ok:
            return self._stop(3, r.get("message", "failed"))

        # 4 - BL files (BL\SINGLE, BL\MULTI and a copy of everything in BL\ALL) with vessel metadata injection
        ok, r4 = self._stage(4, s4.process_file, folder_or_files=str(FOR_RPA), vessel=self.current_vessel)
        if not ok:
            return self._stop(4, r4.get("message", "failed"))

        # 5 - checker result of BL\ALL
        stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        path = out_dir / f"{Path(s5.TEMPLATE_CHECKER).stem}_{stamp}.xlsx"
        
        ok, r = self._stage(5, s5.process_checker_result, folder_or_files=str(BL_ALL), output_path=str(path))
        if not ok:
            return self._stop(5, r.get("message", "failed"))
            
        outputs["Checker result"] = r.get("output_path", str(path))
        outputs["IV番号管理表"] = r.get("iv_kanri_path", "")

        # 6 - blank A18 / A19 in BL\ALL
        ok, r = self._stage(6, s6.process_remove_zero, folder_or_files=str(BL_ALL))
        if not ok:
            return self._stop(6, r.get("message", "failed"))

        # 7 - attach sheets (only multi-unit files have one)
        if _xlsm_in(FOR_RPA_MULTI):
            ok, r = self._stage(7, s7.process_attach, folder_or_files=str(FOR_RPA_MULTI))
            if not ok:
                return self._stop(7, r.get("message", "failed"))
        else:
            self._set_stage(7, "skipped", "no multi-unit files in this batch")
            self._log("Stage 7 skipped: no multi-unit files in this batch")

        # 8 - everything in one workbook
        path = out_dir / f"{s8.OUTPUT_PREFIX}_{stamp}.xlsm"
        ok, r = self._stage(8, s8.process_all_bl, folder_or_files=str(BL_ALL), output_path=str(path))
        if not ok:
            return self._stop(8, r.get("message", "failed"))
        outputs["BL output"] = r.get("output_path", str(path))

        with self._lock:
            self._state["outputs"] = outputs
            n_err, n_warn = len(self._state["errors"]), len(self._state["warnings"])
        message = (f"Create BL finished in {_fmt_time(time.perf_counter() - started)}: "
                f"{r4.get('single', 0)} single and {r4.get('multi', 0)} multi BL file(s). "
                f"Shipping list, checker result and BL output are in {out_dir}.")
        if n_err or n_warn:
            message += f" {n_err} error(s) and {n_warn} warning(s), see the details."
        self._log("Create BL finished")
        self._finish(True, message)





# import os
# import re
# import threading
# import time
# from datetime import datetime
# from pathlib import Path

# from backend.macros import shipping_list as s1
# from backend.macros import copy_IV_base_create_SI as s2
# from backend.macros import input_data_b1 as s3
# from backend.macros import bl_creation_single_and_multi_files as s4
# from backend.macros import checker_result as s5
# from backend.macros import inf_macro as s6
# from backend.macros import move_or_copy_attach as s7
# from backend.macros import all_bl as s8

# # --------------------------------------------------------------------------
# # CONFIG (every default path in one place, built from the user's Downloads folder)
# # --------------------------------------------------------------------------
# BL_ROOT = Path(os.path.expanduser("~")) / "Downloads" / "BLC ALL MACRO"
# WIP = BL_ROOT / "WIP BL"

# SEPARATED_IV = WIP / "SEPARATED IV"
# FOR_RPA = WIP / "FOR RPA"
# FOR_RPA_SINGLE = FOR_RPA / "SINGLE"
# FOR_RPA_MULTI = FOR_RPA / "MULTI"
# BL_SINGLE = WIP / "BL" / "SINGLE"
# BL_MULTI = WIP / "BL" / "MULTI"
# BL_ALL = WIP / "BL" / "ALL"

# # Where the three finished files (shipping list, checker result, BL output) are saved in a one-click run.
# OUTPUT_DIR = WIP / "BL OUTPUT"

# STAGES = [
#     (1, "Shipping List"),
#     (2, "Copy IV Base Create SI"),
#     (3, "Input Data (Base B1)"),
#     (4, "BL Creation"),
#     (5, "Checker Result"),
#     (6, "Removing Zero"),
#     (7, "Move or Copy Attach"),
#     (8, "All BL"),
# ]

# _EXCEL_FILE = re.compile(r"\.(xlsx|xlsm|xlsb|xls)$", re.IGNORECASE)


# # --------------------------------------------------------------------------
# # HELPERS
# # --------------------------------------------------------------------------
# def _as_list(value):
#     if not value:
#         return []
#     if isinstance(value, (str, os.PathLike)):
#         return [os.fspath(value)]
#     return [os.fspath(v) for v in value]


# def _files_in(folder, pattern=_EXCEL_FILE):
#     """The Excel files directly inside a folder (no ~$ lock files)."""
#     folder = Path(folder)
#     if not folder.is_dir():
#         return []
#     return sorted(p for p in folder.iterdir()
#                 if p.is_file() and not p.name.startswith("~$") and pattern.search(p.name))


# def _xlsm_in(folder):
#     return _files_in(folder, re.compile(r"\.xlsm$", re.IGNORECASE))


# def _lock_files(folder):
#     folder = Path(folder)
#     return [p for p in folder.iterdir() if p.is_file() and p.name.startswith("~$")] if folder.is_dir() else []


# def _fmt_time(seconds):
#     return f"{seconds:.1f} s" if seconds < 60 else f"{int(seconds // 60)} min {int(seconds % 60)} s"


# # --------------------------------------------------------------------------
# # RUNNER
# # --------------------------------------------------------------------------
# class PipelineRunner:
#     """Runs stages 1 to 8 in order in a background thread, so the window stays alive and can show progress."""

#     def __init__(self):
#         self._lock = threading.Lock()
#         self._state = self._fresh_state()

#     # ---- state shown in the UI ---------------------------------------------
#     @staticmethod
#     def _fresh_state():
#         return {
#             "running": False, "done": False, "ok": None, "message": "Idle",
#             "current_stage": 0, "total_stages": len(STAGES),
#             "stages": [{"id": i, "name": n, "status": "pending", "message": ""} for i, n in STAGES],
#             "log": [], "errors": [], "warnings": [], "archive_dir": None, "outputs": {},
#         }

#     def status(self):
#         with self._lock:
#             s = dict(self._state)
#             s["stages"] = [dict(x) for x in self._state["stages"]]
#             for key in ("log", "errors", "warnings"):
#                 s[key] = list(self._state[key])
#             s["outputs"] = dict(self._state["outputs"])
#             return s

#     def _log(self, msg):
#         with self._lock:
#             self._state["log"].append(f"{datetime.now():%H:%M:%S}  {msg}")
#             self._state["log"] = self._state["log"][-300:]

#     def _set_stage(self, n, status, message=""):
#         with self._lock:
#             if status == "running":
#                 self._state["current_stage"] = n
#                 self._state["message"] = f"Stage {n} of {len(STAGES)}: {STAGES[n - 1][1]}"
#             st = self._state["stages"][n - 1]
#             st["status"], st["message"] = status, message

#     def _issues(self, n, result):
#         with self._lock:
#             for e in _as_list(result.get("errors")):
#                 self._state["errors"].append(f"[Stage {n}] {e}")
#             for w in _as_list(result.get("warnings")):
#                 self._state["warnings"].append(f"[Stage {n}] {w}")

#     def _finish(self, ok, message):
#         with self._lock:
#             self._state.update(running=False, done=True, ok=ok, message=message)

#     def _stop(self, n, message):
#         """The stage failed: skip every stage after it and end the run."""
#         with self._lock:
#             self._state["stages"][n - 1].update(status="failed", message=message)
#             for s in self._state["stages"]:
#                 if s["status"] == "pending":
#                     s["status"] = "skipped"
#         self._log(f"STOPPED at stage {n}: {message}")
#         self._finish(False, f"Stopped at stage {n} ({STAGES[n - 1][1]}): {message}")

#     # ---- checks before anything is touched -------------------------------------
#     @staticmethod
#     def _excel_problem():
#         try:
#             import pythoncom          # noqa: F401
#             import win32com.client    # noqa: F401
#             import winreg
#             winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
#         except ImportError:
#             return "pywin32 is not installed (pip install pywin32). Stage 3 needs it."
#         except OSError:
#             return "Microsoft Excel does not seem to be installed. Stage 3 needs it."
#         return None

#     def preview(self):
#         """How many old files this run will delete (shown in the confirmation pop-up)."""
#         return {
#             "stage1_clears": 0,
#             "stage2_clears": len(_files_in(FOR_RPA_SINGLE)) + len(_files_in(FOR_RPA_MULTI)),
#             "stage4_clears": sum(len(_files_in(f)) for f in (BL_SINGLE, BL_MULTI, BL_ALL)),
#             "mode": "delete",
#         }

#     def preflight(self, selection):
#         """Everything that could make a later stage fail, checked BEFORE the first file is deleted."""
#         if not BL_ROOT.is_dir():
#             return [f"The folder 'BLC ALL MACRO' was not found in Downloads: {BL_ROOT}"]

#         problems = []
#         for item in selection:
#             if not os.path.exists(item):
#                 problems.append(f"Selected file or folder not found: {item}")
#         if selection == [str(SEPARATED_IV)] and not _files_in(SEPARATED_IV):
#             problems.append(f"No IV files were selected and the SEPARATED IV folder is empty: {SEPARATED_IV}")

#         for folder in (FOR_RPA_SINGLE, FOR_RPA_MULTI, BL_SINGLE, BL_MULTI, BL_ALL):
#             for lock in _lock_files(folder):
#                 problems.append(f"A file is open in Excel ({folder.name}): {lock.name[2:]}. Close it first.")

#         if s1.find_template()[0] is None:
#             problems.append(f"Template not found for Stage 1: {s1.TEMPLATE_NAME} (project templates folder)")
#         for name in (s2.SINGLE_TEMPLATE_NAME, s2.MULTI_TEMPLATE_NAME):
#             if not (Path(s2.TEMPLATE_DIR) / name).is_file():
#                 problems.append(f"Template not found for Stage 2: {Path(s2.TEMPLATE_DIR) / name}")
#         for tpl in s4.TEMPLATES.values():
#             if not Path(tpl).is_file():
#                 problems.append(f"Template not found for Stage 4: {tpl}")
#         if s5.find_template()[0] is None:
#             problems.append(f"Template not found for Stage 5: {s5.TEMPLATE_NAME} (project templates folder)")
#         if s8.find_template()[0] is None:
#             problems.append(f"Template not found for Stage 8: {s8.MODULE_TEMPLATE} (project templates folder)")

#         excel = self._excel_problem()
#         if excel:
#             problems.append(excel)
#         return problems

#     # ---- start -----------------------------------------------------------------
#     def start(self, files, sheet_name="BASE", confirm=None, on_finish=None, output_dir=None):
#         """
#         files: the IV files (or folder) the user selected. Without a selection the SEPARATED IV folder is used.
#         confirm: function(title, text) -> bool, used by Stage 1 when it finds duplicates.
#         on_finish: called when the run ends (releases the app-wide lock).
#         output_dir: where the shipping list, checker result and BL output go (default: OUTPUT_DIR).
#         """
#         selection = _as_list(files) or [str(SEPARATED_IV)]
#         problems = self.preflight(selection)
#         if problems:
#             return {"ok": False, "message": f"Cannot start: {problems[0]}", "problems": problems}

#         with self._lock:
#             if self._state["running"]:
#                 return {"ok": False, "message": "A run is already in progress."}
#             self._state = self._fresh_state()
#             self._state["running"] = True
#             self._state["message"] = "Starting..."

#         out_dir = Path(output_dir) if output_dir else OUTPUT_DIR
#         threading.Thread(target=self._run, args=(selection, sheet_name, out_dir, confirm, on_finish),
#                          daemon=True).start()
#         return {"ok": True, "message": "Create BL started.", "output_dir": str(out_dir)}

#     # ---- the run -----------------------------------------------------------------
#     def _stage(self, n, func, **kwargs):
#         name = STAGES[n - 1][1]
#         self._set_stage(n, "running")
#         self._log(f"Stage {n} ({name}) started")
#         t0 = time.perf_counter()
#         try:
#             result = func(**kwargs) or {}
#         except Exception as exc:
#             result = {"ok": False, "message": f"{name} crashed: {exc!r}"}
#         took = time.perf_counter() - t0
#         ok = bool(result.get("ok"))
#         message = f"{result.get('message', '')} ({_fmt_time(took)})".strip()
#         self._issues(n, result)
#         self._set_stage(n, "done" if ok else "failed", message)
#         self._log(f"Stage {n}: {message}")
#         return ok, result

#     def _run(self, selection, sheet_name, out_dir, confirm, on_finish):
#         try:
#             self._execute(selection, sheet_name, out_dir, confirm)
#         except Exception as exc:
#             self._log(f"Unexpected error: {exc!r}")
#             self._stop(max(1, self.status()["current_stage"]), f"unexpected error: {exc!r}")
#         finally:
#             if on_finish:
#                 try:
#                     on_finish()
#                 except Exception:
#                     pass

#     def _execute(self, selection, sheet_name, out_dir, confirm):
#         started = time.perf_counter()
#         stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
#         out_dir.mkdir(parents=True, exist_ok=True)
#         outputs = {}

#         # 1 - shipping list (the file is saved straight into the output folder, no Save As window)
#         path = out_dir / f"{Path(s1.TEMPLATE_NAME).stem}_{stamp}.xlsm"
#         ok, r = self._stage(1, s1.process_shipping_list, file_paths=selection, sheet_name=sheet_name,
#                             output_path=str(path), confirm=confirm)
#         if not ok:
#             return self._stop(1, r.get("message", "failed"))
#         outputs["Shipping list"] = r.get("output_path", str(path))

#         # 2 - SI files for exactly the IV files that were selected
#         ok, r = self._stage(2, s2.process_si_creation, folder_or_files=selection, sheet_name=sheet_name)
#         if not ok:
#             return self._stop(2, r.get("message", "failed"))

#         # 3 - B1 through Excel, for every file Stage 2 just created
#         rpa_files = [str(p) for p in _xlsm_in(FOR_RPA_SINGLE) + _xlsm_in(FOR_RPA_MULTI)]
#         if not rpa_files:
#             return self._stop(3, "Stage 2 finished but the FOR RPA folders are empty")
#         ok, r = self._stage(3, s3.process_count_rows_b1, file_paths=rpa_files, sheet_name=sheet_name)
#         if not ok:
#             return self._stop(3, r.get("message", "failed"))

#         # 4 - BL files (BL\SINGLE, BL\MULTI and a copy of everything in BL\ALL)
#         ok, r4 = self._stage(4, s4.process_file, folder_or_files=str(FOR_RPA))
#         if not ok:
#             return self._stop(4, r4.get("message", "failed"))

#         # 5 - checker result of BL\ALL
        
#         # Generate output path using the checker template name
#         stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
#         path = out_dir / f"{Path(s5.TEMPLATE_CHECKER).stem}_{stamp}.xlsx"
        
#         ok, r = self._stage(5, s5.process_checker_result, folder_or_files=str(BL_ALL), output_path=str(path))
#         if not ok:
#             return self._stop(5, r.get("message", "failed"))
            
#         outputs["Checker result"] = r.get("output_path", str(path))
#         outputs["IV番号管理表"] = r.get("iv_kanri_path", "")

#         # 6 - blank A18 / A19 in BL\ALL
#         ok, r = self._stage(6, s6.process_remove_zero, folder_or_files=str(BL_ALL))
#         if not ok:
#             return self._stop(6, r.get("message", "failed"))

#         # 7 - attach sheets (only multi-unit files have one)
#         if _xlsm_in(FOR_RPA_MULTI):
#             ok, r = self._stage(7, s7.process_attach, folder_or_files=str(FOR_RPA_MULTI))
#             if not ok:
#                 return self._stop(7, r.get("message", "failed"))
#         else:
#             self._set_stage(7, "skipped", "no multi-unit files in this batch")
#             self._log("Stage 7 skipped: no multi-unit files in this batch")

#         # 8 - everything in one workbook
#         path = out_dir / f"{s8.OUTPUT_PREFIX}_{stamp}.xlsm"
#         ok, r = self._stage(8, s8.process_all_bl, folder_or_files=str(BL_ALL), output_path=str(path))
#         if not ok:
#             return self._stop(8, r.get("message", "failed"))
#         outputs["BL output"] = r.get("output_path", str(path))

#         with self._lock:
#             self._state["outputs"] = outputs
#             n_err, n_warn = len(self._state["errors"]), len(self._state["warnings"])
#         message = (f"Create BL finished in {_fmt_time(time.perf_counter() - started)}: "
#                 f"{r4.get('single', 0)} single and {r4.get('multi', 0)} multi BL file(s). "
#                 f"Shipping list, checker result and BL output are in {out_dir}.")
#         if n_err or n_warn:
#             message += f" {n_err} error(s) and {n_warn} warning(s), see the details."
#         self._log("Create BL finished")
#         self._finish(True, message)