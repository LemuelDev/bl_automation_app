import os
import threading
import webview

from backend.macros.shipping_list import process_shipping_list
from backend.macros.copy_IV_base_create_SI import process_si_creation
from backend.macros.input_data_b1 import process_count_rows_b1
from backend.macros.bl_creation_single_and_multi_files import process_file
from backend.macros.checker_result import process_checker_result
from backend.macros.inf_macro import process_remove_zero
from backend.macros.move_or_copy_attach import process_attach
from backend.macros.all_bl import process_all_bl
from backend.macros.sample import PipelineRunner

def _expand_to_xlsm(folder_or_files):
    """Folder, file path, or list of either -> list of .xlsm file paths (no ~$ temp files)."""
    if isinstance(folder_or_files, (str, os.PathLike)):
        folder_or_files = [folder_or_files]

    files = []
    for item in folder_or_files:
        item = os.fspath(item)
        if os.path.isdir(item):
            for name in sorted(os.listdir(item)):
                if name.lower().endswith(".xlsm") and not name.startswith("~$"):
                    files.append(os.path.join(item, name))
        elif os.path.isfile(item):
            files.append(item)
    return list(dict.fromkeys(files))   # de-duplicate, keep order


class Api:
    def __init__(self):
        self._window = None
        self._macro_lock = threading.Lock()   # only one macro at a time (Stage 3 drives Excel)
        self._pipeline = PipelineRunner()

    def set_window(self, window):
        self._window = window

    # ---------- File / Folder Dialogs ----------

    def select_file(self, file_types=None):
        file_types = file_types or ("Excel files (*.xlsx;*.xlsm;*.xls)",)
        result = self._window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types
        )
        return result[0] if result else None

    def select_files(self, file_types=None):
        file_types = file_types or ("Excel files (*.xlsx;*.xlsm;*.xls)",)
        result = self._window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=True, file_types=file_types
        )
        return list(result) if result else []

    def select_folder(self):
        result = self._window.create_file_dialog(webview.FOLDER_DIALOG)
        return result[0] if result else None
    
    def _ask_save_path(self, suggested_name, start_dir=None, ext="xlsm"):
        labels = {
            "xlsm": "Excel Macro Enabled Workbook (*.xlsm)",
            "xlsx": "Excel Workbook (*.xlsx)",
        }
        kwargs = {
            "save_filename": suggested_name,
            "file_types": (labels[ext],),
        }
        if start_dir and os.path.isdir(start_dir):
            kwargs["directory"] = start_dir
        result = self._window.create_file_dialog(webview.SAVE_DIALOG, **kwargs)
        if isinstance(result, (list, tuple)):          # some pywebview versions return a tuple
            result = result[0] if result else None
        return result or None

    def _confirm(self, title, message):
        """Yes / No pop-up. If the window cannot show one, carry on (the warnings are still in the log)."""
        try:
            return bool(self._window.create_confirmation_dialog(title, message))
        except Exception:
            return True
        
    # helpers for the one-click BL automation
    def pipeline_preview(self):
        """Called by app.js to show confirmation counts of old files."""
        return self._pipeline.preview()

    def start_pipeline(self, payload):
        """Triggers the one-click automated multi-stage background run."""
        files = payload.get("folder_or_files") or payload.get("file_paths", [])
        sheet_name = payload.get("sheet_name", "BASE")
        vessel = payload.get("vessel")

        if not self._macro_lock.acquire(blocking=False):
            return {"ok": False, "message": "Another macro or run is still in progress."}

        result = self._pipeline.start(
            files, 
            sheet_name,
            vessel=vessel,
            confirm=self._confirm,
            on_finish=self._macro_lock.release,
            output_dir=payload.get("output_dir"),
        )
        if not result.get("ok"):
            self._macro_lock.release()  # Free the lock if it failed to start
        return result

    def pipeline_status(self):
        """Polled by app.js to update frontend UI logs and step indicators."""
        return self._pipeline.status()
        
    
    # ---------- Macro Router ----------

    def run_macro(self, macro_id, payload):
        folder_or_files = payload.get("folder_or_files") or payload.get("file_paths", [])
        sheet_name = payload.get("sheet_name", "BASE")
        vessel = payload.get("vessel")

        if not folder_or_files:
            return {"ok": False, "message": "No files or folder selected to process."}

        if not self._macro_lock.acquire(blocking=False):
            return {"ok": False, "message": "Another macro is still running. Please wait for it to finish."}

        try:
            # Stage 1
            if macro_id == "shipping_list":
                template_folder = "../templates/SO.UT.ATTACH.xlsm"
                return process_shipping_list(
                    file_paths=folder_or_files,
                    sheet_name=sheet_name,
                    ask_save_path=self._ask_save_path,
                    confirm=self._confirm,
                )

            # Stage 2: create SI files from the IV files
            if macro_id in ["copy_iv_base_si", "copy_iv_base_create_si", "stage_2"]:
                return process_si_creation(
                    folder_or_files=folder_or_files,
                    sheet_name=sheet_name,
                )

            # Stage 3: input the unit count in BASE!B1 (runs through Excel)
            if macro_id in ["input_data_b1", "stage_3"]:
                files = _expand_to_xlsm(folder_or_files)
                if not files:
                    return {"ok": False, "message": "No .xlsm files found in the selection."}
                return process_count_rows_b1(
                    file_paths=files,
                    sheet_name=sheet_name,
                )

            # Stage 4: BL creation for multi/single unit files with vessel metadata
            if macro_id in ["so_templates", "stage_4"]:
                return process_file(
                    folder_or_files=folder_or_files,
                    vessel=vessel,
                )
            
            # Stage 5: checker result
            if macro_id in ["checker", "stage_5"]:
                return process_checker_result(
                    folder_or_files=folder_or_files,
                    ask_save_path=lambda name, start: self._ask_save_path(name, start, "xlsx"),
                )
                
            # Stage 6: removing zero
            if macro_id in ["remove_zeros", "stage_6"]:
                return process_remove_zero(folder_or_files=folder_or_files)
            
            # Stage 7: move or copy attach sheet
            if macro_id in ["move_copy_attach", "stage_7"]:
                return process_attach(folder_or_files=folder_or_files)

            # Stage 8: combine all BL files into one workbook
            if macro_id in ["all_bl", "stage_8"]:
                return process_all_bl(
                    folder_or_files=folder_or_files,
                    ask_save_path=self._ask_save_path,
                )

            return {"ok": False, "message": f"Macro '{macro_id}' is not implemented yet."}

        except Exception as exc:
            return {"ok": False, "message": f"Macro '{macro_id}' failed: {exc!r}"}
        finally:
            self._macro_lock.release()









# import os
# import threading
# import webview

# from backend.macros.shipping_list import process_shipping_list
# from backend.macros.copy_IV_base_create_SI import process_si_creation
# from backend.macros.input_data_b1 import process_count_rows_b1
# from backend.macros.bl_creation_single_and_multi_files import process_file
# from backend.macros.checker_result import process_checker_result
# from backend.macros.inf_macro import process_remove_zero
# from backend.macros.move_or_copy_attach import process_attach
# from backend.macros.all_bl import process_all_bl
# from backend.macros.sample import PipelineRunner

# def _expand_to_xlsm(folder_or_files):
#     """Folder, file path, or list of either -> list of .xlsm file paths (no ~$ temp files)."""
#     if isinstance(folder_or_files, (str, os.PathLike)):
#         folder_or_files = [folder_or_files]

#     files = []
#     for item in folder_or_files:
#         item = os.fspath(item)
#         if os.path.isdir(item):
#             for name in sorted(os.listdir(item)):
#                 if name.lower().endswith(".xlsm") and not name.startswith("~$"):
#                     files.append(os.path.join(item, name))
#         elif os.path.isfile(item):
#             files.append(item)
#     return list(dict.fromkeys(files))   # de-duplicate, keep order


# class Api:
#     def __init__(self):
#         self._window = None
#         self._macro_lock = threading.Lock()   # only one macro at a time (Stage 3 drives Excel)
#         self._pipeline = PipelineRunner()
#     def set_window(self, window):
#         self._window = window

#     # ---------- File / Folder Dialogs ----------

#     def select_file(self, file_types=None):
#         file_types = file_types or ("Excel files (*.xlsx;*.xlsm;*.xls)",)
#         result = self._window.create_file_dialog(
#             webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types
#         )
#         return result[0] if result else None

#     def select_files(self, file_types=None):
#         file_types = file_types or ("Excel files (*.xlsx;*.xlsm;*.xls)",)
#         result = self._window.create_file_dialog(
#             webview.OPEN_DIALOG, allow_multiple=True, file_types=file_types
#         )
#         return list(result) if result else []

#     def select_folder(self):
#         result = self._window.create_file_dialog(webview.FOLDER_DIALOG)
#         return result[0] if result else None
    
#     def _ask_save_path(self, suggested_name, start_dir=None, ext="xlsm"):
#         labels = {
#             "xlsm": "Excel Macro Enabled Workbook (*.xlsm)",
#             "xlsx": "Excel Workbook (*.xlsx)",
#         }
#         kwargs = {
#             "save_filename": suggested_name,
#             "file_types": (labels[ext],),
#         }
#         if start_dir and os.path.isdir(start_dir):
#             kwargs["directory"] = start_dir
#         result = self._window.create_file_dialog(webview.SAVE_DIALOG, **kwargs)
#         if isinstance(result, (list, tuple)):          # some pywebview versions return a tuple
#             result = result[0] if result else None
#         return result or None

#     def _confirm(self, title, message):
#         """Yes / No pop-up. If the window cannot show one, carry on (the warnings are still in the log)."""
#         try:
#             return bool(self._window.create_confirmation_dialog(title, message))
#         except Exception:
#             return True
        
#     # helpers for the one-click BL automation
#     def pipeline_preview(self):
#         """Called by app.js to show confirmation counts of old files."""
#         return self._pipeline.preview()

#     def start_pipeline(self, payload):
#         """Triggers the one-click automated multi-stage background run."""
#         files = payload.get("folder_or_files") or payload.get("file_paths", [])
#         sheet_name = payload.get("sheet_name", "BASE")

#         if not self._macro_lock.acquire(blocking=False):
#             return {"ok": False, "message": "Another macro or run is still in progress."}

#         result = self._pipeline.start(
#             files, sheet_name,
#             confirm=self._confirm,
#             on_finish=self._macro_lock.release,
#             output_dir=payload.get("output_dir"),
#         )
#         if not result.get("ok"):
#             self._macro_lock.release()  # Free the lock if it failed to start
#         return result

#     def pipeline_status(self):
#         """Polled by app.js to update frontend UI logs and step indicators."""
#         return self._pipeline.status()
        
    
#     # ---------- Macro Router ----------

#     def run_macro(self, macro_id, payload):
#         folder_or_files = payload.get("folder_or_files") or payload.get("file_paths", [])
#         sheet_name = payload.get("sheet_name", "BASE")

#         if not folder_or_files:
#             return {"ok": False, "message": "No files or folder selected to process."}

#         if not self._macro_lock.acquire(blocking=False):
#             return {"ok": False, "message": "Another macro is still running. Please wait for it to finish."}

#         try:
#             # Stage 1
#             if macro_id == "shipping_list":
#                 template_folder = "../templates/SO.UT.ATTACH.xlsm"
#                 return process_shipping_list(
#                     file_paths=folder_or_files,
#                     sheet_name=sheet_name,
#                     ask_save_path=self._ask_save_path,
#                     confirm=self._confirm,
#                 )

#             # Stage 2: create SI files from the IV files
#             if macro_id in ["copy_iv_base_si", "copy_iv_base_create_si", "stage_2"]:
#                 return process_si_creation(
#                     folder_or_files=folder_or_files,
#                     sheet_name=sheet_name,
#                 )

#             # Stage 3: input the unit count in BASE!B1 (runs through Excel)
#             if macro_id in ["input_data_b1", "stage_3"]:
#                 files = _expand_to_xlsm(folder_or_files)
#                 if not files:
#                     return {"ok": False, "message": "No .xlsm files found in the selection."}
#                 return process_count_rows_b1(
#                     file_paths=files,
#                     sheet_name=sheet_name,
#                 )

#             # Stage 4: BL creation for multi-unit files
#             if macro_id in ["so_templates", "stage_4"]:
#                 return process_file(
#                     folder_or_files=folder_or_files,
#                 )
            
#             # Stage 5: checker result
#             if macro_id in ["checker", "stage_5"]:
#                 return process_checker_result(
#                     folder_or_files=folder_or_files,
#                     ask_save_path=lambda name, start: self._ask_save_path(name, start, "xlsx"),
#                 )
                
#             # Stage 6: removing zero
#             if macro_id in ["remove_zeros", "stage_6"]:
#                 return process_remove_zero(folder_or_files=folder_or_files)
            
#             # Stage 7: move or copy attach sheet
#             if macro_id in ["move_copy_attach", "stage_7"]:
#                 return process_attach(folder_or_files=folder_or_files)

#             # Stage 8: combine all BL files into one workbook
#             if macro_id in ["all_bl", "stage_8"]:
#                 return process_all_bl(
#                     folder_or_files=folder_or_files,
#                     ask_save_path=self._ask_save_path,
#                 )
        
        

#             return {"ok": False, "message": f"Macro '{macro_id}' is not implemented yet."}

#         except Exception as exc:
#             return {"ok": False, "message": f"Macro '{macro_id}' failed: {exc!r}"}
#         finally:
#             self._macro_lock.release()