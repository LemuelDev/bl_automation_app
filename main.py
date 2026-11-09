import os
import sys
from pathlib import Path
import webview
from backend.bridge import Api

APP_TITLE = "BL Creation App"

def resource_path(relative_path):
    """ Get the absolute path to a resource, works for both development and PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def ensure_blc_directory_structure() -> dict:
    """
    Checks if the 'BLC ALL MACRO' folder and all corresponding subfolders 
    exist in the current user's Downloads directory. 
    If they do not exist, it creates them automatically.
    """
    downloads_dir = Path.home() / "Downloads"
    root_dir = downloads_dir / "BLC ALL MACRO"
    
    subfolders = [
        root_dir / "WIP BL" / "SEPARATED IV",
        root_dir / "WIP BL" / "BL OUTPUT",
        root_dir / "WIP BL" / "FOR RPA" / "MULTI",
        root_dir / "WIP BL" / "FOR RPA" / "SINGLE",
        root_dir / "WIP BL" / "BL" / "ALL",
        root_dir / "WIP BL" / "BL" / "MULTI",
        root_dir / "WIP BL" / "BL" / "SINGLE",
    ]
    
    created_new = False
    for folder_path in subfolders:
        if not folder_path.exists():
            folder_path.mkdir(parents=True, exist_ok=True)
            created_new = True

    return {
        "success": True,
        "root_path": str(root_dir),
        "created_missing_folders": created_new,
        "message": "Directory structure verified and ready." if not created_new else "Missing workspace folders were automatically created."
    }

def main():
    # 1. Run folder verification and auto-creation immediately on startup
    startup_status = ensure_blc_directory_structure()
    print(f"[Init] {startup_status['message']}")
    print(f"[Init] Workspace Root: {startup_status['root_path']}")

    # 2. Initialize application API and window UI
    api = Api()
    
    # Use resource_path so PyInstaller can locate index.html inside the bundle
    frontend_path = resource_path(os.path.join("frontend", "index.html"))

    window = webview.create_window(
        APP_TITLE,
        frontend_path,
        js_api=api,
        width=1200,
        height=800,
        resizable=False,
        min_size=(1000, 650),
        background_color="#0f1220",
    )
    api.set_window(window)
    webview.start(debug=False)

if __name__ == "__main__":
    main()