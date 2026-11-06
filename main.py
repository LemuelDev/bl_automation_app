import os
import webview
from backend.bridge import Api

APP_TITLE = "BL Creation App"

def main():
    api = Api()
    frontend_path = os.path.join(os.path.dirname(__file__), "frontend", "index.html")

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

