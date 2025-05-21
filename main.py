"""Entry point for the modern web-based GUI using PyWebView and Flask."""
import threading
import webview
import os
import sys
from webapp.server import app
from modules.logging_setup import setup_logging, log_user_action

class WebApi:
    def select_file(self):
        result = webview.windows[0].create_file_dialog(webview.OPEN_DIALOG)
        if result:
            log_user_action("File Selected", result[0])
            return result[0]
        return None

    def select_folder(self):
        result = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
        if result:
            log_user_action("Folder Selected", result[0])
            return result[0]
        return None

def start_server():
    app.run(port=5000)

def main():
    # Initialize logging
    log_file = setup_logging()
    log_user_action("Application Started", f"Log file: {log_file}")
    
    t = threading.Thread(target=start_server, daemon=True)
    t.start()
    api = WebApi()

    # Determine screen size for a slightly-less-than-maximized window
    width, height = 1200, 800
    try:
        if webview.screens:
            screen = webview.screens[0]
            screen_w = getattr(screen, 'width', None) or screen['width']
            screen_h = getattr(screen, 'height', None) or screen['height']
            width = int(screen_w * 0.45)
            height = int(screen_h * 0.9)
    except Exception:
        pass

    # Get the path to the icon file
    icon_path = os.path.join(os.path.dirname(__file__), "assets", "app_icon.ico")

    # Create window with minimal required parameters first
    window = webview.create_window(
        title="Fleet Prof App",
        url="http://127.0.0.1:5000",
        js_api=api,
        width=width,
        height=height
    )
    
    # Set icon after window creation if it exists
    if os.path.exists(icon_path) and sys.platform.startswith(("linux", "freebsd")):
        webview.start(icon=icon_path)
    else:
        webview.start()

if __name__ == "__main__":
    main()
