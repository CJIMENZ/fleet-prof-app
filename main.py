"""Entry point for the modern web-based GUI using PyWebView and Flask."""
import threading
import webview
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
    window = webview.create_window("Fleet Prof App", "http://127.0.0.1:5000", js_api=api)
    webview.start()

if __name__ == "__main__":
    main()
