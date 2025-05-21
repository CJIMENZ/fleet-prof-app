"""Entry point for the modern web-based GUI using PyWebView and Flask."""
import threading
import webview
from webapp.server import app

class WebApi:
    def select_file(self):
        result = webview.windows[0].create_file_dialog(webview.OPEN_DIALOG)
        return result[0] if result else None

    def select_folder(self):
        result = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
        return result[0] if result else None

def start_server():
    app.run(port=5000, threaded=True)

def main():
    t = threading.Thread(target=start_server, daemon=True)
    t.start()
    api = WebApi()
    window = webview.create_window("Fleet Prof App", "http://127.0.0.1:5000", js_api=api)
    webview.start()

if __name__ == "__main__":
    main()
