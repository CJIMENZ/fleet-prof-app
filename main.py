"""
main.py
Entrypoint for the application. Loads config, opens the main GUI window.
"""

from settings_manager import load_config
from gui.main_window import MainWindow

def main():
    config = load_config()  # Load (or create) the config from config.ini
    app = MainWindow(config)  # Pass config to the main window
    app.mainloop()

if __name__ == "__main__":
    main()
