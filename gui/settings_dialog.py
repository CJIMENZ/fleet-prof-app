"""
gui/settings_dialog.py
A modal dialog to allow the user to edit configuration settings such as
Tableau tokens and file paths for Ref/Master files.
"""

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter.filedialog as fd
import os

class SettingsDialog(ttk.Toplevel):
    def __init__(self, parent, config_parser):
        super().__init__(parent)
        self.title("Settings")
        self.config_parser = config_parser
        self.resizable(False, False)

        self._build_ui()

    def _build_ui(self):
        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=ttk.BOTH, expand=True)

        row = 0

        # ---------------------------
        # Tableau Online Config
        # ---------------------------
        # Safely get the tableau_online section
        if "tableau_online" in self.config_parser:
            tab_cfg = self.config_parser["tableau_online"]
        else:
            tab_cfg = {}

        label_token_name = ttk.Label(frame, text="Tableau Token Name:", bootstyle="secondary")
        label_token_name.grid(row=row, column=0, padx=5, pady=5, sticky="e")

        self.token_name_var = ttk.StringVar(value=tab_cfg.get("personal_access_token_name", ""))
        entry_token_name = ttk.Entry(frame, textvariable=self.token_name_var, width=35)
        entry_token_name.grid(row=row, column=1, padx=5, pady=5)

        row += 1
        label_token_secret = ttk.Label(frame, text="Tableau Token Secret:", bootstyle="secondary")
        label_token_secret.grid(row=row, column=0, padx=5, pady=5, sticky="e")

        self.token_secret_var = ttk.StringVar(value=tab_cfg.get("personal_access_token_secret", ""))
        entry_token_secret = ttk.Entry(frame, textvariable=self.token_secret_var, show="*", width=35)
        entry_token_secret.grid(row=row, column=1, padx=5, pady=5)

        # ---------------------------
        # RefData & MasterFile Paths
        # ---------------------------
        row += 1
        if "files" in self.config_parser:
            files_cfg = self.config_parser["files"]
        else:
            files_cfg = {}

        ref_data_path = files_cfg.get("ref_data_path", "")
        master_file_path = files_cfg.get("master_file_path", "")

        # --- RefData path ---
        label_ref_path = ttk.Label(frame, text="Reference File Path:", bootstyle="secondary")
        label_ref_path.grid(row=row, column=0, padx=5, pady=5, sticky="e")

        self.ref_data_var = ttk.StringVar(value=ref_data_path)
        entry_ref_path = ttk.Entry(frame, textvariable=self.ref_data_var, width=35)
        entry_ref_path.grid(row=row, column=1, padx=5, pady=5, sticky="w")

        btn_browse_ref = ttk.Button(
            frame, text="Browse...", command=self.browse_ref_data, bootstyle="secondary"
        )
        btn_browse_ref.grid(row=row, column=2, padx=5, pady=5, sticky="w")

        row += 1
        # --- MasterFile path ---
        label_master_path = ttk.Label(frame, text="Master File Path:", bootstyle="secondary")
        label_master_path.grid(row=row, column=0, padx=5, pady=5, sticky="e")

        self.master_file_var = ttk.StringVar(value=master_file_path)
        entry_master_path = ttk.Entry(frame, textvariable=self.master_file_var, width=35)
        entry_master_path.grid(row=row, column=1, padx=5, pady=5, sticky="w")

        btn_browse_master = ttk.Button(
            frame, text="Browse...", command=self.browse_master_file, bootstyle="secondary"
        )
        btn_browse_master.grid(row=row, column=2, padx=5, pady=5, sticky="w")

        # ---------------------------
        # Theme Selection
        # ---------------------------
        row += 1
        label_theme = ttk.Label(frame, text="Theme:", bootstyle="secondary")
        label_theme.grid(row=row, column=0, padx=5, pady=5, sticky="e")

        themes = ['flatly', 'journal', 'darkly', 'cyborg', 'superhero', 'cosmo', 'solar', 'united', 'lumen', 'pulse', 'sandstone', 'minty', 'yeti']
        self.theme_var = ttk.StringVar(value=self.config_parser.get('appearance', {}).get('theme', 'journal'))
        theme_combo = ttk.Combobox(frame, textvariable=self.theme_var, values=themes, state='readonly', width=33)
        theme_combo.grid(row=row, column=1, padx=5, pady=5, sticky="w")

        # ---------------------------
        # Save Button
        # ---------------------------
        row += 1
        save_btn = ttk.Button(frame, text="Save", command=self.save_and_close, bootstyle="success")
        save_btn.grid(row=row, column=0, columnspan=3, pady=10)

    def browse_ref_data(self):
        """Open file dialog to select the RefData file."""
        file_path = fd.askopenfilename(
            title="Select Reference File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.ref_data_var.set(file_path)

    def browse_master_file(self):
        """Open file dialog to select the MasterFile."""
        file_path = fd.askopenfilename(
            title="Select Master File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.master_file_var.set(file_path)

    def save_and_close(self):
        # Ensure sections exist
        if "tableau_online" not in self.config_parser:
            self.config_parser.add_section("tableau_online")
        if "files" not in self.config_parser:
            self.config_parser.add_section("files")
        if "appearance" not in self.config_parser:
            self.config_parser.add_section("appearance")

        # Tableau
        self.config_parser["tableau_online"]["personal_access_token_name"] = self.token_name_var.get()
        self.config_parser["tableau_online"]["personal_access_token_secret"] = self.token_secret_var.get()

        # Files
        self.config_parser["files"]["ref_data_path"] = self.ref_data_var.get()
        self.config_parser["files"]["master_file_path"] = self.master_file_var.get()

        # Theme
        self.config_parser["appearance"]["theme"] = self.theme_var.get()

        self.destroy()
