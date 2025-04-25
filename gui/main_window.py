# gui/main_window.py

import sys
import os
import datetime
import logging

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter.filedialog as fd
import tkinter.messagebox as mb

from settings_manager import save_config
from gui.settings_dialog import SettingsDialog

from modules.monthly_workflow         import run_fx_and_comparison
from modules.cks_pivot_operations     import pivot_cks_data_to_ref
from modules.view_download_operations import download_all_views
from modules.report_generation        import build_monthly_database
from modules.pnl_pivot_operations     import generate_pnl_pivot           # NEW
from modules.project_vm_adjustment    import generate_project_vm_adj      # NEW
from modules.unalloc_distribution     import run_unalloc_distribution

class MainWindow(ttk.Window):
    def __init__(self, config_parser):
        self.config_parser = config_parser

        chosen_theme = "journal"
        if "appearance" in config_parser and "theme" in config_parser["appearance"]:
            chosen_theme = config_parser["appearance"]["theme"]

        super().__init__(themename=chosen_theme)
        self.title("My App Main Window")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Variables
        self.oracle_usd_var     = ttk.StringVar()
        self.oracle_cad_var     = ttk.StringVar()
        self.oracle_aud_var     = ttk.StringVar()
        self.finance_report_var = ttk.StringVar()

        self._build_ui()

    def _build_ui(self):
        menubar = ttk.Menu(self)
        file_menu = ttk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Settings...", command=self.open_settings_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing)
        menubar.add_cascade(label="File", menu=file_menu)
        self.config(menu=menubar)

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=BOTH, expand=True)

        row = 0
        # Oracle USD
        ttk.Label(frame, text="Oracle USD:", bootstyle="secondary")\
            .grid(row=row, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.oracle_usd_var, width=50)\
            .grid(row=row, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(frame, text="Browse...", command=self.browse_oracle_usd, bootstyle="secondary")\
            .grid(row=row, column=2, padx=5, pady=5, sticky="w")

        # Oracle CAD
        row += 1
        ttk.Label(frame, text="Oracle CAD:", bootstyle="secondary")\
            .grid(row=row, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.oracle_cad_var, width=50)\
            .grid(row=row, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(frame, text="Browse...", command=self.browse_oracle_cad, bootstyle="secondary")\
            .grid(row=row, column=2, padx=5, pady=5, sticky="w")

        # Oracle AUD
        row += 1
        ttk.Label(frame, text="Oracle AUD:", bootstyle="secondary")\
            .grid(row=row, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.oracle_aud_var, width=50)\
            .grid(row=row, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(frame, text="Browse...", command=self.browse_oracle_aud, bootstyle="secondary")\
            .grid(row=row, column=2, padx=5, pady=5, sticky="w")

        # Month End Finance Report
        row += 1
        ttk.Label(frame, text="Month End Finance Report:", bootstyle="secondary")\
            .grid(row=row, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.finance_report_var, width=50)\
            .grid(row=row, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(frame, text="Browse...", command=self.browse_finance_report, bootstyle="secondary")\
            .grid(row=row, column=2, padx=5, pady=5, sticky="w")

        # Buttons
        row += 1
        ttk.Button(frame, text="Update FX/Compare", bootstyle="primary",
                   command=self.on_update_fx_compare_click)\
            .grid(row=row, column=0, columnspan=3, pady=(10,5))

        row += 1
        ttk.Button(frame, text="Pivot CK Data", bootstyle="info",
                   command=self.on_pivot_cks_click)\
            .grid(row=row, column=0, columnspan=3, pady=(5,5))

        row += 1
        ttk.Button(frame, text="Download Views", bootstyle="warning",
                   command=self.on_download_views_click)\
            .grid(row=row, column=0, columnspan=3, pady=(5,5))

        # NEW: Build Monthly Database button
        row += 1
        ttk.Button(frame, text="Build Monthly Database", bootstyle="success",
                   command=self.on_build_monthly_db_click)\
            .grid(row=row, column=0, columnspan=3, pady=(5,5))

        # NEW: Generate PnL Pivot
        row += 1
        ttk.Button(frame, text="Generate PnL Pivot", bootstyle="secondary",
                   command=self.on_generate_pnl_pivot)\
            .grid(row=row, column=0, columnspan=3, pady=(5,5))

        # NEW: Generate Project VM Adjustment sheet
        row += 1
        ttk.Button(frame, text="Generate Project VM Adj", bootstyle="secondary",
                   command=self.on_generate_project_vm_adj)\
            .grid(row=row, column=0, columnspan=3, pady=(5,10))

        # NEW: Create Unallocated Distributions
        row += 1
        ttk.Button(frame, text="Create Unallocated Distributions", bootstyle="secondary",
                   command=self.on_create_unalloc_distributions)\
            .grid(row=row, column=0, columnspan=3, pady=(5,10))

    # Browse handlers
    def browse_oracle_usd(self):
        path = fd.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")])
        if path: self.oracle_usd_var.set(path)

    def browse_oracle_cad(self):
        path = fd.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")])
        if path: self.oracle_cad_var.set(path)

    def browse_oracle_aud(self):
        path = fd.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")])
        if path: self.oracle_aud_var.set(path)

    def browse_finance_report(self):
        path = fd.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")])
        if path: self.finance_report_var.set(path)

    # Handlers for workflows… (unchanged)
    def on_update_fx_compare_click(self):
        usd_path = self.oracle_usd_var.get()
        cad_path = self.oracle_cad_var.get()
        if not usd_path or not cad_path:
            mb.showerror("Missing Files", "Select both USD and CAD.")
            return
        try:
            ref_fp = self.config_parser["files"]["ref_data_path"]
        except KeyError:
            mb.showerror("Config Error", "No ref_data_path in config.")
            return
        try:
            run_fx_and_comparison(
                config_parser   = self.config_parser,
                oracle_usd_path = usd_path,
                oracle_cad_path = cad_path,
                ref_file_path   = ref_fp,
                oracle_aud_path = self.oracle_aud_var.get() or None
            )
            mb.showinfo("Success", "FX updated and comparison finished.")
        except Exception as e:
            mb.showerror("Error", f"An error occurred:\n{e}")

    def on_pivot_cks_click(self):
        fin = self.finance_report_var.get()
        if not fin:
            mb.showerror("Missing File", "Select the Month End Finance file (CKs).")
            return
        try:
            ref_fp = self.config_parser["files"]["ref_data_path"]
        except KeyError:
            mb.showerror("Config Error", "No ref_data_path in config.")
            return
        now  = datetime.datetime.today()
        last = (now.replace(day=1) - datetime.timedelta(days=1)).strftime("%b-%y")
        try:
            pivot_cks_data_to_ref(
                finance_file     = fin,
                ref_file         = ref_fp,
                target_month_str = last
            )
            mb.showinfo("Success", f"CK Pivot done for {last}.")
        except Exception as e:
            mb.showerror("Error", f"Pivot failed:\n{e}")

    def on_download_views_click(self):
        folder = fd.askdirectory(title="Select folder to save Tableau exports")
        if not folder:
            mb.showerror("No Folder", "Select a folder first.")
            return
        try:
            out = download_all_views(self.config_parser, folder)
            mb.showinfo("Done", f"Views saved to:\n{out}")
        except Exception as e:
            logging.error(e)
            mb.showerror("Error", str(e))

    def on_build_monthly_db_click(self):
        tbl = fd.askopenfilename(title="Select Tableau Exports Workbook",
                                 filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")])
        if not tbl: return
        try: ref_fp = self.config_parser["files"]["ref_data_path"]
        except KeyError: ref_fp = ""
        if not ref_fp:
            ref_fp = fd.askopenfilename(title="Select Reference Data Workbook",
                                         filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")])
            if not ref_fp: return
        out = fd.asksaveasfilename(title="Save Monthly Database As",
                                   defaultextension=".xlsx",
                                   filetypes=[("Excel files","*.xlsx"),("All files","*.*")])
        if not out: return
        try:
            build_monthly_database(
                tableau_exports_path=tbl,
                ref_data_path       =ref_fp,
                output_path         =out
            )
            mb.showinfo("Success", f"Monthly database built:\n{out}")
        except Exception as e:
            logging.error(e)
            mb.showerror("Error", f"Failed to build monthly database:\n{e}")

    # NEW: Generate PnL Pivot
    def on_generate_pnl_pivot(self):
        file = fd.askopenfilename(title="Select MonthDataFile.xlsx",
                                   filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")])
        if not file: return
        try:
            generate_pnl_pivot(file)
            mb.showinfo("Done", "PnL Pivot sheet added.")
        except Exception as e:
            logging.error(e)
            mb.showerror("Error", str(e))

    # NEW: Generate Project VM Adjustment sheet
    def on_generate_project_vm_adj(self):
        """Prompt for the MonthData workbook and build the Project-VM Adj sheet."""

        file = fd.askopenfilename(
            title="Select MonthDataFile.xlsx",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not file:   # user cancelled
            return

        try:
            generate_project_vm_adj(file)          # ← pass the path argument
            mb.showinfo("Success", "Project VM adjustment sheet created.")
        except Exception as e:
            logging.error("Project-VM Adj failed: %s", e, exc_info=True)
            mb.showerror("Error",
                         f"Failed to generate project VM adjustment sheet:\n{e}")

    def on_create_unalloc_distributions(self):
        """Create unallocated distributions for the current workbook."""
        # Helper to shift a date by ±n months
        def shift_month(d: datetime.date, delta: int) -> datetime.date:
            y = d.year + ((d.month - 1 + delta) // 12)
            m = ((d.month - 1 + delta) % 12) + 1
            return d.replace(year=y, month=m, day=1)

        # Build a list of the last 12 months (labels like 'Mar-25')
        today = datetime.date.today()
        prev_end = today.replace(day=1) - datetime.timedelta(days=1)
        prev_start = prev_end.replace(day=1)
        months = [shift_month(prev_start, -i) for i in range(12)]
        labels = [d.strftime('%b-%y') for d in months]

        # Create pop-up dialog
        dlg = ttk.Toplevel(self)
        dlg.title("Select Report Month")
        dlg.transient(self)
        dlg.grab_set()

        # Add content to dialog
        ttk.Label(dlg, text="Report Month:", bootstyle="secondary")\
            .pack(padx=10, pady=(10, 0))
        
        sel_var = ttk.StringVar(value=labels[0])
        combo = ttk.Combobox(dlg, values=labels, textvariable=sel_var, state="readonly", width=10)
        combo.pack(padx=10, pady=5)

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(pady=(0,10))

        def _ok():
            dlg.selected = sel_var.get()
            dlg.destroy()

        ttk.Button(btn_frame, text="OK", command=_ok, bootstyle="primary")\
            .pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy, bootstyle="secondary")\
            .pack(side="left", padx=5)

        self.wait_window(dlg)
        sel = getattr(dlg, 'selected', None)
        if not sel:
            return  # User cancelled

        # Parse 'Mon-YY' into start/end dates
        dt = datetime.datetime.strptime(sel, '%b-%y')
        month_start = dt.date().replace(day=1)
        next_month = shift_month(month_start, 1)
        month_end = next_month - datetime.timedelta(days=1)

        # Get the workbook path
        wb_path = fd.askopenfilename(
            title="Select Workbook",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not wb_path:
            return  # User cancelled

        # Call the distribution routine
        try:
            run_unalloc_distribution(wb_path, month_start, month_end)
            mb.showinfo("Success", "Unallocated distributions created successfully.")
        except Exception as e:
            mb.showerror("Error", f"Failed to create unallocated distributions:\n{str(e)}")

    def open_settings_dialog(self):
        dlg = SettingsDialog(self, self.config_parser)
        self.wait_window(dlg)

    def on_closing(self):
        save_config(self.config_parser)
        self.destroy()
        sys.exit(0)
