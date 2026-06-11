"""
OFAC Scanner GUI - Enhanced Version
- Determinate progress bar (File X of Y)
- Open Output Folder prompt after completion
- All four tabs: Scan, Extract Headers, History, Settings
- Searchable company combobox, sync between tabs, manual/optional auto-refresh
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import csv
import json
import threading
import time
from datetime import datetime, date
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# ScrolledText fallback
try:
    from ttkbootstrap.widgets import ScrolledText
except ImportError:
    from tkinter.scrolledtext import ScrolledText

# -------- Scanner backend ----------
from ofac_scanner_core import (
    ENV_VAR_FOLDER, ENV_VAR_CSV, APP_NAME, SCRIPT_PATH,
    COMPANY_HEADERS_PATH, BASE_OUTPUT_FOLDER, SEVEN_ZIP_PATH,
    FILE_EXTENSIONS_EXCEL, FILE_EXTENSIONS_TEXT, FILE_EXTENSIONS_ARCHIVE,
    OUTPUT_COLUMNS, LOG_HEADERS,
    set_settings_file, load_settings, save_settings,
    instance_check_or_focus, get_env_var, set_user_env_var,
    ScannerJob,
    process_files_direct,
    FolderHandler, start_watching, process_config_and_scan
)

# -------- Header extraction backend ----------
from header_extractor_core import (
    HeaderExtractionJob,
    process_header_extraction
)

# ---------- Settings file path ----------
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ofac_settings.json")
set_settings_file(SETTINGS_FILE)

SETUP_WINDOW_TITLE = "OFAC Scanner Setup"
SCANNER_WINDOW_TITLE = "OFAC Scanner - Enhanced"

# ==================== SETUP GUI ====================
class SetupGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(SETUP_WINDOW_TITLE)
        self.root.geometry("650x500")
        self.root.resizable(False, False)
        self.main_frame = tb.Frame(self.root, padding=30)
        self.main_frame.pack(fill=BOTH, expand=YES)

        tb.Label(self.main_frame, text="OFAC Scanner Setup", font=("Helvetica", 18, "bold")).pack(pady=15)
        tb.Label(self.main_frame, text="Watch Folder Path:", font=("Helvetica", 10)).pack(anchor=W, pady=(10,0))
        folder_frame = tb.Frame(self.main_frame)
        folder_frame.pack(fill=X, pady=5)
        self.folder_var = tk.StringVar()
        tb.Entry(folder_frame, textvariable=self.folder_var, bootstyle="info").pack(side=LEFT, fill=X, expand=True, padx=(0,5))
        tb.Button(folder_frame, text="Browse", command=self.browse_folder, bootstyle="secondary").pack(side=RIGHT)

        tb.Label(self.main_frame, text="Company Passwords CSV Path:", font=("Helvetica", 10)).pack(anchor=W, pady=(10,0))
        csv_frame = tb.Frame(self.main_frame)
        csv_frame.pack(fill=X, pady=5)
        self.csv_var = tk.StringVar()
        tb.Entry(csv_frame, textvariable=self.csv_var, bootstyle="info").pack(side=LEFT, fill=X, expand=True, padx=(0,5))
        tb.Button(csv_frame, text="Browse", command=self.browse_csv, bootstyle="secondary").pack(side=RIGHT)

        info = tb.Label(self.main_frame, text="The watch folder will be monitored for new config files.\nPassword CSV must have columns: Code, Password",
                        font=("Helvetica", 9), foreground="gray")
        info.pack(pady=20)

        btn_frame = tb.Frame(self.main_frame)
        btn_frame.pack(fill=X, pady=30)
        self.save_btn = tb.Button(btn_frame, text="Save & Launch Scanner", command=self.save_and_launch,
                                  bootstyle="success", width=25)
        self.save_btn.pack()
        self.status_label = tb.Label(self.main_frame, text="", font=("Helvetica", 9))
        self.status_label.pack(pady=5)

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_var.set(path)
            self.status_label.config(text=f"Selected: {path}")

    def browse_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path:
            self.csv_var.set(path)
            self.status_label.config(text=f"Selected: {path}")

    def save_and_launch(self):
        folder = self.folder_var.get()
        csv_path = self.csv_var.get()
        if not folder or not csv_path:
            messagebox.showerror("Error", "Both fields are required")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("Error", f"Folder does not exist: {folder}")
            return
        if not os.path.isfile(csv_path):
            messagebox.showerror("Error", f"CSV file does not exist: {csv_path}")
            return

        settings = load_settings()
        settings["folder"] = folder
        settings["csv"] = csv_path
        save_settings(settings)

        messagebox.showinfo("Success", "Settings saved! The scanner will now launch.")
        self.main_frame.destroy()
        EnhancedScannerGUI(self.root, folder_path=folder, csv_path=csv_path)


# ==================== MAIN SCANNER GUI ====================
class EnhancedScannerGUI:
    def __init__(self, root, folder_path=None, csv_path=None):
        self.root = root
        self.root.title(SCANNER_WINDOW_TITLE)
        try:
            self.root.state('zoomed')   # Windows
        except:
            self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}")
        self.root.minsize(800, 600)

        saved = load_settings()
        theme = saved.get("theme", "flatly")
        try:
            self.root.style.theme_use(theme)
        except:
            pass

        self.folder_path = folder_path or get_env_var(ENV_VAR_FOLDER)
        self.csv_path = csv_path or get_env_var(ENV_VAR_CSV)

        if not self.folder_path or not self.csv_path:
            messagebox.showerror("Error", "Configuration missing. Run setup first.")
            root.destroy()
            return

        self.company_data = self.load_csv()
        self.scan_thread = None
        self.stop_requested = False
        self.selected_files_set = set()
        self.ext_selected_files_set = set()

        # Shared company code variable for synchronization
        self.current_company_code = tk.StringVar()
        self.current_company_code.trace('w', self.on_company_code_changed)

        self.notebook = tb.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=YES, padx=10, pady=10)

        # --- Scan Tab ---
        self.scan_tab = tb.Frame(self.notebook)
        self.notebook.add(self.scan_tab, text="Scan")
        self.build_scan_tab()

        # --- Extract Headers Tab ---
        self.extract_tab = tb.Frame(self.notebook)
        self.notebook.add(self.extract_tab, text="Extract Headers")
        self.build_extract_headers_tab()

        # --- History Tab ---
        self.history_tab = tb.Frame(self.notebook)
        self.notebook.add(self.history_tab, text="History")
        self.build_history_tab()

        # --- Settings Tab ---
        self.settings_tab = tb.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="Settings")
        self.build_settings_tab()

        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = tb.Label(self.root, textvariable=self.status_var, bootstyle="info", anchor=W)
        self.status_bar.pack(side=BOTTOM, fill=X)

        self.progress = tb.Progressbar(self.root, bootstyle="success", mode="indeterminate")
        self.progress.pack(side=BOTTOM, fill=X, padx=10, pady=5)
        self.progress.pack_forget()

        self.refresh_file_list()
        self.refresh_extract_file_list()
        # Start auto-refresh loop (will check checkbox)
        self.root.after(1000, self._auto_refresh_loop)

    # ---------- CSV loading ----------
    def load_csv(self):
        data = []
        if not os.path.exists(self.csv_path):
            return data
        try:
            with open(self.csv_path, 'r', newline='', encoding='utf-8-sig') as f:
                sample = f.read(8192)
                f.seek(0)
                has_header = csv.Sniffer().has_header(sample)
                if has_header:
                    reader = csv.DictReader(f)
                    if reader.fieldnames:
                        code_key = next((k for k in reader.fieldnames if k.lower() == 'code'), None)
                        pwd_key = next((k for k in reader.fieldnames if k.lower() == 'password'), None)
                        if code_key and pwd_key:
                            for row in reader:
                                data.append({'Code': row[code_key].strip(), 'Password': row[pwd_key].strip()})
                        else:
                            f.seek(0)
                            next(reader, None)
                            for row in reader:
                                if len(row) >= 2:
                                    data.append({'Code': row[0].strip(), 'Password': row[1].strip()})
                else:
                    reader = csv.reader(f)
                    for row in reader:
                        if len(row) >= 2:
                            data.append({'Code': row[0].strip(), 'Password': row[1].strip()})
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV: {e}")
        return data

    def append_csv(self, new_row):
        try:
            with open(self.csv_path, 'a', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['Code', 'Password'])
                writer.writerow(new_row)
        except:
            pass

    # ---------- Searchable combobox factory ----------
    def create_searchable_combobox(self, parent, label_text, values, variable):
        """Returns a frame containing a label, entry, and listbox for searchable dropdown."""
        frame = tb.Frame(parent)
        tb.Label(frame, text=label_text).pack(side=tk.LEFT, padx=5)

        entry_var = tk.StringVar()
        entry = tb.Entry(frame, textvariable=entry_var, bootstyle="info")
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # Listbox popup (initially hidden)
        listbox_frame = tb.Frame(parent, bootstyle="light")
        listbox = tk.Listbox(listbox_frame, height=6, exportselection=False)
        scrollbar = tb.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        listbox_frame.pack_forget()

        # Store values list for dynamic updates
        frame.values = list(values)

        def update_listbox(event=None):
            search_term = entry_var.get().strip().lower()
            filtered = [v for v in frame.values if search_term in v.lower()] if search_term else frame.values
            listbox.delete(0, tk.END)
            for item in filtered:
                listbox.insert(tk.END, item)
            if filtered:
                listbox.selection_set(0)
                listbox.see(0)
            # Show/hide listbox only if entry has focus and there are items
            if filtered and entry.focus_get() == entry:
                # Place listbox_frame below the frame
                listbox_frame.pack(fill=tk.X, pady=(0,5))
            else:
                listbox_frame.pack_forget()

        def on_entry_focus_in(event):
            update_listbox()

        def on_entry_focus_out(event):
            self.root.after(200, lambda: listbox_frame.pack_forget() if listbox.focus_get() != listbox else None)

        def on_listbox_select(event):
            if listbox.curselection():
                selected = listbox.get(listbox.curselection()[0])
                entry_var.set(selected)
                variable.set(selected)
                listbox_frame.pack_forget()

        entry.bind("<KeyRelease>", update_listbox)
        entry.bind("<FocusIn>", on_entry_focus_in)
        entry.bind("<FocusOut>", on_entry_focus_out)
        listbox.bind("<<ListboxSelect>>", on_listbox_select)
        listbox.bind("<FocusOut>", lambda e: self.root.after(200, listbox_frame.pack_forget))

        # Store references for later updates
        frame.entry = entry
        frame.entry_var = entry_var
        frame.listbox_frame = listbox_frame
        frame.listbox = listbox
        return frame

    def update_searchable_combobox_values(self, combobox_frame, new_values):
        """Update the internal values list of a searchable combobox and refresh display."""
        combobox_frame.values = list(new_values)
        # Trigger update on the entry (clears filter)
        combobox_frame.entry_var.set("")
        combobox_frame.entry.event_generate('<KeyRelease>')

    # ---------- Company code sync ----------
    def on_company_code_changed(self, *args):
        """Sync company code between tabs."""
        code = self.current_company_code.get()
        # Update the underlying StringVars used by password selection
        if hasattr(self, 'code_var'):
            self.code_var.set(code)
        if hasattr(self, 'ext_code_var'):
            self.ext_code_var.set(code)
        # Refresh password lists
        if hasattr(self, 'update_password_options'):
            self.update_password_options()
        if hasattr(self, 'ext_update_password_options'):
            self.ext_update_password_options()

    # ---------- Build Scan Tab ----------
    def build_scan_tab(self):
        canvas = tk.Canvas(self.scan_tab, highlightthickness=0)
        scrollbar = tb.Scrollbar(self.scan_tab, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        content_frame = tb.Frame(canvas)
        content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=content_frame, anchor="nw")

        self.scan_tab.grid_rowconfigure(0, weight=1)
        self.scan_tab.grid_columnconfigure(0, weight=1)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        def on_canvas_configure(event):
            canvas.itemconfig(1, width=event.width)
        canvas.bind("<Configure>", on_canvas_configure)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel, add="+")

        # File Filters
        filter_frame = tb.LabelFrame(content_frame, text="File Filters")
        filter_frame.pack(fill=X, padx=10, pady=(5, 2))
        self.filter_excel = tk.BooleanVar(value=True)
        self.filter_csv = tk.BooleanVar(value=True)
        self.filter_archive = tk.BooleanVar(value=True)
        cb_frame = tb.Frame(filter_frame)
        cb_frame.pack(fill=X, padx=5, pady=2)
        tb.Checkbutton(cb_frame, text="Excel (.xlsx, .xls, .xlsm, .xlsb)",
                       variable=self.filter_excel, bootstyle="primary").pack(anchor="w", pady=1)
        tb.Checkbutton(cb_frame, text="CSV/Text (.csv, .txt, .rpt)",
                       variable=self.filter_csv, bootstyle="primary").pack(anchor="w", pady=1)
        tb.Checkbutton(cb_frame, text="Archives (.zip, .7z, .rar, .tar)",
                       variable=self.filter_archive, bootstyle="primary").pack(anchor="w", pady=1)

        btn_frame = tb.Frame(filter_frame)
        btn_frame.pack(fill=X, pady=5)
        tb.Button(btn_frame, text="Refresh List", command=self.refresh_file_list, bootstyle="secondary").pack(side=tk.LEFT, padx=5)
        self.auto_refresh_var = tk.BooleanVar(value=False)
        tb.Checkbutton(btn_frame, text="Auto-refresh every 2 sec", variable=self.auto_refresh_var, bootstyle="primary").pack(side=tk.LEFT, padx=10)

        # File List
        list_frame = tb.LabelFrame(content_frame, text="Files in Watch Folder")
        list_frame.pack(fill=BOTH, expand=YES, padx=10, pady=2)
        self.file_canvas = tk.Canvas(list_frame, highlightthickness=0)
        file_scrollbar = tb.Scrollbar(list_frame, orient="vertical", command=self.file_canvas.yview)
        self.file_canvas.configure(yscrollcommand=file_scrollbar.set)
        self.file_inner = tb.Frame(self.file_canvas)
        self.file_canvas.create_window((0, 0), window=self.file_inner, anchor="nw")
        self.file_inner.bind("<Configure>", lambda e: self.file_canvas.configure(scrollregion=self.file_canvas.bbox("all")))
        self.file_canvas.bind("<Configure>", lambda e: self.file_canvas.itemconfig(1, width=e.width))
        self.file_canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        file_scrollbar.pack(side=RIGHT, fill=Y)
        self.file_vars = {}

        select_frame = tb.Frame(content_frame)
        select_frame.pack(fill=X, padx=10, pady=2)
        tb.Button(select_frame, text="Select All", command=self.select_all_files_checkbox,
                  bootstyle="info").pack(side=LEFT, padx=5)
        tb.Button(select_frame, text="Clear All", command=self.clear_selection_checkbox,
                  bootstyle="secondary").pack(side=LEFT, padx=5)

        # Scan Configuration
        config_frame = tb.LabelFrame(content_frame, text="Scan Configuration")
        config_frame.pack(fill=X, padx=10, pady=2)

        row1 = tb.Frame(config_frame)
        row1.pack(fill=X, padx=5, pady=(5, 2))
        # Searchable company combobox
        company_codes = sorted(set([d['Code'] for d in self.company_data]))
        self.searchable_company_scan = self.create_searchable_combobox(
            row1, "Company Code:", company_codes, self.current_company_code
        )
        self.searchable_company_scan.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        tb.Button(row1, text="+ New Code", command=self.add_new_code, bootstyle="success").pack(side=tk.LEFT, padx=5)

        # For backward compatibility with password methods, keep code_var
        self.code_var = tk.StringVar()
        self.current_company_code.trace('w', lambda *a: self.code_var.set(self.current_company_code.get()))
        self.code_var.set(self.current_company_code.get())

        pass_frame = tb.LabelFrame(config_frame, text="Passwords")
        pass_frame.pack(fill=BOTH, expand=YES, padx=5, pady=2)

        search_frame = tb.Frame(pass_frame)
        search_frame.pack(fill=X, padx=5, pady=2)
        tb.Label(search_frame, text="Search:").pack(side=LEFT)
        self.pass_search_var = tk.StringVar()
        self.pass_search_var.trace("w", lambda *a: self.filter_password_list())
        pass_search_entry = tb.Entry(search_frame, textvariable=self.pass_search_var, bootstyle="info")
        pass_search_entry.pack(side=LEFT, fill=X, expand=YES, padx=5)

        pass_canvas = tk.Canvas(pass_frame, highlightthickness=0, height=100)
        pass_scrollbar = tb.Scrollbar(pass_frame, orient="vertical", command=pass_canvas.yview)
        pass_canvas.configure(yscrollcommand=pass_scrollbar.set)

        self.pass_inner = tb.Frame(pass_canvas)
        pass_canvas.create_window((0, 0), window=self.pass_inner, anchor="nw")
        self.pass_inner.bind("<Configure>", lambda e: pass_canvas.configure(scrollregion=pass_canvas.bbox("all")))
        pass_canvas.bind("<Configure>", lambda e: pass_canvas.itemconfig(1, width=e.width))

        pass_canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        pass_scrollbar.pack(side=RIGHT, fill=Y)

        self.pass_vars = {}
        self.all_passwords = []

        btn_frame2 = tb.Frame(pass_frame)
        btn_frame2.pack(fill=X, pady=2)
        tb.Button(btn_frame2, text="Select All", command=self.select_all_passwords, bootstyle="info").pack(side=LEFT, padx=5)
        tb.Button(btn_frame2, text="Clear All", command=self.clear_all_passwords, bootstyle="secondary").pack(side=LEFT, padx=5)
        tb.Button(btn_frame2, text="+ New Password", command=self.add_new_pass, bootstyle="success").pack(side=RIGHT, padx=5)

        # Date picker
        row3 = tb.Frame(config_frame)
        row3.pack(fill=X, padx=5, pady=(5, 5))
        tb.Label(row3, text="Email Received Date:").pack(side=LEFT, padx=5)
        try:
            from ttkbootstrap.widgets import DateEntry
            self.date_entry = DateEntry(row3, width=12)
            self.date_entry.set_date(date.today())
        except ImportError:
            self.date_entry = tk.Entry(row3, width=12)
            self.date_entry.insert(0, date.today().isoformat())
        self.date_entry.pack(side=LEFT, padx=5)

        # Action Buttons
        action_frame = tb.Frame(content_frame)
        action_frame.pack(fill=X, padx=10, pady=(5, 10))
        tb.Button(action_frame, text="Queue for Watcher", command=self.queue_for_watcher,
                  bootstyle="warning", width=20).pack(side=LEFT, padx=5)
        tb.Button(action_frame, text="Run Scanner Now", command=self.run_scanner_now,
                  bootstyle="primary", width=20).pack(side=LEFT, padx=5)
        self.stop_btn = tb.Button(action_frame, text="Stop Scan", command=self.stop_scan,
                                  bootstyle="danger", width=20, state=DISABLED)
        self.stop_btn.pack(side=LEFT, padx=5)

        # Log Output
        log_frame = tb.LabelFrame(content_frame, text="Log Output")
        log_frame.pack(fill=BOTH, expand=YES, padx=10, pady=(5, 10))
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)
        self.log_text = ScrolledText(log_frame, height=8, wrap=WORD)
        self.log_text.pack(fill=BOTH, expand=YES)

    # ---------- Build Extract Headers Tab ----------
    def build_extract_headers_tab(self):
        canvas = tk.Canvas(self.extract_tab, highlightthickness=0)
        scrollbar = tb.Scrollbar(self.extract_tab, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        content_frame = tb.Frame(canvas)
        content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=content_frame, anchor="nw")

        self.extract_tab.grid_rowconfigure(0, weight=1)
        self.extract_tab.grid_columnconfigure(0, weight=1)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        def on_canvas_configure(event):
            canvas.itemconfig(1, width=event.width)
        canvas.bind("<Configure>", on_canvas_configure)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel, add="+")

        # File Filters
        filter_frame = tb.LabelFrame(content_frame, text="File Filters")
        filter_frame.pack(fill=X, padx=10, pady=(5, 2))
        self.ext_filter_excel = tk.BooleanVar(value=True)
        self.ext_filter_csv = tk.BooleanVar(value=True)
        self.ext_filter_archive = tk.BooleanVar(value=True)
        cb_frame = tb.Frame(filter_frame)
        cb_frame.pack(fill=X, padx=5, pady=2)
        tb.Checkbutton(cb_frame, text="Excel (.xlsx, .xls, .xlsm, .xlsb)",
                       variable=self.ext_filter_excel, bootstyle="primary").pack(anchor="w", pady=1)
        tb.Checkbutton(cb_frame, text="CSV/Text (.csv, .txt, .rpt)",
                       variable=self.ext_filter_csv, bootstyle="primary").pack(anchor="w", pady=1)
        tb.Checkbutton(cb_frame, text="Archives (.zip, .7z, .rar, .tar)",
                       variable=self.ext_filter_archive, bootstyle="primary").pack(anchor="w", pady=1)

        btn_frame = tb.Frame(filter_frame)
        btn_frame.pack(fill=X, pady=5)
        tb.Button(btn_frame, text="Refresh List", command=self.refresh_extract_file_list, bootstyle="secondary").pack(side=tk.LEFT, padx=5)
        # Auto-refresh checkbox already present; no need for another one here

        # File List
        list_frame = tb.LabelFrame(content_frame, text="Files in Watch Folder")
        list_frame.pack(fill=BOTH, expand=YES, padx=10, pady=2)
        self.ext_file_canvas = tk.Canvas(list_frame, highlightthickness=0)
        ext_file_scrollbar = tb.Scrollbar(list_frame, orient="vertical", command=self.ext_file_canvas.yview)
        self.ext_file_canvas.configure(yscrollcommand=ext_file_scrollbar.set)
        self.ext_file_inner = tb.Frame(self.ext_file_canvas)
        self.ext_file_canvas.create_window((0, 0), window=self.ext_file_inner, anchor="nw")
        self.ext_file_inner.bind("<Configure>", lambda e: self.ext_file_canvas.configure(scrollregion=self.ext_file_canvas.bbox("all")))
        self.ext_file_canvas.bind("<Configure>", lambda e: self.ext_file_canvas.itemconfig(1, width=e.width))
        self.ext_file_canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        ext_file_scrollbar.pack(side=RIGHT, fill=Y)
        self.ext_file_vars = {}

        select_frame = tb.Frame(content_frame)
        select_frame.pack(fill=X, padx=10, pady=2)
        tb.Button(select_frame, text="Select All", command=self.ext_select_all_files,
                  bootstyle="info").pack(side=LEFT, padx=5)
        tb.Button(select_frame, text="Clear All", command=self.ext_clear_all_files,
                  bootstyle="secondary").pack(side=LEFT, padx=5)

        # Configuration
        config_frame = tb.LabelFrame(content_frame, text="Extraction Configuration")
        config_frame.pack(fill=X, padx=10, pady=2)

        row1 = tb.Frame(config_frame)
        row1.pack(fill=X, padx=5, pady=(5, 2))
        # Searchable company combobox for extract tab
        company_codes = sorted(set([d['Code'] for d in self.company_data]))
        self.searchable_company_extract = self.create_searchable_combobox(
            row1, "Company Code:", company_codes, self.current_company_code
        )
        self.searchable_company_extract.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        tb.Button(row1, text="+ New Code", command=self.add_new_code, bootstyle="success").pack(side=tk.LEFT, padx=5)

        self.ext_code_var = tk.StringVar()
        self.current_company_code.trace('w', lambda *a: self.ext_code_var.set(self.current_company_code.get()))
        self.ext_code_var.set(self.current_company_code.get())

        pass_frame = tb.LabelFrame(config_frame, text="Passwords")
        pass_frame.pack(fill=BOTH, expand=YES, padx=5, pady=2)

        search_frame = tb.Frame(pass_frame)
        search_frame.pack(fill=X, padx=5, pady=2)
        tb.Label(search_frame, text="Search:").pack(side=LEFT)
        self.ext_pass_search_var = tk.StringVar()
        self.ext_pass_search_var.trace("w", lambda *a: self.ext_filter_password_list())
        tb.Entry(search_frame, textvariable=self.ext_pass_search_var, bootstyle="info").pack(side=LEFT, fill=X, expand=YES, padx=5)

        pass_canvas = tk.Canvas(pass_frame, highlightthickness=0, height=100)
        pass_scrollbar = tb.Scrollbar(pass_frame, orient="vertical", command=pass_canvas.yview)
        pass_canvas.configure(yscrollcommand=pass_scrollbar.set)

        self.ext_pass_inner = tb.Frame(pass_canvas)
        pass_canvas.create_window((0, 0), window=self.ext_pass_inner, anchor="nw")
        self.ext_pass_inner.bind("<Configure>", lambda e: pass_canvas.configure(scrollregion=pass_canvas.bbox("all")))
        pass_canvas.bind("<Configure>", lambda e: pass_canvas.itemconfig(1, width=e.width))

        pass_canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        pass_scrollbar.pack(side=RIGHT, fill=Y)
        self.ext_pass_vars = {}
        self.ext_all_passwords = []

        btn_frame2 = tb.Frame(pass_frame)
        btn_frame2.pack(fill=X, pady=2)
        tb.Button(btn_frame2, text="Select All", command=self.ext_select_all_passwords, bootstyle="info").pack(side=LEFT, padx=5)
        tb.Button(btn_frame2, text="Clear All", command=self.ext_clear_all_passwords, bootstyle="secondary").pack(side=LEFT, padx=5)
        tb.Button(btn_frame2, text="+ New Password", command=self.add_new_pass, bootstyle="success").pack(side=RIGHT, padx=5)

        row3 = tb.Frame(config_frame)
        row3.pack(fill=X, padx=5, pady=(5, 5))
        tb.Label(row3, text="Extraction Date:").pack(side=LEFT, padx=5)
        try:
            from ttkbootstrap.widgets import DateEntry
            self.ext_date_entry = DateEntry(row3, width=12)
            self.ext_date_entry.set_date(date.today())
        except ImportError:
            self.ext_date_entry = tk.Entry(row3, width=12)
            self.ext_date_entry.insert(0, date.today().isoformat())
        self.ext_date_entry.pack(side=LEFT, padx=5)

        action_frame = tb.Frame(content_frame)
        action_frame.pack(fill=X, padx=10, pady=(5, 10))
        tb.Button(action_frame, text="Extract Headers", command=self.run_header_extraction_now,
                  bootstyle="success", width=20).pack(side=LEFT, padx=5)

        log_frame = tb.LabelFrame(content_frame, text="Log Output")
        log_frame.pack(fill=BOTH, expand=YES, padx=10, pady=(5, 10))
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)
        self.ext_log_text = ScrolledText(log_frame, height=8, wrap=WORD)
        self.ext_log_text.pack(fill=BOTH, expand=YES)

    # ---------- History Tab ----------
    def build_history_tab(self):
        frame = tb.Frame(self.history_tab)
        frame.pack(fill=BOTH, expand=YES, padx=10, pady=10)
        tb.Button(frame, text="Load Recent Logs", command=self.load_history, bootstyle="primary").pack(pady=5)
        columns = ("date", "company", "files", "output")
        self.history_tree = tb.Treeview(frame, columns=columns, show="headings", bootstyle="primary")
        self.history_tree.heading("date", text="Scan Date")
        self.history_tree.heading("company", text="Company")
        self.history_tree.heading("files", text="Files Processed")
        self.history_tree.heading("output", text="Output Folder")
        self.history_tree.pack(fill=BOTH, expand=YES)
        scroll = tb.Scrollbar(frame, orient=VERTICAL, command=self.history_tree.yview)
        scroll.pack(side=RIGHT, fill=Y)
        self.history_tree.configure(yscrollcommand=scroll.set)

    # ---------- Settings Tab ----------
    def build_settings_tab(self):
        frame = tb.Frame(self.settings_tab)
        frame.pack(fill=BOTH, expand=YES, padx=20, pady=20)
        tb.Label(frame, text="Theme", font=("Helvetica", 12, "bold")).pack(anchor=W)
        saved = load_settings()
        current_theme = saved.get("theme", "flatly")
        self.theme_var = tk.StringVar(value=current_theme)
        themes = ["flatly", "darkly", "cyborg", "solar", "superhero", "journal", "litera", "lumen", "minty", "pulse", "sandstone", "simplex", "spacelab", "united", "yeti"]
        theme_combo = tb.Combobox(frame, values=themes, textvariable=self.theme_var, bootstyle="primary")
        theme_combo.pack(fill=X, pady=5)
        tb.Button(frame, text="Apply Theme", command=self.change_theme, bootstyle="info").pack(pady=5)
        tb.Label(frame, text="Watch Folder", font=("Helvetica", 12, "bold")).pack(anchor=W, pady=(20,0))
        watch_frame = tb.Frame(frame)
        watch_frame.pack(fill=X, pady=5)
        self.watch_path_var = tk.StringVar(value=self.folder_path)
        tb.Entry(watch_frame, textvariable=self.watch_path_var).pack(side=LEFT, fill=X, expand=True, padx=(0,5))
        tb.Button(watch_frame, text="Change", command=self.change_watch_folder, bootstyle="warning").pack(side=RIGHT)
        tb.Label(frame, text="Passwords CSV", font=("Helvetica", 12, "bold")).pack(anchor=W, pady=(20,0))
        csv_frame = tb.Frame(frame)
        csv_frame.pack(fill=X, pady=5)
        self.csv_path_var = tk.StringVar(value=self.csv_path)
        tb.Entry(csv_frame, textvariable=self.csv_path_var).pack(side=LEFT, fill=X, expand=True, padx=(0,5))
        tb.Button(csv_frame, text="Change", command=self.change_csv_file, bootstyle="warning").pack(side=RIGHT)
        tb.Button(frame, text="Save Settings", command=self.save_settings, bootstyle="success").pack(pady=20)

    # ---------- File list refresh for Scan tab ----------
    def refresh_file_list(self):
        if not self.folder_path:
            return
        self.selected_files_set = {f for f, var in self.file_vars.items() if var.get() == 1}
        for widget in self.file_inner.winfo_children():
            widget.destroy()
        self.file_vars.clear()
        try:
            files_info = []
            for f in os.listdir(self.folder_path):
                full = os.path.join(self.folder_path, f)
                if os.path.isfile(full) and not f.endswith('.json'):
                    ext = os.path.splitext(f)[1].lower().replace('.', '')
                    if ext in FILE_EXTENSIONS_EXCEL and not self.filter_excel.get():
                        continue
                    if ext in FILE_EXTENSIONS_TEXT and not self.filter_csv.get():
                        continue
                    if ext in FILE_EXTENSIONS_ARCHIVE and not self.filter_archive.get():
                        continue
                    size = os.path.getsize(full) // 1024
                    mod_time = datetime.fromtimestamp(os.path.getmtime(full)).strftime("%Y-%m-%d %H:%M:%S")
                    ftype = "Excel" if ext in FILE_EXTENSIONS_EXCEL else "Text" if ext in FILE_EXTENSIONS_TEXT else "Archive"
                    files_info.append((f, size, mod_time, ftype))
            files_info.sort(key=lambda x: x[0])
            for fname, size, mod_time, ftype in files_info:
                checked = 1 if fname in self.selected_files_set else 0
                var = tk.IntVar(value=checked)
                self.file_vars[fname] = var
                row_frame = tb.Frame(self.file_inner)
                row_frame.pack(fill=X, pady=2)
                tb.Checkbutton(row_frame, variable=var, bootstyle="primary").pack(side=LEFT, padx=5)
                tb.Label(row_frame, text=fname, font=("Helvetica", 10, "bold"), anchor="w").pack(side=LEFT, fill=X, expand=True, padx=5)
                tb.Label(row_frame, text=f"{size} KB  |  {mod_time}  |  {ftype}", font=("Helvetica", 9), foreground="gray").pack(side=RIGHT, padx=5)
        except Exception as e:
            self.log(f"Refresh error: {e}")

    # ---------- File list refresh for Extract tab ----------
    def refresh_extract_file_list(self):
        if not self.folder_path:
            return
        self.ext_selected_files_set = {f for f, var in self.ext_file_vars.items() if var.get() == 1}
        for widget in self.ext_file_inner.winfo_children():
            widget.destroy()
        self.ext_file_vars.clear()
        try:
            files_info = []
            for f in os.listdir(self.folder_path):
                full = os.path.join(self.folder_path, f)
                if os.path.isfile(full) and not f.endswith('.json'):
                    ext = os.path.splitext(f)[1].lower().replace('.', '')
                    if ext in FILE_EXTENSIONS_EXCEL and not self.ext_filter_excel.get():
                        continue
                    if ext in FILE_EXTENSIONS_TEXT and not self.ext_filter_csv.get():
                        continue
                    if ext in FILE_EXTENSIONS_ARCHIVE and not self.ext_filter_archive.get():
                        continue
                    size = os.path.getsize(full) // 1024
                    mod_time = datetime.fromtimestamp(os.path.getmtime(full)).strftime("%Y-%m-%d %H:%M:%S")
                    ftype = "Excel" if ext in FILE_EXTENSIONS_EXCEL else "Text" if ext in FILE_EXTENSIONS_TEXT else "Archive"
                    files_info.append((f, size, mod_time, ftype))
            files_info.sort(key=lambda x: x[0])
            for fname, size, mod_time, ftype in files_info:
                checked = 1 if fname in self.ext_selected_files_set else 0
                var = tk.IntVar(value=checked)
                self.ext_file_vars[fname] = var
                row_frame = tb.Frame(self.ext_file_inner)
                row_frame.pack(fill=X, pady=2)
                tb.Checkbutton(row_frame, variable=var, bootstyle="primary").pack(side=LEFT, padx=5)
                tb.Label(row_frame, text=fname, font=("Helvetica", 10, "bold"), anchor="w").pack(side=LEFT, fill=X, expand=True, padx=5)
                tb.Label(row_frame, text=f"{size} KB  |  {mod_time}  |  {ftype}", font=("Helvetica", 9), foreground="gray").pack(side=RIGHT, padx=5)
        except Exception as e:
            self.ext_log(f"Refresh error: {e}")

    # ---------- Scan tab selections ----------
    def get_selected_files(self):
        return [fname for fname, var in self.file_vars.items() if var.get() == 1]

    def select_all_files_checkbox(self):
        for var in self.file_vars.values():
            var.set(1)

    def clear_selection_checkbox(self):
        for var in self.file_vars.values():
            var.set(0)

    # ---------- Extract tab selections ----------
    def ext_get_selected_files(self):
        return [fname for fname, var in self.ext_file_vars.items() if var.get() == 1]

    def ext_select_all_files(self):
        for var in self.ext_file_vars.values():
            var.set(1)

    def ext_clear_all_files(self):
        for var in self.ext_file_vars.values():
            var.set(0)

    # ---------- Password handling for Scan tab ----------
    def update_password_options(self, event=None):
        selected = self.code_var.get()
        self.all_passwords = [d['Password'] for d in self.company_data if d['Code'] == selected]
        self.filter_password_list()

    def filter_password_list(self):
        search_term = self.pass_search_var.get().strip().lower()
        for widget in self.pass_inner.winfo_children():
            widget.destroy()
        self.pass_vars.clear()
        if not self.all_passwords:
            tb.Label(self.pass_inner, text="No passwords for this company", foreground="gray").pack(pady=10)
            return
        filtered = [p for p in self.all_passwords if search_term in p.lower()]
        if not filtered:
            tb.Label(self.pass_inner, text="No passwords match", foreground="gray").pack(pady=10)
            return
        for pwd in filtered:
            var = tk.IntVar(value=0)
            self.pass_vars[pwd] = var
            row = tb.Frame(self.pass_inner)
            row.pack(fill=X, pady=2)
            tb.Checkbutton(row, variable=var, bootstyle="primary").pack(side=LEFT, padx=5)
            tb.Label(row, text=pwd, anchor="w").pack(side=LEFT, fill=X, expand=True)

    def select_all_passwords(self):
        for var in self.pass_vars.values():
            var.set(1)

    def clear_all_passwords(self):
        for var in self.pass_vars.values():
            var.set(0)

    def get_selected_passwords(self):
        return [pwd for pwd, var in self.pass_vars.items() if var.get() == 1]

    # ---------- Password handling for Extract tab ----------
    def ext_update_password_options(self, event=None):
        selected = self.ext_code_var.get()
        self.ext_all_passwords = [d['Password'] for d in self.company_data if d['Code'] == selected]
        self.ext_filter_password_list()

    def ext_filter_password_list(self):
        search_term = self.ext_pass_search_var.get().strip().lower()
        for widget in self.ext_pass_inner.winfo_children():
            widget.destroy()
        self.ext_pass_vars.clear()
        if not self.ext_all_passwords:
            tb.Label(self.ext_pass_inner, text="No passwords for this company", foreground="gray").pack(pady=10)
            return
        filtered = [p for p in self.ext_all_passwords if search_term in p.lower()]
        if not filtered:
            tb.Label(self.ext_pass_inner, text="No passwords match", foreground="gray").pack(pady=10)
            return
        for pwd in filtered:
            var = tk.IntVar(value=0)
            self.ext_pass_vars[pwd] = var
            row = tb.Frame(self.ext_pass_inner)
            row.pack(fill=X, pady=2)
            tb.Checkbutton(row, variable=var, bootstyle="primary").pack(side=LEFT, padx=5)
            tb.Label(row, text=pwd, anchor="w").pack(side=LEFT, fill=X, expand=True)

    def ext_select_all_passwords(self):
        for var in self.ext_pass_vars.values():
            var.set(1)

    def ext_clear_all_passwords(self):
        for var in self.ext_pass_vars.values():
            var.set(0)

    def ext_get_selected_passwords(self):
        return [pwd for pwd, var in self.ext_pass_vars.items() if var.get() == 1]

    def ext_get_date(self):
        try:
            d = self.ext_date_entry.get_date()
            if isinstance(d, datetime):
                d = d.date()
            return d.strftime("%m%d%Y")
        except AttributeError:
            raw = self.ext_date_entry.get()
            try:
                parsed = datetime.strptime(raw, "%Y-%m-%d")
                return parsed.strftime("%m%d%Y")
            except:
                return date.today().strftime("%m%d%Y")

    # ---------- Dialogs ----------
    def add_new_code(self):
        dialog = tb.Toplevel(self.root)
        dialog.title("New Company Code")
        dialog.geometry("300x200")
        dialog.attributes('-topmost', True)
        tb.Label(dialog, text="Company Code:").pack(pady=5)
        code_entry = tb.Entry(dialog)
        code_entry.pack(pady=5)
        tb.Label(dialog, text="Default Password:").pack(pady=5)
        pwd_entry = tb.Entry(dialog)
        pwd_entry.pack(pady=5)
        def save():
            code = code_entry.get().strip()
            pwd = pwd_entry.get().strip()
            if code and pwd:
                new_row = {'Code': code, 'Password': pwd}
                self.company_data.append(new_row)
                self.append_csv(new_row)
                # Update company dropdown values
                new_codes = sorted(set([d['Code'] for d in self.company_data]))
                self.update_searchable_combobox_values(self.searchable_company_scan, new_codes)
                self.update_searchable_combobox_values(self.searchable_company_extract, new_codes)
                # Set the new code as current
                self.current_company_code.set(code)
                dialog.destroy()
            else:
                messagebox.showwarning("Warning", "Both fields required")
        tb.Button(dialog, text="Save", command=save, bootstyle="success").pack(pady=10)
        dialog.transient(self.root)
        dialog.grab_set()
        self.root.wait_window(dialog)

    def add_new_pass(self):
        code = self.current_company_code.get()
        if not code:
            messagebox.showwarning("Warning", "Select a company code first.")
            return
        dialog = tb.Toplevel(self.root)
        dialog.title("New Password")
        dialog.geometry("300x150")
        dialog.attributes('-topmost', True)
        tb.Label(dialog, text=f"New password for {code}:").pack(pady=5)
        pwd_entry = tb.Entry(dialog)
        pwd_entry.pack(pady=5)
        def save():
            pwd = pwd_entry.get().strip()
            if pwd:
                new_row = {'Code': code, 'Password': pwd}
                self.company_data.append(new_row)
                self.append_csv(new_row)
                self.update_password_options()
                self.ext_update_password_options()
                dialog.destroy()
        tb.Button(dialog, text="Save", command=save, bootstyle="success").pack(pady=10)
        dialog.transient(self.root)
        dialog.grab_set()
        self.root.wait_window(dialog)

    # ---------- Scanner actions ----------
    def queue_for_watcher(self):
        files = self.get_selected_files()
        if not files:
            messagebox.showwarning("Warning", "No files selected.")
            return
        code = self.code_var.get()
        passwords = self.get_selected_passwords()
        if not code or not passwords:
            messagebox.showerror("Error", "Select company and at least one password")
            return
        email_received_date = self.get_email_date()
        config = {
            "processed_at": datetime.now().isoformat(),
            "user": os.getlogin(),
            "email_received_date": email_received_date,
            "company_code": code,
            "passwords": passwords,
            "files": files
        }
        json_path = os.path.join(self.folder_path, f"configuration_{int(time.time())}.json")
        with open(json_path, 'w') as f:
            json.dump(config, f, indent=4)
        self.log(f"Queued {len(files)} files to watcher: {json_path}")
        messagebox.showinfo("Queued", f"Configuration written: {json_path}\nThe watcher will process it automatically.")
        self.refresh_file_list()

    def run_scanner_now(self):
        files = self.get_selected_files()
        if not files:
            messagebox.showwarning("Warning", "No files selected.")
            return
        code = self.code_var.get()
        passwords = self.get_selected_passwords()
        if not code or not passwords:
            messagebox.showerror("Error", "Select company and at least one password")
            return
        email_received_date = self.get_email_date()

        self.stop_btn.config(state=NORMAL)
        self.stop_requested = False
        total_files = len(files)

        self.progress.configure(mode="determinate", maximum=total_files, value=0)
        self.progress.pack(side=BOTTOM, fill=X, padx=10, pady=5)
        self.status_var.set("Scanning...")

        def stop_flag():
            return self.stop_requested

        def progress_update(current, total):
            self.root.after(0, lambda: self._update_progress(current, total))

        def scan():
            try:
                job = ScannerJob(self.folder_path, code, passwords, email_received_date, files)
                def log_callback(msg):
                    self.root.after(0, lambda: self.log(msg))
                process_files_direct(job, progress_callback=log_callback,
                                     stop_flag=stop_flag, progress_update=progress_update)
                self.root.after(0, lambda: self.log("Scanning completed!"))
                self.root.after(0, lambda: self._ask_open_output_folder(job.output_root))
            except Exception as e:
                self.root.after(0, lambda: self.log(f"Error: {e}"))
                self.root.after(0, lambda: messagebox.showerror("Scanner Error", str(e)))
            finally:
                self.root.after(0, self.scan_finished)

        self.scan_thread = threading.Thread(target=scan, daemon=True)
        self.scan_thread.start()

    def _update_progress(self, current, total):
        self.progress['value'] = current
        self.status_var.set(f"Processing file {current} of {total}")

    def _ask_open_output_folder(self, output_root):
        if messagebox.askyesno("Done", f"Output saved to:\n{output_root}\n\nOpen folder?"):
            os.startfile(output_root)

    def stop_scan(self):
        self.stop_requested = True
        self.log("Stop requested – will terminate after current file.")
        self.stop_btn.config(state=DISABLED)

    def scan_finished(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.stop_btn.config(state=DISABLED)
        self.status_var.set("Ready")

    # ---------- Header Extraction action ----------
    def run_header_extraction_now(self):
        files = self.ext_get_selected_files()
        if not files:
            messagebox.showwarning("Warning", "No files selected.")
            return
        code = self.ext_code_var.get().strip()
        passwords = self.ext_get_selected_passwords()
        if not code or not passwords:
            messagebox.showerror("Error", "Select company and at least one password")
            return

        date_str = self.ext_get_date()
        job = HeaderExtractionJob(
            input_folder=self.folder_path,
            company_code=code,
            passwords=passwords,
            date_str=date_str,
            file_names=files
        )

        total_files = len(files)
        self.progress.configure(mode="determinate", maximum=total_files, value=0)
        self.progress.pack(side=BOTTOM, fill=X, padx=10, pady=5)
        self.status_var.set("Extracting headers...")
        self.ext_log("Header extraction started...")

        def progress_update(current, total):
            self.root.after(0, lambda: self._update_progress(current, total))

        def task():
            try:
                def log_callback(msg):
                    self.root.after(0, lambda: self.ext_log(msg))
                process_header_extraction(job, progress_callback=log_callback,
                                          progress_update=progress_update)
                self.root.after(0, lambda: self.ext_log(f"Extraction complete. Output: {job.header_output_file}"))
                self.root.after(0, lambda: self._ask_open_output_folder(os.path.dirname(job.header_output_file)))
            except Exception as e:
                self.root.after(0, lambda: self.ext_log(f"Error: {e}"))
                self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally:
                self.root.after(0, self.header_extraction_finished)

        threading.Thread(target=task, daemon=True).start()

    def header_extraction_finished(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.status_var.set("Ready")
        self.ext_log("Header extraction finished.")

    # ---------- Logging ----------
    def log(self, message, level="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)

    def ext_log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.ext_log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.ext_log_text.see(tk.END)

    def get_email_date(self):
        try:
            d = self.date_entry.get_date()
            if isinstance(d, datetime):
                d = d.date()
            return d.strftime("%Y-%m-%d")
        except AttributeError:
            raw = self.date_entry.get()
            try:
                parsed = datetime.strptime(raw, "%Y-%m-%d")
                return parsed.strftime("%Y-%m-%d")
            except:
                return date.today().strftime("%Y-%m-%d")

    # ---------- Auto-refresh loop (conditional) ----------
    def _auto_refresh_loop(self):
        """If auto-refresh is enabled, refresh file lists every 2 seconds."""
        if self.auto_refresh_var.get():
            self.refresh_file_list()
            self.refresh_extract_file_list()
        self.root.after(2000, self._auto_refresh_loop)

    # ---------- History ----------
    def load_history(self):
        output_root = BASE_OUTPUT_FOLDER
        if not os.path.exists(output_root):
            self.log("Output folder not found.")
            return
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        for folder in os.listdir(output_root):
            folder_path = os.path.join(output_root, folder)
            if os.path.isdir(folder_path):
                log_file = os.path.join(folder_path, f"Log_{folder}.csv")
                if os.path.exists(log_file):
                    try:
                        import polars as pl
                        df = pl.read_csv(log_file)
                        company = df['Company Code'].unique().to_list()[0] if 'Company Code' in df.columns else "Unknown"
                        file_count = df['File Name'].n_unique() if 'File Name' in df.columns else 0
                        self.history_tree.insert("", tk.END, values=(folder, company, file_count, folder_path))
                    except:
                        pass

    # ---------- Settings actions ----------
    def change_theme(self):
        theme = self.theme_var.get()
        try:
            self.root.style.theme_use(theme)
            self.log(f"Theme changed to {theme}")
        except Exception as e:
            messagebox.showerror("Theme Error", str(e))

    def change_watch_folder(self):
        new_path = filedialog.askdirectory()
        if new_path:
            self.watch_path_var.set(new_path)

    def change_csv_file(self):
        new_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if new_path:
            self.csv_path_var.set(new_path)

    def save_settings(self):
        settings = load_settings()
        settings["folder"] = self.watch_path_var.get()
        settings["csv"] = self.csv_path_var.get()
        settings["theme"] = self.theme_var.get()
        save_settings(settings)
        set_user_env_var(ENV_VAR_FOLDER, self.watch_path_var.get())
        set_user_env_var(ENV_VAR_CSV, self.csv_path_var.get())
        self.folder_path = self.watch_path_var.get()
        self.csv_path = self.csv_path_var.get()
        self.company_data = self.load_csv()
        new_codes = sorted(set([d['Code'] for d in self.company_data]))
        self.update_searchable_combobox_values(self.searchable_company_scan, new_codes)
        self.update_searchable_combobox_values(self.searchable_company_extract, new_codes)
        messagebox.showinfo("Settings", "Settings saved. Restart the application for full effect.")


# ==================== MAIN ====================
if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == "--watch":
            start_watching()
        elif sys.argv[1] == "--process":
            instance_check_or_focus(SCANNER_WINDOW_TITLE)
            root = tb.Window()
            try:
                root.style.theme_use("flatly")
            except:
                pass
            app = EnhancedScannerGUI(root)
            root.mainloop()
        elif sys.argv[1] == "--process-config" and len(sys.argv) == 3:
            process_config_and_scan(sys.argv[2])
        else:
            print("Usage:")
            print("  python ofac_gui.py              : Launch setup/GUI")
            print("  python ofac_gui.py --watch      : Start folder watcher")
            print("  python ofac_gui.py --process-config <json_path> : Run scanner from config")
    else:
        instance_check_or_focus(SETUP_WINDOW_TITLE)
        root = tb.Window()
        try:
            root.style.theme_use("flatly")
        except:
            pass
        saved = load_settings()
        folder = saved.get("folder")
        csv_path = saved.get("csv")
        theme = saved.get("theme", "flatly")
        if folder and csv_path:
            try:
                root.style.theme_use(theme)
            except:
                pass
            EnhancedScannerGUI(root, folder_path=folder, csv_path=csv_path)
        else:
            SetupGUI(root)
        root.mainloop()
