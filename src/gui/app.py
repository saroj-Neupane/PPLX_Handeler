"""Main PPLX GUI Application."""

import os
import sys

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from src.config.manager import PPLXConfigManager, get_available_configs
from src.core.handler import PPLXHandler
from src.gui.constants import THEME
from src.gui.frames.file_list import PPLXFileListFrame
from src.gui.frames.processing import ProcessingFrame

try:
    from PIL import Image, ImageTk
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False


def _apply_theme(style: ttk.Style) -> None:
    """Apply modern minimal theme to ttk widgets."""
    bg = THEME["bg"]
    panel = THEME["bg_panel"]
    card = THEME["bg_card"]
    purple = THEME["purple"]
    purple_light = THEME["purple_light"]
    purple_dim = THEME["purple_dim"]
    text = THEME["text"]
    text_muted = THEME["text_muted"]
    border = THEME["border"]

    style.configure(".", background=panel, foreground=text, font=("Segoe UI", 10))
    style.configure("TFrame", background=panel)
    style.configure("TLabel", background=panel, foreground=text)

    style.configure("TLabelframe", background=panel, foreground=text, bordercolor=border)
    style.configure(
        "TLabelframe.Label", background=panel, foreground=THEME["text_secondary"],
        font=("Segoe UI", 9, "bold"),
    )

    style.configure("TButton", background=card, foreground=purple, padding=(12, 6))
    style.map("TButton", background=[("active", border)], foreground=[("active", purple)])

    # Accent button - filled purple for primary actions
    style.configure(
        "Accent.TButton",
        background=purple, foreground="white",
        padding=(20, 10), font=("Segoe UI", 11, "bold"),
    )
    style.map(
        "Accent.TButton",
        background=[("active", purple_light), ("disabled", purple_dim)],
        foreground=[("disabled", panel)],
    )

    style.configure("TEntry", fieldbackground=card, foreground=text, insertcolor=purple)
    style.configure("TCombobox", fieldbackground=card, foreground=text, background=card)
    style.map("TCombobox", fieldbackground=[("readonly", card)], background=[("readonly", card)])

    style.configure(
        "Horizontal.TProgressbar",
        background=purple, troughcolor=card, thickness=6,
    )

    style.configure("TPanedwindow", background=panel)
    style.configure("Vertical.TScrollbar", troughcolor=card, background=border)
    style.configure("TCheckbutton", background=panel, foreground=text)
    style.map("TCheckbutton", background=[("active", panel)], foreground=[("active", purple)])

    # Muted label style
    style.configure("Muted.TLabel", foreground=text_muted, font=("Segoe UI", 9))
    style.configure("Secondary.TLabel", foreground=THEME["text_secondary"])
    style.configure("Heading.TLabel", foreground=text, font=("Segoe UI", 11, "bold"))


def _get_icon_path() -> str | None:
    """Return path to handler.ico in assets/ (script and PyInstaller bundle)."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    for subpath in ("assets/handler.ico", "handler.ico"):
        path = os.path.join(base, subpath)
        if os.path.exists(path):
            return path
    return None


class PPLXGUIApp:
    """Main PPLX GUI Application."""

    def __init__(self):
        self.config_manager = PPLXConfigManager()
        self.root = tk.Tk()
        self.setup_window()
        self.setup_ui()

    def setup_window(self):
        self.root.title("PPLX Handler")
        self.root.geometry("780x650")
        self.root.minsize(620, 500)

        icon_path = _get_icon_path()
        if icon_path:
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
        _apply_theme(style)

        self.root.configure(bg=THEME["bg"])
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_ui(self):
        # Outer wrapper centers content with max width
        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        main = ttk.Frame(outer, padding=20)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)

        # Cap content width so it doesn't stretch on wide screens
        def _cap_width(event=None):
            max_w = 900
            if outer.winfo_width() > max_w:
                pad_x = (outer.winfo_width() - max_w) // 2
                main.grid_configure(padx=pad_x)
            else:
                main.grid_configure(padx=0)
        outer.bind("<Configure>", _cap_width)

        # -- Row 0: Header --
        header = tk.Frame(main, bg=THEME["bg_panel"])
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 20))

        icon_path = _get_icon_path()
        if icon_path and PILLOW_AVAILABLE:
            try:
                img = Image.open(icon_path)
                img = img.resize((28, 28), Image.Resampling.LANCZOS)
                self._header_icon = ImageTk.PhotoImage(img)
                tk.Label(
                    header, image=self._header_icon, bg=THEME["bg_panel"],
                ).pack(side="left", padx=(0, 10))
            except Exception:
                pass

        tk.Label(
            header, text="PPLX Handler",
            font=("Segoe UI", 20, "bold"),
            fg=THEME["purple"], bg=THEME["bg_panel"],
        ).pack(side="left")

        # Config profile selector (right side of header)
        configs = get_available_configs()
        self.config_var = tk.StringVar(value=self.config_manager.config_name)
        config_combo = ttk.Combobox(
            header, textvariable=self.config_var,
            values=configs, state="readonly", width=12,
        )
        config_combo.pack(side="right")
        config_combo.bind("<<ComboboxSelected>>", self._on_config_changed)
        tk.Label(
            header, text="Config", font=("Segoe UI", 9),
            fg=THEME["text_secondary"], bg=THEME["bg_panel"],
        ).pack(side="right", padx=(0, 6))

        # -- Row 1: Excel file selection --
        excel_row = ttk.Frame(main)
        excel_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        excel_row.columnconfigure(1, weight=1)

        ttk.Label(
            excel_row, text="Excel File", style="Heading.TLabel",
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))

        self.excel_file_var = tk.StringVar()
        self.excel_file_var.set(self.config_manager.get("excel_file_path", "No file selected"))
        ttk.Label(
            excel_row, textvariable=self.excel_file_var, style="Muted.TLabel",
        ).grid(row=0, column=1, sticky="ew", padx=(0, 8))

        ttk.Button(
            excel_row, text="Browse", command=self.select_excel_file,
        ).grid(row=0, column=2)

        # -- Row 2: Midspan heights Excel selection (optional) --
        heights_row = ttk.Frame(main)
        heights_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        heights_row.columnconfigure(1, weight=1)

        ttk.Label(
            heights_row, text="Midspan Heights File", style="Heading.TLabel",
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))

        self.midspan_file_var = tk.StringVar()
        self.midspan_file_var.set(self.config_manager.get("midspan_heights_file_path", "No file selected"))
        ttk.Label(
            heights_row, textvariable=self.midspan_file_var, style="Muted.TLabel",
        ).grid(row=0, column=1, sticky="ew", padx=(0, 8))

        ttk.Button(
            heights_row, text="Browse", command=self.select_midspan_file,
        ).grid(row=0, column=2)

        # Subtle divider
        tk.Frame(main, height=1, bg=THEME["border"]).grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=(0, 16),
        )

        # -- Row 3: File lists side by side --
        self.existing_frame = PPLXFileListFrame(
            main, self.config_manager, category="EXISTING",
        )
        self.existing_frame.grid(row=4, column=0, sticky="nsew", padx=(0, 8), pady=(0, 12))

        self.proposed_frame = PPLXFileListFrame(
            main, self.config_manager, category="PROPOSED",
        )
        self.proposed_frame.grid(row=4, column=1, sticky="nsew", padx=(8, 0), pady=(0, 12))

        main.rowconfigure(4, weight=1)

        # -- Row 4: Processing --
        self.processing_frame = ProcessingFrame(
            main,
            self.config_manager,
            self.existing_frame,
            self.proposed_frame,
        )
        self.processing_frame.grid(row=5, column=0, columnspan=2, sticky="nsew")
        main.rowconfigure(5, weight=2)

    def _on_config_changed(self, event=None):
        name = self.config_var.get()
        if name and name != self.config_manager.config_name:
            self.config_manager.switch_config(name)

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if file_path:
            self.excel_file_var.set(file_path)
            self.config_manager.set("excel_file_path", file_path)

    def select_midspan_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Node & Midspan Heights Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if file_path:
            self.midspan_file_var.set(file_path)
            self.config_manager.set("midspan_heights_file_path", file_path)

    def show_batch_report(self):
        from datetime import datetime
        files = self.existing_frame.get_files() + self.proposed_frame.get_files()
        if not files:
            messagebox.showwarning("No Files", "Please select PPLX files first")
            return
        report_window = tk.Toplevel(self.root)
        report_window.title("Batch Report")
        report_window.geometry("600x400")
        report_window.configure(bg=THEME["bg_panel"])
        report_text = scrolledtext.ScrolledText(
            report_window,
            wrap=tk.WORD,
            bg=THEME["bg_card"],
            fg=THEME["text"],
            insertbackground=THEME["purple"],
            font=("Consolas", 10),
        )
        report_text.pack(fill="both", expand=True, padx=10, pady=10)
        report_text.insert(tk.END, f"PPLX Files Batch Report\n")
        report_text.insert(
            tk.END,
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n",
        )
        report_text.insert(tk.END, f"Total Files: {len(files)}\n\n")
        for i, file_path in enumerate(files, 1):
            try:
                handler = PPLXHandler(file_path)
                info = handler.get_file_info()
                aux_data = handler.get_aux_data()
                report_text.insert(tk.END, f"{i}. {os.path.basename(file_path)}\n")
                report_text.insert(tk.END, f"   Date: {info.get('date', 'Unknown')}\n")
                report_text.insert(tk.END, f"   User: {info.get('user', 'Unknown')}\n")
                non_default = {k: v for k, v in aux_data.items() if v != "Unset"}
                report_text.insert(
                    tk.END,
                    f"   Aux Data: {non_default if non_default else 'All unset'}\n",
                )
                report_text.insert(tk.END, "\n")
            except Exception as e:
                report_text.insert(tk.END, f"{i}. {os.path.basename(file_path)}\n")
                report_text.insert(tk.END, f"   Error: {str(e)}\n\n")

    def export_structure(self):
        files = self.existing_frame.get_files() + self.proposed_frame.get_files()
        if not files:
            messagebox.showwarning("No Files", "Please select a PPLX file first")
            return
        file_path = files[0]
        output_file = filedialog.asksaveasfilename(
            title="Export Structure to JSON",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if output_file:
            try:
                handler = PPLXHandler(file_path)
                handler.export_structure_to_json(output_file)
                messagebox.showinfo(
                    "Success", f"Structure exported to:\n{output_file}"
                )
            except Exception as e:
                messagebox.showerror(
                    "Error", f"Failed to export structure:\n{str(e)}"
                )

    def on_closing(self):
        if hasattr(self, "existing_frame"):
            self.existing_frame.cleanup_temp_dir()
        if hasattr(self, "proposed_frame"):
            self.proposed_frame.cleanup_temp_dir()
        self.root.destroy()

    def run(self):
        self.root.mainloop()
