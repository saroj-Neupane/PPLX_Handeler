"""Main PPLX GUI Application."""

import os
import sys

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from src.config.manager import PPLXConfigManager
from src.core.handler import PPLXHandler
from src.gui.constants import THEME
from src.gui.frames.file_list import PPLXFileListFrame
from src.gui.frames.aux_data import AuxDataEditFrame
from src.gui.frames.processing import ProcessingFrame

try:
    from PIL import Image, ImageTk
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False


def _apply_theme(style: ttk.Style) -> None:
    """Apply White & Purple theme to ttk widgets."""
    bg, panel, card = THEME["bg_dark"], THEME["bg_panel"], THEME["bg_card"]
    purple, text, text_muted = (
        THEME["purple"],
        THEME["text"],
        THEME["text_muted"],
    )
    style.configure(".", background=panel, foreground=text, font=("Segoe UI", 10))
    style.configure("TFrame", background=panel)
    style.configure("TLabel", background=panel, foreground=text)
    style.configure("TLabelframe", background=panel, foreground=purple)
    style.configure("TLabelframe.Label", background=panel, foreground=purple, font=("Segoe UI", 10, "bold"))
    style.configure("TButton", background=card, foreground=purple, padding=(12, 6))
    style.map("TButton", background=[("active", purple)], foreground=[("active", "white")])
    style.configure("TEntry", fieldbackground=card, foreground=text, insertcolor=purple)
    style.configure("TCombobox", fieldbackground=card, foreground=text, background=card)
    style.map("TCombobox", fieldbackground=[("readonly", card)], background=[("readonly", card)])
    style.configure("Horizontal.TProgressbar", background=purple, troughcolor=card, thickness=8)
    style.configure("TPanedwindow", background=panel)
    style.configure("Vertical.TScrollbar", troughcolor=card, background=panel)
    style.configure("TCheckbutton", background=panel, foreground=text)
    style.map("TCheckbutton", background=[("active", panel)], foreground=[("active", purple)])


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
        self.root.geometry(self.config_manager.get("window_geometry", "1000x700"))
        self.root.minsize(800, 600)

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

        self.root.configure(bg=THEME["bg_dark"])
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="12")
        main_frame.pack(fill="both", expand=True)

        header = tk.Frame(main_frame, bg=THEME["bg_dark"], height=56)
        header.pack(fill="x", pady=(0, 14))
        header.pack_propagate(False)

        icon_path = _get_icon_path()
        if icon_path and PILLOW_AVAILABLE:
            try:
                from PIL import Image
                img = Image.open(icon_path)
                img = img.resize((40, 40), Image.Resampling.LANCZOS)
                self._header_icon = ImageTk.PhotoImage(img)
                icon_label = tk.Label(
                    header, image=self._header_icon, bg=THEME["bg_dark"]
                )
                icon_label.pack(side="left", padx=(16, 12), pady=8)
            except Exception:
                pass

        title_label = tk.Label(
            header,
            text="PPLX Handler",
            font=("Segoe UI", 18, "bold"),
            fg=THEME["purple"],
            bg=THEME["bg_dark"],
        )
        title_label.pack(side="left", pady=14)

        sep = tk.Frame(main_frame, height=2, bg=THEME["purple"])
        sep.pack(fill="x", pady=(0, 10))

        self.aux_frame = AuxDataEditFrame(main_frame, self.config_manager)

        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)

        left_col = ttk.Frame(paned)
        right_col = ttk.Frame(paned)
        paned.add(left_col, weight=1)
        paned.add(right_col, weight=1)

        self.aux_frame.setup_ui(
            toolbar_parent=main_frame,
            excel_parent=left_col,
            keywords_parent=left_col,
            fields_parent=left_col,
        )

        self.aux_frame.toolbar_frame.pack(fill="x", pady=(0, 10))
        paned.pack(fill="both", expand=True)

        sources_frame = ttk.LabelFrame(left_col, text="PPLX Sources", padding=8)
        sources_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        sources_frame.grid_rowconfigure(0, weight=1)
        sources_frame.columnconfigure(0, weight=1)
        sources_frame.columnconfigure(1, weight=1)
        sources_frame.rowconfigure(0, weight=1)

        self.existing_frame = PPLXFileListFrame(
            sources_frame, self.config_manager, category="EXISTING"
        )
        self.existing_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.proposed_frame = PPLXFileListFrame(
            sources_frame, self.config_manager, category="PROPOSED"
        )
        self.proposed_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        self.aux_frame.excel_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        self.aux_frame.keywords_frame.grid(row=2, column=0, sticky="ew", pady=(0, 8))

        self.aux_frame.fields_frame.grid(row=3, column=0, sticky="ew", pady=(0, 0))

        left_col.columnconfigure(0, weight=1)
        left_col.rowconfigure(0, weight=1)

        self.processing_frame = ProcessingFrame(
            right_col,
            self.config_manager,
            self.existing_frame,
            self.proposed_frame,
            self.aux_frame,
        )
        self.processing_frame.pack(fill="both", expand=True)

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
        self.config_manager.set("window_geometry", self.root.geometry())
        if hasattr(self, "existing_frame"):
            self.existing_frame.cleanup_temp_dir()
        if hasattr(self, "proposed_frame"):
            self.proposed_frame.cleanup_temp_dir()
        self.root.destroy()

    def run(self):
        self.root.mainloop()
