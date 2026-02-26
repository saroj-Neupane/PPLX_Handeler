"""File list frame for selecting PPLX files from folder or ZIP."""

import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import List

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from src.gui.constants import THEME


class PPLXFileListFrame(ttk.Frame):
    """Frame for displaying and managing selected PPLX files."""

    def __init__(self, parent, config_manager, category: str):
        super().__init__(parent)
        self.config_manager = config_manager
        self.category = category
        self.files: List[str] = []
        self.current_folder = ""
        self.source_path = ""
        self.source_type = "folder"
        self.display_name = ""
        self.temp_directory = None
        normalized_category = self.category.lower().replace(" ", "_")
        self.config_key = f"last_{normalized_category}_folder_path"

        self.setup_ui()

        last_source = self.config_manager.get(self.config_key, "")
        if last_source:
            try:
                if last_source.lower().endswith(".zip") and os.path.exists(last_source):
                    self.load_zip_source(last_source, remember=False)
                elif os.path.isdir(last_source):
                    self.load_directory_source(last_source, remember=False)
            except Exception as exc:
                print(f"Warning: Could not auto-load previous source: {exc}")

    def setup_ui(self):
        """Setup the file list UI with minimal design."""
        # Section header row
        header_row = ttk.Frame(self)
        header_row.pack(fill="x", pady=(0, 8))

        ttk.Label(
            header_row, text=self.category, style="Heading.TLabel",
        ).pack(side="left")

        self.count_label = ttk.Label(header_row, text="", style="Muted.TLabel")
        self.count_label.pack(side="right")

        # Select button
        ttk.Button(
            self, text=f"Select Source", command=self.select_folder,
        ).pack(fill="x", pady=(0, 8))

        # File list
        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True)

        self.file_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            height=6,
            bg=THEME["bg_card"],
            fg=THEME["text"],
            selectbackground=THEME["purple"],
            selectforeground="white",
            font=("Segoe UI", 9),
            highlightthickness=1,
            highlightcolor=THEME["border"],
            highlightbackground=THEME["border"],
            relief="flat",
            bd=0,
        )
        scrollbar = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.file_listbox.yview,
        )
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        self.file_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Source path
        self.folder_label = ttk.Label(self, text="", style="Muted.TLabel")
        self.folder_label.pack(anchor="w", pady=(6, 0))

    def select_folder(self):
        """Select folder or ZIP containing PPLX files."""
        initial_dir = self.config_manager.get(self.config_key, "")
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = os.getcwd()

        zip_source = filedialog.askopenfilename(
            title=f"Select {self.category} ZIP Archive (Cancel to pick a folder)",
            initialdir=initial_dir,
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")],
        )

        if zip_source:
            source_path = Path(zip_source)
            if source_path.is_file() and source_path.suffix.lower() == ".zip":
                self.load_zip_source(str(source_path))
            elif source_path.is_dir():
                self.load_directory_source(str(source_path))
            else:
                messagebox.showerror(
                    "Invalid Selection",
                    "Please select a folder or a .zip archive containing PPLX files.",
                )
            return

        folder = filedialog.askdirectory(
            title=f"Select {self.category} Folder with PPLX Files",
            initialdir=initial_dir,
            mustexist=True,
        )
        if folder:
            self.load_directory_source(folder)

    def load_folder_files(self):
        """Load all PPLX files from the current folder."""
        if not getattr(self, "current_folder", None) or not self.current_folder:
            return
        if not os.path.exists(self.current_folder):
            messagebox.showerror("Error", f"Source path not found: {self.current_folder}")
            return

        self.files.clear()
        try:
            pplx_files = []
            for root, _, filenames in os.walk(self.current_folder):
                for filename in filenames:
                    if filename.lower().endswith(".pplx"):
                        pplx_files.append(os.path.join(root, filename))
            self.files = sorted(pplx_files)

            source_label = self.display_name or os.path.basename(self.current_folder)
            if self.source_type == "zip":
                source_label = f"{source_label} [ZIP]"
            self.folder_label.config(text=source_label)
            self.update_display()
        except Exception as e:
            messagebox.showerror("Error", f"Error reading folder: {str(e)}")

    def clear_files(self):
        """Clear all files from the list."""
        self.cleanup_temp_dir()
        self.files.clear()
        self.current_folder = ""
        self.source_path = ""
        self.display_name = ""
        self.source_type = "folder"
        self.folder_label.config(text="")
        self.update_display()

    def update_display(self):
        """Update the file list display."""
        self.file_listbox.delete(0, tk.END)
        for file_path in self.files:
            try:
                display_name = os.path.relpath(file_path, self.current_folder)
            except ValueError:
                display_name = os.path.basename(file_path)
            self.file_listbox.insert(tk.END, display_name.replace("\\", "/"))

        count = len(self.files)
        if count > 0:
            self.count_label.config(text=f"{count} file{'s' if count != 1 else ''}")
        else:
            self.count_label.config(text="")

    def get_files(self) -> List[str]:
        """Get the list of selected files."""
        return self.files.copy()

    def get_working_directory(self) -> str:
        """Return the working directory."""
        return self.current_folder

    def get_source_path(self) -> str:
        """Return the path originally selected by the user."""
        return self.source_path or self.current_folder

    def get_current_folder(self) -> str:
        """Backward-compatible alias for working directory."""
        return self.current_folder

    def cleanup_temp_dir(self):
        """Remove any temporary extraction directory."""
        if self.temp_directory and os.path.exists(self.temp_directory):
            shutil.rmtree(self.temp_directory, ignore_errors=True)
        self.temp_directory = None

    def destroy(self):
        """Ensure temp data is cleaned up when frame is destroyed."""
        self.cleanup_temp_dir()
        super().destroy()

    def load_directory_source(self, folder_path: str, remember: bool = True):
        """Configure frame to use a regular folder."""
        if not os.path.isdir(folder_path):
            messagebox.showerror("Error", f"Folder not found: {folder_path}")
            return
        self.cleanup_temp_dir()
        self.source_type = "folder"
        self.source_path = folder_path
        self.display_name = os.path.basename(os.path.normpath(folder_path)) or folder_path
        self.current_folder = folder_path
        if remember:
            self.config_manager.set(self.config_key, folder_path)
        self.load_folder_files()

    def load_zip_source(self, zip_path: str, remember: bool = True):
        """Configure frame to use a ZIP archive."""
        if not os.path.exists(zip_path):
            messagebox.showerror("Error", f"ZIP file not found: {zip_path}")
            return
        self.cleanup_temp_dir()
        temp_dir = ""
        try:
            temp_dir = tempfile.mkdtemp(prefix=f"pplx_{self.category.lower()}_")
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(temp_dir)
        except Exception as exc:
            if temp_dir:
                shutil.rmtree(temp_dir, ignore_errors=True)
            messagebox.showerror("Error", f"Failed to extract ZIP:\n{exc}")
            return
        self.temp_directory = temp_dir
        self.source_type = "zip"
        self.source_path = zip_path
        self.display_name = os.path.basename(zip_path)
        self.current_folder = temp_dir
        if remember:
            self.config_manager.set(self.config_key, zip_path)
        self.load_folder_files()
