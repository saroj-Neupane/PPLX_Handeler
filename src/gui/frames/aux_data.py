"""Aux Data editing frame with Excel integration."""

import os
from typing import Dict, List

import tkinter as tk
from tkinter import ttk, filedialog

from src.core.logic import analyze_mr_note_for_aux_data
from src.core.utils import parse_keywords
from src.excel.loader import load_excel_data
from src.gui.constants import AUX_AUTO_FILL_CONFIG, PLACEHOLDER_AUX1, PLACEHOLDER_AUX2, THEME

try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


def _get_auto_fill_enabled(aux_frame, field_index: int, default: bool = False) -> bool:
    """Return True if aux field is set to auto-fill from Excel."""
    try:
        var = getattr(aux_frame, f"auto_fill_aux{field_index + 1}_var", None)
        return bool(var.get()) if var else default
    except Exception:
        return default


class AuxDataEditFrame(ttk.Frame):
    """Frame for editing Aux Data fields."""

    def __init__(self, parent, config_manager):
        super().__init__(parent)
        self.config_manager = config_manager
        self.aux_entries = []
        self.ignore_scid_keywords = self.config_manager.get("ignore_scid_keywords", "")

    def setup_ui(self, toolbar_parent, excel_parent, keywords_parent, fields_parent):
        """Setup the Aux Data editing UI. Parents must be different to avoid pack/grid conflict."""
        self.excel_frame = ttk.LabelFrame(excel_parent, text="Node-Section-Connection File", padding=8)
        self.excel_frame.columnconfigure(0, weight=1)
        self.excel_file_var = tk.StringVar()
        self.excel_file_var.set(self.config_manager.get("excel_file_path", "No file selected"))
        ttk.Label(self.excel_frame, textvariable=self.excel_file_var, foreground=THEME["purple"]).grid(
            row=0, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Button(self.excel_frame, text="Browse", command=self.select_excel_file).grid(
            row=0, column=1, padx=(4, 0)
        )

        self.aux_fields = [
            ("Aux Data 1", "Pole Owner", True),
            ("Aux Data 2", "Pole Tag", True),
            ("Aux Data 3", "Condition", True),
            ("Aux Data 4", "Make Ready Type", False),
            ("Aux Data 5", "Proposed Riser", False),
        ]

        self.fields_frame = ttk.LabelFrame(fields_parent, text="Aux Data", padding=8)
        self.fields_frame.columnconfigure(1, weight=1, minsize=120)
        for i, (field_name, description, is_editable) in enumerate(self.aux_fields):
            row_frame = ttk.Frame(self.fields_frame)
            row_frame.grid(row=i, column=0, columnspan=2, sticky="ew", pady=3)
            row_frame.columnconfigure(0, minsize=220)
            row_frame.columnconfigure(1, weight=1, minsize=100)
            row_frame.columnconfigure(2, minsize=80)

            label_text = f"{field_name} ({description}):"
            lbl = ttk.Label(row_frame, text=label_text, anchor="w")
            lbl.grid(row=0, column=0, sticky="w", padx=(0, 8))

            if is_editable:
                entry = ttk.Entry(row_frame, width=25)
                entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))
                if i == 2:
                    entry.insert(0, "Auto (EXISTING/PROPOSED)")
                    entry.config(state="readonly")

                for idx, config_key, cb_label, placeholder, default_checked in AUX_AUTO_FILL_CONFIG:
                    if i == idx:
                        var = tk.BooleanVar()
                        setattr(self, f"auto_fill_aux{idx + 1}_var", var)
                        ttk.Checkbutton(
                            row_frame,
                            text=cb_label,
                            variable=var,
                            command=lambda i=idx: self._toggle_aux_auto_fill(i),
                        ).grid(row=0, column=2, sticky="w")
                        checkbox_state = self.config_manager.get(config_key, default_checked)
                        var.set(checkbox_state)
                        if checkbox_state:
                            entry.config(state="readonly")
                            entry.delete(0, tk.END)
                            entry.insert(0, placeholder)
                        break

                self.aux_entries.append(entry)
            else:
                auto_label = ttk.Label(
                    row_frame, text="(Auto Filled)", foreground=THEME["text_muted"]
                )
                auto_label.grid(row=0, column=1, sticky="w", padx=(0, 0))
                self.aux_entries.append(auto_label)

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if file_path:
            self.excel_file_var.set(file_path)
            self.config_manager.set("excel_file_path", file_path)

    def get_excel_path(self) -> str:
        value = self.excel_file_var.get()
        if value and value != "No file selected":
            return value
        return self.config_manager.get("excel_file_path", "")

    def _toggle_aux_auto_fill(self, field_index: int):
        config_key, _, placeholder, *_ = next(
            c for c in AUX_AUTO_FILL_CONFIG if c[0] == field_index
        )
        var = getattr(self, f"auto_fill_aux{field_index + 1}_var")
        state = var.get()
        self.config_manager.set(config_key, state)
        entry = self.aux_entries[field_index]
        if state:
            entry.config(state="normal")
            entry.delete(0, tk.END)
            entry.insert(0, placeholder)
            entry.config(state="readonly")
        else:
            entry.config(state="normal")

    def analyze_mr_note(self, mr_note: str) -> tuple:
        return analyze_mr_note_for_aux_data(
            mr_note,
            comm_keywords=parse_keywords(self.config_manager.get("comm_keywords", "")),
            power_keywords=parse_keywords(self.config_manager.get("power_keywords", "")),
            pco_keywords=parse_keywords(self.config_manager.get("pco_keywords", "")),
            aux5_keywords=parse_keywords(self.config_manager.get("aux5_keywords", "")),
            power_label=self.config_manager.get("power_label", "POWER"),
        )

    def get_aux_values(self) -> Dict[int, str]:
        values = {}
        placeholders = (PLACEHOLDER_AUX1, PLACEHOLDER_AUX2)
        entry_index = 0
        for i, (_, _, is_editable) in enumerate(self.aux_fields):
            if not is_editable:
                entry_index += 1
                continue
            if i in {c[0] for c in AUX_AUTO_FILL_CONFIG} and _get_auto_fill_enabled(
                self, i
            ):
                entry_index += 1
                continue
            entry = self.aux_entries[entry_index]
            entry_index += 1
            if hasattr(entry, "get"):
                value = entry.get().strip()
                if value and value not in placeholders:
                    values[i + 1] = value
        return values

    def set_readonly_field(self, field_index: int, value: str):
        if 0 <= field_index < len(self.aux_entries):
            entry = self.aux_entries[field_index]
            was_readonly = str(entry.cget("state")) == "readonly"
            if was_readonly:
                entry.config(state="normal")
            entry.delete(0, tk.END)
            entry.insert(0, value)
            if was_readonly:
                entry.config(state="readonly")

    def load_excel_data(self, log_callback=None) -> Dict[str, Dict]:
        excel_path = self.config_manager.get("excel_file_path", "")
        return load_excel_data(excel_path, log_callback=log_callback)

    def get_valid_scids(self) -> set:
        excel_data = self.load_excel_data()
        return set(excel_data.keys())
