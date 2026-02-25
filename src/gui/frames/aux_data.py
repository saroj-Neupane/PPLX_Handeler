"""Aux Data editing frame with Excel integration."""

import os
from typing import Dict, List

import tkinter as tk
from tkinter import ttk, filedialog

from src.config.manager import get_available_configs
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
        self.toolbar_frame = ttk.Frame(toolbar_parent)
        self.toolbar_frame.columnconfigure(1, weight=1)
        ttk.Label(self.toolbar_frame, text="Config:").grid(row=0, column=0, padx=(0, 4), sticky="w")
        self.config_var = tk.StringVar()
        self.config_combo = ttk.Combobox(
            self.toolbar_frame, textvariable=self.config_var, state="readonly", width=16
        )
        self._refresh_config_combo()
        saved = self.config_manager.config_name
        names = list(self.config_combo["values"])
        self.config_var.set(saved if saved in names else (names[0] if names else "OPPD"))
        self.config_combo.grid(row=0, column=1, padx=(0, 12), sticky="w")
        self.config_combo.bind("<<ComboboxSelected>>", self._on_config_selected)

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

        self.keywords_frame = ttk.LabelFrame(keywords_parent, text="Keywords", padding=8)
        self.keywords_frame.columnconfigure(1, weight=1)
        self.comm_owners_var = tk.StringVar(value=self.config_manager.get("comm_owners", ""))
        self.power_owners_var = tk.StringVar(value=self.config_manager.get("power_owners", ""))
        self.pco_keywords_var = tk.StringVar(value=self.config_manager.get("pco_keywords", ""))
        self.aux5_keywords_var = tk.StringVar(value=self.config_manager.get("aux5_keywords", ""))

        for row, (label_text, var) in enumerate(
            [
                ("Comm:", self.comm_owners_var),
                ("Power:", self.power_owners_var),
                ("PCO:", self.pco_keywords_var),
                ("Riser:", self.aux5_keywords_var),
            ]
        ):
            ttk.Label(self.keywords_frame, text=label_text, width=8, anchor="w").grid(
                row=row, column=0, sticky="w", pady=2, padx=(0, 4)
            )
            entry = ttk.Entry(self.keywords_frame, textvariable=var)
            entry.grid(row=row, column=1, sticky="ew", pady=2)
            entry.bind("<FocusOut>", self.save_owner_config)

        self.aux_fields = [
            ("Aux Data 1", "Pole Owner", True),
            ("Aux Data 2", "Pole Tag", True),
            ("Aux Data 3", "Condition", True),
            ("Aux Data 4", "Make Ready Type", False),
            ("Aux Data 5", "Proposed Riser", False),
        ]

        self.fields_frame = ttk.LabelFrame(fields_parent, text="Aux Data", padding=8)
        self.fields_frame.columnconfigure(1, weight=1, minsize=120)
        saved_values = self.config_manager.get("last_aux_values", {})

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
                saved_key = f"aux_data_{i + 1}"
                if i == 2:
                    entry.insert(0, "Auto (EXISTING/PROPOSED)")
                    entry.config(state="readonly")
                elif saved_key in saved_values:
                    entry.insert(0, saved_values[saved_key])

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

                if i != 2:
                    entry.bind("<FocusOut>", self.auto_save_values)
                    entry.bind("<KeyRelease>", self.auto_save_values)
                self.aux_entries.append(entry)
            else:
                auto_label = ttk.Label(
                    row_frame, text="(Auto Filled)", foreground=THEME["text_muted"]
                )
                auto_label.grid(row=0, column=1, sticky="w", padx=(0, 0))
                self.aux_entries.append(auto_label)

    def _refresh_config_combo(self):
        """Load config names from config/ folder."""
        self.config_combo["values"] = get_available_configs()

    def _on_config_selected(self, event=None):
        val = self.config_var.get()
        if val:
            self.config_manager.switch_config(val)

    def get_selected_config_power_label(self) -> str:
        return self.config_manager.get_power_label()

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

    def auto_save_values(self, event=None):
        placeholders = (PLACEHOLDER_AUX1, PLACEHOLDER_AUX2, "Auto (EXISTING/PROPOSED)")
        values = {}
        entry_index = 0
        for i, (_, _, is_editable) in enumerate(self.aux_fields):
            if not is_editable:
                entry_index += 1
                continue
            entry = self.aux_entries[entry_index]
            entry_index += 1
            if hasattr(entry, "get"):
                value = entry.get().strip()
                if value and value not in placeholders:
                    values[f"aux_data_{i + 1}"] = value
        self.config_manager.set("last_aux_values", values)

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

    def save_owner_config(self, event=None):
        self.config_manager.set("comm_owners", self.comm_owners_var.get())
        self.config_manager.set("power_owners", self.power_owners_var.get())
        self.config_manager.set("pco_keywords", self.pco_keywords_var.get())
        self.config_manager.set("aux5_keywords", self.aux5_keywords_var.get())

    def analyze_mr_note(self, mr_note: str) -> tuple:
        def get_keywords(var_attr: str) -> List[str]:
            var = getattr(self, var_attr, None)
            return parse_keywords(var.get() if var else "")

        return analyze_mr_note_for_aux_data(
            mr_note,
            comm_keywords=get_keywords("comm_owners_var"),
            power_keywords=get_keywords("power_owners_var"),
            pco_keywords=get_keywords("pco_keywords_var"),
            aux5_keywords=get_keywords("aux5_keywords_var"),
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
