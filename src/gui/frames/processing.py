"""Processing frame for batch PPLX file processing."""

import os
import subprocess
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

from src.core.logic import extract_scid_from_filename
from src.core.processor import process_single_file
from src.core.utils import parse_keywords
from src.excel.changelog import write_change_log
from src.excel.loader import load_excel_data
from src.gui.constants import THEME


def _wire_spec_base_path():
    """Base path for OPPD shapefiles."""
    return Path(__file__).resolve().parents[3] / "data" / "OPPD" / "shape"


try:
    from src.core.wire_spec_from_excel import build_wire_spec_comparison
    _WIRE_SPEC_IMPORT_ERROR = None
except ImportError as _e:
    build_wire_spec_comparison = None
    _WIRE_SPEC_IMPORT_ERROR = str(_e)


class ProcessingFrame(ttk.Frame):
    """Frame for processing controls and output."""

    def __init__(self, parent, config_manager, existing_frame, proposed_frame):
        super().__init__(parent)
        self.config_manager = config_manager
        self.existing_frame = existing_frame
        self.proposed_frame = proposed_frame
        self.is_processing = False
        self.active_output_root = ""
        self.setup_ui()

    def setup_ui(self):
        # Process button - prominent accent style
        self.process_button = ttk.Button(
            self, text="Process Files", style="Accent.TButton",
            command=self.start_processing,
        )
        self.process_button.pack(fill="x", pady=(0, 10))

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self, variable=self.progress_var, maximum=100,
        )
        self.progress_bar.pack(fill="x", pady=(0, 12))

        # Log header
        ttk.Label(self, text="Log", style="Heading.TLabel").pack(anchor="w", pady=(0, 4))

        # Log output
        self.output_text = scrolledtext.ScrolledText(
            self,
            height=8,
            wrap=tk.WORD,
            bg=THEME["bg_card"],
            fg=THEME["text"],
            insertbackground=THEME["purple"],
            font=("Consolas", 9),
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightcolor=THEME["border"],
            highlightbackground=THEME["border"],
        )
        self.output_text.pack(fill="both", expand=True)

    def log_message(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}\n"

        def _update():
            self.output_text.insert(tk.END, line)
            self.output_text.see(tk.END)

        self.output_text.after(0, _update)

    def start_processing(self):
        if self.is_processing:
            return
        existing_files = self.existing_frame.get_files()
        proposed_files = self.proposed_frame.get_files()
        if not existing_files and not proposed_files:
            self.log_message("ERROR: Please select at least one folder containing PPLX files")
            return

        excel_path = self.config_manager.get("excel_file_path", "")
        if not excel_path or not os.path.exists(excel_path):
            self.log_message("ERROR: Please select a valid Excel file before processing.")
            messagebox.showerror(
                "Excel Required",
                "Please select a valid Excel file before processing.",
            )
            return

        self.is_processing = True
        self.process_button.config(state="disabled", text="Processing...")
        self.output_text.delete(1.0, tk.END)

        category_data = [
            {
                "name": "EXISTING",
                "files": existing_files,
                "source_folder": self.existing_frame.get_source_path(),
            },
            {
                "name": "PROPOSED",
                "files": proposed_files,
                "source_folder": self.proposed_frame.get_source_path(),
            },
        ]

        timestamp = datetime.now().strftime("%Y_%m_%d_%H%M%S")
        project_root = Path(__file__).resolve().parents[3]
        processed_base = project_root / "output"
        processed_base.mkdir(parents=True, exist_ok=True)
        prefix = "O-Calcs"
        if excel_path:
            base_name = os.path.basename(excel_path)
            prefix_candidate = base_name.split(" ")[0].strip() or Path(base_name).stem
            if prefix_candidate:
                prefix = f"{prefix_candidate}_O-Calcs"

        output_root = str(processed_base / f"{prefix}_{timestamp}")
        self.active_output_root = output_root

        thread = threading.Thread(
            target=self.process_files,
            args=(category_data, output_root, timestamp),
        )
        thread.daemon = True
        thread.start()

    def process_files(self, category_data: List[Dict], output_root: str, timestamp: str):
        try:
            os.makedirs(output_root, exist_ok=True)
            excel_path = self.config_manager.get("excel_file_path", "")
            excel_data = load_excel_data(excel_path, log_callback=self.log_message)
            valid_scids = set(excel_data.keys()) if excel_data else set()

            total_files = sum(len(c["files"]) for c in category_data)
            if total_files == 0:
                self.log_message("No PPLX files found to process.")
                return

            self.log_message(f"Starting processing of {total_files} files across categories")
            self.log_message(f"Output root directory: {output_root}")
            if excel_data:
                self.log_message(f"{len(valid_scids)} non-underground poles found in Excel data")
            else:
                self.log_message("No Excel data loaded - processing all files")
            processed_count = 0
            summary: Dict[str, Dict] = {}

            for category in category_data:
                name = category["name"]
                files = category["files"]
                source_folder = category.get("source_folder") or ""
                output_dir = os.path.join(output_root, name)
                os.makedirs(output_dir, exist_ok=True)
                summary[name] = {
                    "successful": 0, "failed": 0, "skipped": 0,
                    "csv_path": "", "output_dir": output_dir,
                }

                if not files:
                    self.log_message(f"\nCategory: {name} (no files selected, directory created)")
                    continue

                condition_value = name
                csv_data = []
                wire_spec_data = []

                self.log_message(f"\nCategory: {name}")
                if source_folder:
                    self.log_message(f" Source: {source_folder}")
                self.log_message(f" Processing {len(files)} file{'s' if len(files) != 1 else ''}")

                proc_kwargs = dict(
                    condition_value=condition_value,
                    output_dir=output_dir,
                    excel_data=excel_data,
                    valid_scids=valid_scids,
                    auto_fill_aux1=self.config_manager.get("auto_fill_aux1", False),
                    auto_fill_aux2=self.config_manager.get("auto_fill_aux2", False),
                    keyword_payload={
                        "comm_keywords": parse_keywords(self.config_manager.get("comm_keywords", "")),
                        "power_keywords": parse_keywords(self.config_manager.get("power_keywords", "")),
                        "pco_keywords": parse_keywords(self.config_manager.get("pco_keywords", "")),
                        "aux5_keywords": parse_keywords(self.config_manager.get("aux5_keywords", "")),
                        "power_label": self.config_manager.get("power_label", "POWER"),
                    },
                )

                def _process(fp):
                    return process_single_file(fp, **proc_kwargs)

                max_workers = min(8, max(1, (os.cpu_count() or 1)))
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    for result in executor.map(_process, files):
                        for entry in result["logs"]:
                            self.log_message(entry)
                        status = result["status"]
                        if status == "success":
                            summary[name]["successful"] += 1
                            if result["csv_row"]:
                                csv_data.append(result["csv_row"])
                        elif status == "failed":
                            summary[name]["failed"] += 1
                        elif status == "skipped":
                            summary[name]["skipped"] += 1
                        processed_count += 1
                        pct = (processed_count / total_files) * 100
                        self.progress_bar.after(0, lambda p=pct: self.progress_var.set(p))

                # OPPD config: wire spec comparison
                if self.config_manager.config_name == "OPPD" and files:
                    if not build_wire_spec_comparison:
                        self.log_message(f"  Wire spec skipped: {_WIRE_SPEC_IMPORT_ERROR or 'import unavailable'}")
                    else:
                        shape_base = _wire_spec_base_path()
                        if excel_path and shape_base.exists():
                            self.log_message("  Starting wire spec comparison...")
                            try:
                                wire_spec_data = build_wire_spec_comparison(
                                    Path(excel_path), files, shape_base, extract_scid_from_filename,
                                    log_callback=self.log_message,
                                )
                                if wire_spec_data:
                                    self.log_message(f"  Wire spec comparison: {len(wire_spec_data)} rows")
                            except Exception as e:
                                self.log_message(f"  Wire spec comparison failed: {e}")
                        else:
                            self.log_message(f"  Wire spec skipped: shape dir not found at {shape_base}")

                if csv_data or wire_spec_data:
                    change_log_path = os.path.join(output_root, f"{name}_change_log_{timestamp}.xlsx")
                    wire_spec_mapping = self.config_manager.get("wire_spec_mapping", {})
                    try:
                        if write_change_log(change_log_path, csv_data, wire_spec_data, wire_spec_mapping):
                            self.log_message(f"{name} change log saved: {change_log_path}")
                            summary[name]["csv_path"] = change_log_path
                        else:
                            self.log_message("openpyxl not available; skipping change_log.xlsx")
                    except Exception as e:
                        self.log_message(f"Error saving {name} change log: {str(e)}")

            self.progress_bar.after(0, lambda: self.progress_var.set(100))
            self.log_message("\nProcessing complete!")
            for name, stats in summary.items():
                self.log_message(
                    f"{name} - Successful: {stats['successful']}, "
                    f"Failed: {stats['failed']}, Skipped: {stats['skipped']}"
                )
                if stats.get("csv_path"):
                    self.log_message(f"{name} CSV: {stats['csv_path']}")
                self.log_message(f"{name} Output: {stats['output_dir']}")
            self.log_message(f"Output root directory: {output_root}")

            try:
                if os.name == "nt":
                    os.startfile(output_root)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", output_root])
                else:
                    subprocess.Popen(["xdg-open", output_root])
            except Exception as open_err:
                self.log_message(f"Info: Unable to open output folder: {open_err}")

        except Exception as e:
            self.log_message(f"Critical error: {str(e)}")
            self.output_text.after(0, lambda: messagebox.showerror("Error", f"Processing failed: {str(e)}"))
        finally:
            self.is_processing = False
            self.process_button.after(0, lambda: self.process_button.config(state="normal", text="Process Files"))
