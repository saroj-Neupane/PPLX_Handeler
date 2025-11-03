#!/usr/bin/env python3
"""
PPLX GUI - Graphical User Interface for PPLX File Editor

A user-friendly GUI application for editing pplx XML files, specifically
designed for managing Aux Data fields in pole line engineering files.

Features:
- File selection and browsing
- Visual editing of Aux Data fields
- Automatic output to 'Processed PPLX' folder
- Path memory using JSON configuration
- Batch processing capabilities

Author: AI Assistant
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import json
from pathlib import Path
from typing import List, Dict, Optional
import threading
from datetime import datetime
import csv

# Import our existing PPLX handler
from pplx_handeler import PPLXHandler, PPLXBatchProcessor

# Import shared configuration and logic
from pplx_config import (
    analyze_mr_note_for_aux_data,
    extract_scid_from_filename,
    clean_scid_keywords,
    normalize_scid_for_excel_lookup,
    DEFAULT_COMM_OWNERS,
    DEFAULT_POWER_OWNERS,
    DEFAULT_IGNORE_SCID_KEYWORDS
)

# Excel support
try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


class PPLXConfigManager:
    """Manages application configuration and path memory."""
    
    def __init__(self, config_file: str = "pplx_gui_config.json"):
        # Try to find the config file in multiple locations
        self.config_file = self._find_config_file(config_file)
        self.config = self.load_config()
    
    def _find_config_file(self, config_file: str) -> str:
        """Find the config file in multiple possible locations."""
        import sys
        
        # Check if we're running from a PyInstaller bundle
        if getattr(sys, 'frozen', False):
            # Running from executable - look in executable directory
            executable_dir = os.path.dirname(sys.executable)
            possible_paths = [
                os.path.join(executable_dir, config_file),  # Same directory as executable
                os.path.join(os.getcwd(), config_file),  # Current working directory
                config_file,  # Current directory
            ]
            
            # Debug: Print paths being checked (only in development)
            if os.getenv('PPLX_DEBUG'):
                print(f"[DEBUG] Running from executable: {sys.executable}")
                print(f"[DEBUG] Executable directory: {executable_dir}")
                print(f"[DEBUG] Checking config paths:")
                for path in possible_paths:
                    print(f"[DEBUG]   - {path} (exists: {os.path.exists(path)})")
        else:
            # Running from Python script - look in script directory
            script_dir = os.path.dirname(os.path.abspath(__file__))
            possible_paths = [
                os.path.join(script_dir, config_file),  # Same directory as script
                os.path.join(os.getcwd(), config_file),  # Current working directory
                config_file,  # Current directory
                os.path.join(os.path.dirname(__file__), config_file),  # Same directory as script
                os.path.join(os.path.dirname(os.path.abspath(__file__)), config_file),  # Script directory
            ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        # If not found anywhere, return the path relative to executable or script
        if getattr(sys, 'frozen', False):
            return os.path.join(os.path.dirname(sys.executable), config_file)
        else:
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), config_file)
    
    def load_config(self) -> Dict:
        """Load configuration from JSON file."""
        default_config = {
            "last_input_directory": "",
            "last_output_directory": "",
            "window_geometry": "1000x700",
            "recent_files": [],
            "default_aux_values": [""] * 8
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                # Merge with defaults to handle new keys
                for key, value in default_config.items():
                    if key not in config:
                        config[key] = value
                return config
        except Exception as e:
            print(f"Error loading config: {e}")
        
        return default_config
    
    def save_config(self):
        """Save configuration to JSON file."""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            print(f"Error saving config: {e}")
    
    def get(self, key: str, default=None):
        """Get configuration value."""
        return self.config.get(key, default)
    
    def set(self, key: str, value):
        """Set configuration value."""
        self.config[key] = value
        self.save_config()
    
    def add_recent_file(self, file_path: str):
        """Add file to recent files list."""
        recent = self.config.get("recent_files", [])
        if file_path in recent:
            recent.remove(file_path)
        recent.insert(0, file_path)
        recent = recent[:10]  # Keep only last 10
        self.set("recent_files", recent)


class PPLXFileListFrame(ttk.Frame):
    """Frame for displaying and managing selected PPLX files."""
    
    def __init__(self, parent, config_manager: PPLXConfigManager):
        super().__init__(parent)
        self.config_manager = config_manager
        self.files = []
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the file list UI."""
        # Title
        title_label = ttk.Label(self, text="PPLX Files in Selected Folder", font=("Arial", 12, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Select folder button
        ttk.Button(self, text="Select Folder", command=self.select_folder).pack(pady=(0, 10))
        
        # File list with scrollbar
        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True)
        
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.file_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Current folder label
        self.folder_label = ttk.Label(self, text="No folder selected", foreground="gray")
        self.folder_label.pack(pady=(5, 0))
        
        # Status label
        self.status_label = ttk.Label(self, text="Select a folder containing PPLX files")
        self.status_label.pack(pady=(10, 0))
    
    def select_folder(self):
        """Select folder containing PPLX files."""
        initial_dir = self.config_manager.get("last_input_directory", "")
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = os.getcwd()
        
        folder = filedialog.askdirectory(
            title="Select Folder with PPLX Files",
            initialdir=initial_dir
        )
        
        if folder:
            self.current_folder = folder
            self.config_manager.set("last_folder_path", folder)
            self.load_folder_files()
    
    def load_folder_files(self):
        """Load all PPLX files from the current folder."""
        if not hasattr(self, 'current_folder') or not self.current_folder:
            return
            
        self.files.clear()
        
        try:
            # Find all pplx files in folder
            pplx_files = []
            for file in os.listdir(self.current_folder):
                if file.lower().endswith('.pplx'):
                    file_path = os.path.join(self.current_folder, file)
                    pplx_files.append(file_path)
            
            self.files = sorted(pplx_files)  # Sort alphabetically
            
            # Update folder label
            self.folder_label.config(text=f"Folder: {os.path.basename(self.current_folder)}", foreground="black")
            
            self.update_display()
            
            # Log file count instead of popup
            if self.files:
                print(f"Found {len(self.files)} PPLX files in folder")
            else:
                print("No PPLX files found in selected folder")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error reading folder: {str(e)}")
    
    def clear_files(self):
        """Clear all files from the list."""
        self.files.clear()
        if hasattr(self, 'current_folder'):
            self.current_folder = ""
        self.folder_label.config(text="No folder selected", foreground="gray")
        self.update_display()
    
    def update_display(self):
        """Update the file list display."""
        self.file_listbox.delete(0, tk.END)
        
        for file_path in self.files:
            filename = os.path.basename(file_path)
            self.file_listbox.insert(tk.END, filename)
        
        count = len(self.files)
        if count > 0:
            self.status_label.config(text=f"{count} PPLX file{'s' if count != 1 else ''} found")
        else:
            self.status_label.config(text="Select a folder containing PPLX files")
    
    def get_files(self) -> List[str]:
        """Get the list of selected files."""
        return self.files.copy()


class AuxDataEditFrame(ttk.Frame):
    """Frame for editing Aux Data fields."""
    
    def __init__(self, parent, config_manager: PPLXConfigManager):
        super().__init__(parent)
        self.config_manager = config_manager
        self.aux_entries = []
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the Aux Data editing UI."""
        # Title
        title_label = ttk.Label(self, text="Aux Data Fields", font=("Arial", 12, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Instructions
        instructions = ttk.Label(
            self, 
            text="Enter values for Aux Data 1 and 3. Other fields will be auto-filled based on logic.",
            foreground="gray"
        )
        instructions.pack(pady=(0, 10))
        
        # Excel file selection
        excel_frame = ttk.Frame(self)
        excel_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        ttk.Label(excel_frame, text="Excel Data File:", width=15).pack(side="left")
        self.excel_file_var = tk.StringVar()
        self.excel_file_var.set(self.config_manager.get("excel_file_path", "No file selected"))
        
        excel_label = ttk.Label(excel_frame, textvariable=self.excel_file_var, foreground="blue")
        excel_label.pack(side="left", fill="x", expand=True, padx=(5, 5))
        
        ttk.Button(excel_frame, text="Browse", command=self.select_excel_file).pack(side="left")
        
        # Owner configuration frame
        owners_frame = ttk.LabelFrame(self, text="Owner Configuration", padding=10)
        owners_frame.pack(fill="x", padx=10, pady=(10, 0))
        
        # Comm Owners
        comm_frame = ttk.Frame(owners_frame)
        comm_frame.pack(fill="x", pady=2)
        ttk.Label(comm_frame, text="Comm Owners:", width=15).pack(side="left")
        self.comm_owners_var = tk.StringVar()
        self.comm_owners_var.set(self.config_manager.get("comm_owners", DEFAULT_COMM_OWNERS))
        comm_entry = ttk.Entry(comm_frame, textvariable=self.comm_owners_var)
        comm_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        comm_entry.bind('<FocusOut>', self.save_owner_config)
        
        # Power Owners
        power_frame = ttk.Frame(owners_frame)
        power_frame.pack(fill="x", pady=2)
        ttk.Label(power_frame, text="Power Owners:", width=15).pack(side="left")
        self.power_owners_var = tk.StringVar()
        self.power_owners_var.set(self.config_manager.get("power_owners", DEFAULT_POWER_OWNERS))
        power_entry = ttk.Entry(power_frame, textvariable=self.power_owners_var)
        power_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        power_entry.bind('<FocusOut>', self.save_owner_config)
        
        # Instructions
        ttk.Label(owners_frame, text="Separate multiple owners with commas", 
                 foreground="gray", font=("Arial", 8)).pack(pady=(5, 0))
        
        # SCID Keywords Configuration frame
        scid_frame = ttk.LabelFrame(self, text="SCID Keywords Configuration", padding=10)
        scid_frame.pack(fill="x", padx=10, pady=(10, 0))
        
        # Ignore SCID Keywords
        ignore_frame = ttk.Frame(scid_frame)
        ignore_frame.pack(fill="x", pady=2)
        ttk.Label(ignore_frame, text="Ignore SCID Keywords:", width=20).pack(side="left")
        self.ignore_scid_keywords_var = tk.StringVar()
        self.ignore_scid_keywords_var.set(self.config_manager.get("ignore_scid_keywords", DEFAULT_IGNORE_SCID_KEYWORDS))
        ignore_entry = ttk.Entry(ignore_frame, textvariable=self.ignore_scid_keywords_var)
        ignore_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        ignore_entry.bind('<FocusOut>', self.save_scid_config)
        
        # Instructions for SCID keywords
        ttk.Label(scid_frame, text="Keywords to remove from SCID before processing (comma-separated)", 
                 foreground="gray", font=("Arial", 8)).pack(pady=(5, 0))
        
        # Define the 5 Aux Data fields with their descriptions
        self.aux_fields = [
            ("Aux Data 1", "Pole Owner", True),     # User editable or auto-filled
            ("Aux Data 2", "Pole Tag", False),      # Auto-filled
            ("Aux Data 3", "Condition", True),      # User editable  
            ("Aux Data 4", "Make Ready Type", False), # Auto-filled
            ("Aux Data 5", "Proposed Riser", False)  # Auto-filled
        ]
        
        # Create entry fields for editable fields and labels for auto-filled
        fields_frame = ttk.Frame(self)
        fields_frame.pack(fill="x", padx=10)
        
        saved_values = self.config_manager.get("last_aux_values", {})
        
        for i, (field_name, description, is_editable) in enumerate(self.aux_fields):
            row_frame = ttk.Frame(fields_frame)
            row_frame.pack(fill="x", pady=2)
            
            # Label with description in parentheses
            label_text = f"{field_name} ({description}):"
            label = ttk.Label(row_frame, text=label_text, width=30)
            label.pack(side="left")
            
            if is_editable:
                # Create entry for user input
                entry = ttk.Entry(row_frame)
                entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
                
                # Load saved value if available first
                saved_key = f"aux_data_{i+1}"
                if saved_key in saved_values:
                    entry.insert(0, saved_values[saved_key])
                
                # Add checkbox for Aux Data 1 auto-fill
                if i == 0:  # Aux Data 1
                    self.auto_fill_aux1_var = tk.BooleanVar()
                    auto_checkbox = ttk.Checkbutton(
                        row_frame, 
                        text="Auto-fill from Excel", 
                        variable=self.auto_fill_aux1_var,
                        command=self.toggle_aux1_auto_fill
                    )
                    auto_checkbox.pack(side="left", padx=(5, 0))
                    
                    # Load saved checkbox state and apply it
                    checkbox_state = self.config_manager.get("auto_fill_aux1", False)
                    self.auto_fill_aux1_var.set(checkbox_state)
                    
                    # Apply initial state without triggering the callback
                    if checkbox_state:
                        entry.config(state="readonly")
                        entry.delete(0, tk.END)
                        entry.insert(0, "(Will auto-fill from Excel)")
                
                # Bind to auto-save when value changes
                entry.bind('<FocusOut>', self.auto_save_values)
                entry.bind('<KeyRelease>', self.auto_save_values)
                
                self.aux_entries.append(entry)
            else:
                # Show "(Auto Filled)" for non-editable fields
                auto_label = ttk.Label(row_frame, text="(Auto Filled)", foreground="gray", style="Gray.TLabel")
                auto_label.pack(side="left", fill="x", expand=True, padx=(5, 0))
                self.aux_entries.append(auto_label)
    
    def select_excel_file(self):
        """Select Excel file for data lookup."""
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_file_var.set(file_path)
            self.config_manager.set("excel_file_path", file_path)
    
    def auto_save_values(self, event=None):
        """Auto-save current values when they change."""
        values = {}
        entry_index = 0
        for i, (field_name, description, is_editable) in enumerate(self.aux_fields):
            if is_editable:
                entry = self.aux_entries[entry_index]
                if hasattr(entry, 'get'):  # Make sure it's an Entry widget
                    value = entry.get().strip()
                    values[f"aux_data_{i+1}"] = value
                entry_index += 1
            else:
                entry_index += 1
        
        self.config_manager.set("last_aux_values", values)
    
    def toggle_aux1_auto_fill(self):
        """Toggle auto-fill for Aux Data 1 and save state."""
        state = self.auto_fill_aux1_var.get()
        self.config_manager.set("auto_fill_aux1", state)
        
        if state:
            # If auto-fill is enabled, disable the entry and try to fill from Excel
            entry = self.aux_entries[0]
            entry.config(state="readonly")
            self.fill_aux1_from_excel()
        else:
            # If auto-fill is disabled, enable the entry for manual input
            entry = self.aux_entries[0]
            entry.config(state="normal")
    
    def fill_aux1_from_excel(self):
        """Fill Aux Data 1 from Excel pole_tag_company column."""
        # This will be called when a file is selected for processing
        # For now, just show placeholder
        entry = self.aux_entries[0]
        entry.config(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, "(Will auto-fill from Excel)")
        entry.config(state="readonly")
    
    def save_owner_config(self, event=None):
        """Save owner configuration when values change."""
        self.config_manager.set("comm_owners", self.comm_owners_var.get())
        self.config_manager.set("power_owners", self.power_owners_var.get())
    
    def save_scid_config(self, event=None):
        """Save SCID configuration when values change."""
        self.config_manager.set("ignore_scid_keywords", self.ignore_scid_keywords_var.get())
    
    def analyze_mr_note(self, mr_note: str) -> tuple:
        """Analyze mr_note to determine Aux Data 4 and 5 values."""
        # Get owner lists from GUI configuration
        comm_owners = self.comm_owners_var.get().split(',')
        power_owners = self.power_owners_var.get().split(',')
        
        return analyze_mr_note_for_aux_data(mr_note, comm_owners, power_owners)
    
    def get_aux_values(self) -> Dict[int, str]:
        """Get the current Aux Data values for processing."""
        values = {}
        entry_index = 0
        for i, (field_name, description, is_editable) in enumerate(self.aux_fields):
            if is_editable:
                entry = self.aux_entries[entry_index]
                if hasattr(entry, 'get'):  # Make sure it's an Entry widget
                    value = entry.get().strip()
                    if value:  # Only include non-empty values
                        values[i + 1] = value
                entry_index += 1
            else:
                # For auto-filled fields, get the current auto-filled value
                # This will be implemented when we add the logic
                entry_index += 1
        return values
    
    def set_readonly_field(self, field_index: int, value: str):
        """Set value in a readonly field (0-based index)."""
        if 0 <= field_index < len(self.aux_entries):
            entry = self.aux_entries[field_index]
            # Temporarily enable readonly fields to set value
            was_readonly = str(entry.cget('state')) == 'readonly'
            if was_readonly:
                entry.config(state='normal')
            entry.delete(0, tk.END)
            entry.insert(0, value)
            if was_readonly:
                entry.config(state='readonly')
    
    
    def load_excel_data(self, log_callback=None) -> Dict[str, Dict]:
        """Load and filter Excel data based on node_type=pole and pole_status!=underground"""
        excel_path = self.config_manager.get("excel_file_path", "")
        if log_callback:
            log_callback(f"Looking for Excel file at: {excel_path}")
            log_callback(f"Current working directory: {os.getcwd()}")
        
        if not excel_path or not os.path.exists(excel_path):
            if log_callback:
                log_callback(f"Excel Error: File not found or path empty: {excel_path}")
                log_callback(f"Please check the Excel file path in the configuration")
            return {}
        
        if not EXCEL_AVAILABLE:
            if log_callback:
                log_callback("Excel Support: openpyxl library not available. Please install it.")
            return {}
        
        try:
            workbook = openpyxl.load_workbook(excel_path)
            if 'nodes' not in workbook.sheetnames:
                if log_callback:
                    log_callback("Excel Error: No 'nodes' sheet found in Excel file")
                return {}
            
            sheet = workbook['nodes']
            data = {}
            
            # Get headers
            headers = {}
            for col in range(1, sheet.max_column + 1):
                header = sheet.cell(row=1, column=col).value
                if header:
                    headers[header.lower()] = col
            
            # Check required columns
            required_cols = ['scid', 'node_type', 'pole_status']
            missing_cols = [col for col in required_cols if col not in headers]
            if missing_cols:
                if log_callback:
                    log_callback(f"Excel Error: Missing columns in Excel file: {missing_cols}")
                return {}
            
            # Read data with filtering
            valid_count = 0
            total_count = 0
            for row in range(2, sheet.max_row + 1):
                total_count += 1
                scid = sheet.cell(row=row, column=headers['scid']).value
                node_type = sheet.cell(row=row, column=headers['node_type']).value
                pole_status = sheet.cell(row=row, column=headers['pole_status']).value
                
                # Apply filters: node_type = pole and pole_status != 'underground'
                if (node_type and str(node_type).lower() == 'pole' and 
                    pole_status and str(pole_status).lower() != 'underground'):
                    
                    # Store all row data for this SCID
                    row_data = {}
                    for col_name, col_num in headers.items():
                        cell_value = sheet.cell(row=row, column=col_num).value
                        row_data[col_name] = str(cell_value) if cell_value is not None else ""
                    
                    data[str(scid)] = row_data
                    valid_count += 1
            
            workbook.close()
            if log_callback:
                log_callback(f"Excel loaded successfully: {valid_count} valid entries from {total_count} total rows")
            return data
            
        except Exception as e:
            if log_callback:
                log_callback(f"Excel Error: Error reading Excel file: {str(e)}")
            return {}
    
    def get_valid_scids(self) -> set:
        """Get set of valid SCIDs from Excel data"""
        excel_data = self.load_excel_data()
        return set(excel_data.keys())
    
    def apply_auto_fill_logic(self, scid: str, excel_data: Dict[str, Dict] = None, log_callback=None):
        """Apply logic to auto-fill Aux Data 2, 4, and 5 based on SCID and Excel data."""
        if excel_data is None:
            excel_data = self.load_excel_data(log_callback=log_callback)
        
        if scid not in excel_data:
            # Set default values if SCID not found
            self.set_readonly_field(1, "NO DATA")  # Aux Data 2
            self.set_readonly_field(3, "NO DATA")  # Aux Data 4  
            self.set_readonly_field(4, "NO DATA")  # Aux Data 5
            return
        
        row_data = excel_data[scid]
        
        # Auto-fill logic - you can customize these based on Excel columns
        # Aux Data 2 (Pole Tag) - using SCID as tag for now
        pole_tag = f"POLE_{scid}"
        self.set_readonly_field(1, pole_tag)
        
        # Aux Data 4 (Make Ready Type) - placeholder logic
        # You can map this to specific Excel columns
        make_ready_type = row_data.get('make_ready_type', 'STANDARD')
        self.set_readonly_field(3, make_ready_type)
        
        # Aux Data 5 (Proposed) - placeholder logic  
        # You can map this to specific Excel columns
        proposed = row_data.get('proposed_status', 'NEW')
        self.set_readonly_field(4, proposed)



class ProcessingFrame(ttk.Frame):
    """Frame for processing controls and output."""
    
    def __init__(self, parent, config_manager: PPLXConfigManager, file_frame: PPLXFileListFrame, aux_frame: AuxDataEditFrame):
        super().__init__(parent)
        self.config_manager = config_manager
        self.file_frame = file_frame
        self.aux_frame = aux_frame
        self.is_processing = False
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the processing UI."""
        # Title
        title_label = ttk.Label(self, text="Processing", font=("Arial", 12, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Output information
        info_frame = ttk.Frame(self)
        info_frame.pack(fill="x", pady=(0, 10))
        
        info_text = ttk.Label(
            info_frame, 
            text="Modified files will be saved to 'Modified PPLX' folder\ninside the selected folder",
            foreground="gray",
            justify="center"
        )
        info_text.pack()
        
        # Process button
        self.process_button = ttk.Button(self, text="Process Files", command=self.start_processing)
        self.process_button.pack(pady=(0, 10))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", pady=(0, 10))
        
        # Output text area
        output_frame = ttk.LabelFrame(self, text="Processing Log", padding=5)
        output_frame.pack(fill="both", expand=True)
        
        self.output_text = scrolledtext.ScrolledText(output_frame, height=10, wrap=tk.WORD)
        self.output_text.pack(fill="both", expand=True)
    
    def get_output_directory(self) -> str:
        """Get the output directory based on the selected folder."""
        if hasattr(self.file_frame, 'current_folder') and self.file_frame.current_folder:
            return os.path.join(self.file_frame.current_folder, "Modified PPLX")
        else:
            # Fallback to current directory if no folder selected
            return os.path.join(os.getcwd(), "Modified PPLX")
    
    def log_message(self, message: str):
        """Add a message to the output log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.output_text.insert(tk.END, formatted_message)
        self.output_text.see(tk.END)
        self.output_text.update()
    
    def start_processing(self):
        """Start the file processing in a separate thread."""
        if self.is_processing:
            return
        
        files = self.file_frame.get_files()
        if not files:
            self.log_message("ERROR: Please select a folder containing PPLX files")
            return
        
        # Check if folder is selected
        if not hasattr(self.file_frame, 'current_folder') or not self.file_frame.current_folder:
            self.log_message("ERROR: Please select a folder first")
            return
        
        aux_values = self.aux_frame.get_aux_values()
        if not aux_values:
            self.log_message("WARNING: No Aux Data values specified. Files will be copied with auto-filled data only.")
        
        # Disable the process button
        self.is_processing = True
        self.process_button.config(state="disabled", text="Processing...")
        
        # Clear the output log
        self.output_text.delete(1.0, tk.END)
        
        # Start processing in a separate thread
        thread = threading.Thread(target=self.process_files, args=(files, aux_values))
        thread.daemon = True
        thread.start()
    
    def process_files(self, files: List[str], aux_values: Dict[int, str]):
        """Process the selected files with SCID filtering."""
        try:
            output_dir = self.get_output_directory()
            
            # Create output directory if it doesn't exist
            os.makedirs(output_dir, exist_ok=True)
            
            # Load Excel data and get valid SCIDs
            excel_data = self.aux_frame.load_excel_data(log_callback=self.log_message)
            valid_scids = set(excel_data.keys()) if excel_data else set()
            
            self.log_message(f"Starting processing of {len(files)} files")
            self.log_message(f"Output directory: {output_dir}")
            
            if excel_data:
                self.log_message(f"{len(valid_scids)} non-underground poles found in Excel data")
            else:
                self.log_message("No Excel data loaded - processing all files")
            
            if aux_values:
                self.log_message(f"User Aux Data values: {aux_values}")
            
            successful = 0
            failed = 0
            skipped = 0
            
            # Prepare CSV data for log file
            csv_data = []
            
            for i, file_path in enumerate(files):
                try:
                    # Update progress
                    progress = (i / len(files)) * 100
                    self.progress_var.set(progress)
                    
                    filename = os.path.basename(file_path)
                    
                    # Extract SCID from filename (for Excel lookup - no keyword filtering)
                    scid = extract_scid_from_filename(filename)
                    
                    # Extract pole number from original filename and apply SCID keyword filtering (for filename generation)
                    pole_number = extract_scid_from_filename(filename)
                    clean_pole_number = clean_scid_keywords(pole_number, self.aux_frame.ignore_scid_keywords_var.get())
                    
                    # Check if SCID is valid (if Excel data is loaded)
                    if excel_data and scid not in valid_scids:
                        self.log_message(f"Skipping {filename}: SCID '{scid}' not found in Excel data")
                        skipped += 1
                        continue
                    
                    self.log_message(f"Processing: {filename} (SCID: {scid}, Pole Number: {pole_number} -> {clean_pole_number})")
                    
                    # Load the file
                    handler = PPLXHandler(file_path)
                    
                    # Apply user-entered Aux Data modifications
                    if aux_values:
                        for aux_num, value in aux_values.items():
                            success = handler.set_aux_data(aux_num, value)
                            if success:
                                self.log_message(f"  Set Aux Data {aux_num}: {value}")
                            else:
                                self.log_message(f"  Warning: Could not set Aux Data {aux_num}")
                    
                    # Apply auto-filled data from Excel
                    if excel_data and scid in excel_data:
                        row_data = excel_data[scid]
                        
                        # Auto-fill Aux Data 1 (Pole Owner) if checkbox is checked
                        if hasattr(self.aux_frame, 'auto_fill_aux1_var') and self.aux_frame.auto_fill_aux1_var.get():
                            pole_owner = row_data.get('pole_tag_company', '')
                            if pole_owner:
                                handler.set_aux_data(1, pole_owner)
                                self.log_message(f"  Auto-filled Aux Data 1: {pole_owner}")
                        
                        # Auto-fill Aux Data 2 (Pole Tag) from pole_tag_tagtext
                        pole_tag = row_data.get('pole_tag_tagtext', f"POLE_{scid}")
                        handler.set_aux_data(2, pole_tag)
                        self.log_message(f"  Auto-filled Aux Data 2: {pole_tag}")
                        
                        # Analyze mr_note for Aux Data 4 and 5
                        mr_note = row_data.get('mr_note', '')
                        aux_data_4, aux_data_5 = self.aux_frame.analyze_mr_note(mr_note)
                        
                        # Auto-fill Aux Data 4 (Make Ready Type)
                        handler.set_aux_data(4, aux_data_4)
                        self.log_message(f"  Auto-filled Aux Data 4: {aux_data_4}")
                        if mr_note:
                            self.log_message(f"    Based on mr_note: {mr_note[:50]}{'...' if len(mr_note) > 50 else ''}")
                        
                        # Auto-fill Aux Data 5 (Proposed Riser)
                        handler.set_aux_data(5, aux_data_5)
                        self.log_message(f"  Auto-filled Aux Data 5: {aux_data_5}")
                    
                    # Get pole tag from Excel or user input
                    pole_tag = "Unknown"
                    if excel_data and scid in excel_data:
                        pole_tag = excel_data[scid].get('pole_tag_tagtext', pole_tag)
                    
                    # Get condition from user input
                    condition = aux_values.get(3, "Unknown")  # Aux Data 3
                    
                    # Get MR Note from Excel data
                    mr_note = ""
                    if excel_data and scid in excel_data:
                        mr_note = excel_data[scid].get('mr_note', '')
                    
                    # Get final Aux Data values from the handler
                    final_aux_data = handler.get_aux_data()
                    
                    # Clean values for filename (remove invalid characters)
                    clean_pole_number_safe = ''.join(c for c in clean_pole_number if c.isalnum() or c in '-_')
                    clean_pole_tag = ''.join(c for c in str(pole_tag) if c.isalnum() or c in '-_')
                    clean_condition = ''.join(c for c in str(condition) if c.isalnum() or c in '-_')
                    
                    # Create the new filename using cleaned pole number
                    new_filename = f"{clean_pole_number_safe}_{clean_pole_tag}_{clean_condition}.pplx"
                    output_file = os.path.join(output_dir, new_filename)
                    
                    # Set Pole Number from cleaned pole number
                    handler.set_pole_attribute('Pole Number', clean_pole_number)
                    self.log_message(f"  Set Pole Number: {clean_pole_number}")
                    
                    # Set DescriptionOverride to match the output filename (without extension)
                    description_override = os.path.splitext(new_filename)[0]  # Remove .pplx extension
                    handler.set_pole_attribute('DescriptionOverride', description_override)
                    self.log_message(f"  Set DescriptionOverride: {description_override}")
                    
                    # Save the modified file (overwrite if exists)
                    handler.save_file(output_file)
                    self.log_message(f"  Saved: {os.path.basename(output_file)}")
                    
                    # Add data to CSV
                    csv_row = {
                        'File Name': filename,
                        'MR Note': mr_note,
                        'Aux Data 1': final_aux_data.get('Aux Data 1', 'Unset'),
                        'Aux Data 2': final_aux_data.get('Aux Data 2', 'Unset'),
                        'Aux Data 3': final_aux_data.get('Aux Data 3', 'Unset'),
                        'Aux Data 4': final_aux_data.get('Aux Data 4', 'Unset'),
                        'Aux Data 5': final_aux_data.get('Aux Data 5', 'Unset')
                    }
                    csv_data.append(csv_row)
                    
                    successful += 1
                    
                except Exception as e:
                    self.log_message(f"  Error processing {filename}: {str(e)}")
                    failed += 1
            
            # Update progress to 100%
            self.progress_var.set(100)
            
            # Generate CSV log file
            if csv_data:
                csv_file_path = os.path.join(output_dir, "log.csv")
                try:
                    with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
                        fieldnames = ['File Name', 'MR Note', 'Aux Data 1', 'Aux Data 2', 'Aux Data 3', 'Aux Data 4', 'Aux Data 5']
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        writer.writeheader()
                        writer.writerows(csv_data)
                    
                    self.log_message(f"CSV log file saved: {csv_file_path}")
                except Exception as e:
                    self.log_message(f"Error saving CSV log file: {str(e)}")
            
            # Final summary
            self.log_message(f"\nProcessing complete!")
            self.log_message(f"Successful: {successful}")
            self.log_message(f"Failed: {failed}")
            if excel_data:
                self.log_message(f"Skipped (SCID not in Excel): {skipped}")
            self.log_message(f"Output directory: {output_dir}")
            
            # Log completion summary
            if failed == 0 and skipped == 0:
                self.log_message(f"SUCCESS: All {successful} files processed successfully!")
            elif failed == 0:
                self.log_message(f"COMPLETED: {successful} files processed successfully, {skipped} skipped (SCID not in Excel data)")
            else:
                message = f"COMPLETED: {successful} files processed successfully, {failed} failed"
                if skipped > 0:
                    message += f", {skipped} skipped (SCID not in Excel data)"
                self.log_message(message)
        
        except Exception as e:
            self.log_message(f"Critical error: {str(e)}")
            messagebox.showerror("Error", f"Processing failed: {str(e)}")
        
        finally:
            # Re-enable the process button
            self.is_processing = False
            self.process_button.config(state="normal", text="Process Files")


class PPLXGUIApp:
    """Main PPLX GUI Application."""
    
    def __init__(self):
        self.config_manager = PPLXConfigManager()
        self.root = tk.Tk()
        self.setup_window()
        self.setup_ui()
        self.setup_menu()
    
    def setup_window(self):
        """Setup the main window."""
        self.root.title("PPLX File Editor - Pole Line Engineering Tools")
        self.root.geometry(self.config_manager.get("window_geometry", "1000x700"))
        
        # Set minimum size
        self.root.minsize(800, 600)
        
        # Configure style
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
        
        # Bind window close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_ui(self):
        """Setup the main UI."""
        # Create main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # Create paned window for resizable sections
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True)
        
        # Left panel - File selection and Aux Data editing
        left_panel = ttk.Frame(paned)
        paned.add(left_panel, weight=1)
        
        # File list frame
        self.file_frame = PPLXFileListFrame(left_panel, self.config_manager)
        self.file_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Auto-load last used folder
        last_folder = self.config_manager.get("last_folder_path", "")
        if last_folder and os.path.exists(last_folder):
            self.file_frame.current_folder = last_folder
            self.file_frame.load_folder_files()
        
        # Aux data editing frame
        self.aux_frame = AuxDataEditFrame(left_panel, self.config_manager)
        self.aux_frame.pack(fill="x")
        
        # Right panel - Processing controls and output
        right_panel = ttk.Frame(paned)
        paned.add(right_panel, weight=1)
        
        self.processing_frame = ProcessingFrame(right_panel, self.config_manager, self.file_frame, self.aux_frame)
        self.processing_frame.pack(fill="both", expand=True)
    
    def setup_menu(self):
        """Setup minimal menu bar."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Help menu only
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
    

    
    def show_batch_report(self):
        """Show batch report dialog."""
        files = self.file_frame.get_files()
        if not files:
            messagebox.showwarning("No Files", "Please select PPLX files first")
            return
        
        # Create a new window for the report
        report_window = tk.Toplevel(self.root)
        report_window.title("Batch Report")
        report_window.geometry("600x400")
        
        # Create report text
        report_text = scrolledtext.ScrolledText(report_window, wrap=tk.WORD)
        report_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Generate report
        report_text.insert(tk.END, f"PPLX Files Batch Report\n")
        report_text.insert(tk.END, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        report_text.insert(tk.END, f"Total Files: {len(files)}\n\n")
        
        for i, file_path in enumerate(files, 1):
            try:
                handler = PPLXHandler(file_path)
                info = handler.get_file_info()
                aux_data = handler.get_aux_data()
                
                report_text.insert(tk.END, f"{i}. {os.path.basename(file_path)}\n")
                report_text.insert(tk.END, f"   Date: {info.get('date', 'Unknown')}\n")
                report_text.insert(tk.END, f"   User: {info.get('user', 'Unknown')}\n")
                
                # Show non-default aux data
                non_default = {k: v for k, v in aux_data.items() if v != 'Unset'}
                if non_default:
                    report_text.insert(tk.END, f"   Aux Data: {non_default}\n")
                else:
                    report_text.insert(tk.END, f"   Aux Data: All unset\n")
                
                report_text.insert(tk.END, "\n")
                
            except Exception as e:
                report_text.insert(tk.END, f"{i}. {os.path.basename(file_path)}\n")
                report_text.insert(tk.END, f"   Error: {str(e)}\n\n")
    
    def export_structure(self):
        """Export XML structure of selected file."""
        files = self.file_frame.get_files()
        if not files:
            messagebox.showwarning("No Files", "Please select a PPLX file first")
            return
        
        # Use first file for structure export
        file_path = files[0]
        
        # Ask for output location
        output_file = filedialog.asksaveasfilename(
            title="Export Structure to JSON",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if output_file:
            try:
                handler = PPLXHandler(file_path)
                handler.export_structure_to_json(output_file)
                messagebox.showinfo("Success", f"Structure exported to:\n{output_file}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export structure:\n{str(e)}")
    
    def show_about(self):
        """Show about dialog."""
        about_text = """PPLX File Editor v1.0

A GUI application for editing PPLX XML files used in pole line engineering.

Features:
• Edit Aux Data fields (1-8)
• Batch processing capabilities
• Automatic output organization
• Path memory and recent files
• Structure analysis and reporting

Created for pole line engineering workflow optimization."""
        
        messagebox.showinfo("About PPLX File Editor", about_text)
    
    def on_closing(self):
        """Handle application closing."""
        # Save window geometry
        self.config_manager.set("window_geometry", self.root.geometry())
        
        # Close the application
        self.root.destroy()
    
    def run(self):
        """Run the application."""
        self.root.mainloop()


def main():
    """Main function to run the GUI application."""
    try:
        app = PPLXGUIApp()
        app.run()
    except Exception as e:
        print(f"Error starting application: {e}")
        messagebox.showerror("Startup Error", f"Failed to start application:\n{str(e)}")


if __name__ == "__main__":
    main() 