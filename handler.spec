# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for PPLX GUI Application
"""

import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Paths
current_dir = os.getcwd()
main_script = os.path.join(current_dir, "app.py")
icon_file = os.path.join(current_dir, "assets", "handler.ico")

# Data files: config folder (includes _active.json, state.json, OPPD.json), assets folder
datas = []

# Config folder
config_dir = os.path.join(current_dir, "config")
if os.path.isdir(config_dir):
    datas.append((config_dir, "config"))

# Assets (icon)
assets_dir = os.path.join(current_dir, "assets")
if os.path.isdir(assets_dir):
    datas.append((assets_dir, "assets"))

# Fallback: single icon file
if os.path.exists(icon_file):
    datas.append((icon_file, "assets"))

# Optional: pplx_structure_sample.json if exists
sample = os.path.join(current_dir, "pplx_structure_sample.json")
if os.path.exists(sample):
    datas.append((sample, "."))

# Collect package data files
try:
    import openpyxl
    datas.extend(collect_data_files("openpyxl"))
except ImportError:
    pass

# Hidden imports
hiddenimports = [
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "tkinter.scrolledtext",
    "src",
    "src.config",
    "src.config.manager",
    "src.core",
    "src.core.handler",
    "src.core.logic",
    "src.core.utils",
    "src.excel",
    "src.excel.loader",
    "src.excel.fill_details",
    "src.gui",
    "src.gui.app",
    "src.gui.constants",
    "src.gui.frames",
    "src.gui.frames.file_list",
    "src.gui.frames.aux_data",
    "src.gui.frames.processing",
    "PIL",
    "PIL.Image",
    "PIL.ImageTk",
] + collect_submodules("openpyxl")

# Analysis
a = Analysis(
    [main_script],
    pathex=[current_dir],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib",
        "scipy",
        "cv2",
        "sklearn",
        "tensorflow",
        "torch",
        "jupyter",
        "notebook",
        "IPython",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

# Package
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# Executable
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="PPLX_Handler",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_file if os.path.exists(icon_file) else None,
    version=None,
)
