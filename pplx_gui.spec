# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for PPLX GUI Application

This spec file creates a standalone executable for the PPLX File Editor GUI.
It includes all necessary dependencies and data files.

Usage:
    pyinstaller pplx_gui.spec

Output:
    dist/pplx_gui.exe (Windows)
    dist/pplx_gui (Linux/macOS)
"""

import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Get the current directory
current_dir = os.getcwd()

# Define the main script
main_script = os.path.join(current_dir, 'pplx_gui.py')

# Collect all data files that need to be included
datas = []

# Add configuration files if they exist
config_files = [
    'pplx_gui_config.json',
    'pplx_structure_sample.json',
    'requirements.txt'
]

for config_file in config_files:
    config_path = os.path.join(current_dir, config_file)
    if os.path.exists(config_path):
        # Include config files in the root of the executable directory
        datas.append((config_path, '.'))

# Collect openpyxl data files if available
try:
    import openpyxl
    openpyxl_datas = collect_data_files('openpyxl')
    datas.extend(openpyxl_datas)
except ImportError:
    pass

# Collect pandas data files if available
try:
    import pandas
    pandas_datas = collect_data_files('pandas')
    datas.extend(pandas_datas)
except ImportError:
    pass

# Hidden imports for modules that might not be detected automatically
hiddenimports = [
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.scrolledtext',
    'xml.etree.ElementTree',
    'json',
    'csv',
    'threading',
    'datetime',
    'pathlib',
    'pplx_handeler',
    'pplx_config',
    'openpyxl',
    'pandas',
    'openpyxl.workbook',
    'openpyxl.worksheet',
    'openpyxl.utils',
    'pandas.io.excel',
    'pandas.io.common'
]

# Analysis configuration
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
        # Exclude modules we don't need to reduce size
        'matplotlib',
        'numpy.random._pickle',
        'scipy',
        'PIL',
        'cv2',
        'sklearn',
        'tensorflow',
        'torch',
        'jupyter',
        'notebook',
        'IPython'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

# Remove duplicate binaries and data files
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# Create the executable
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='pplx_gui',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to False for GUI application (no console window)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # You can add an icon file here if you have one
    version=None,  # You can add a version file here if you have one
)

# Optional: Create a directory distribution instead of single file
# Uncomment the following lines if you prefer a directory distribution
# which is faster to start but includes more files

# exe = EXE(
#     pyz,
#     a.scripts,
#     [],
#     exclude_binaries=True,
#     name='pplx_gui',
#     debug=False,
#     bootloader_ignore_signals=False,
#     strip=False,
#     upx=True,
#     console=False,
#     disable_windowed_traceback=False,
#     argv_emulation=False,
#     target_arch=None,
#     codesign_identity=None,
#     entitlements_file=None,
#     icon=None,
# )
# 
# coll = COLLECT(
#     exe,
#     a.binaries,
#     a.zipfiles,
#     a.datas,
#     strip=False,
#     upx=True,
#     upx_exclude=[],
#     name='pplx_gui',
# )
