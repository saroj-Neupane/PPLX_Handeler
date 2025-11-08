# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for PPLX GUI Application
"""

import os
from PyInstaller.utils.hooks import collect_data_files

# Paths
current_dir = os.getcwd()
main_script = os.path.join(current_dir, 'pplx_gui.py')
icon_file = os.path.join(current_dir, 'handeler.ico')

# Data files to include
datas = []

# Configuration files
config_files = ['pplx_gui_config.json', 'pplx_structure_sample.json']
for config_file in config_files:
    config_path = os.path.join(current_dir, config_file)
    if os.path.exists(config_path):
        datas.append((config_path, '.'))

# Collect package data files
try:
    import openpyxl
    datas.extend(collect_data_files('openpyxl'))
except ImportError:
    pass

# Hidden imports
hiddenimports = [
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.scrolledtext',
    'pplx_handeler',
    'pplx_config',
    'openpyxl',
    'openpyxl.workbook',
    'openpyxl.worksheet',
    'openpyxl.utils',
]

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
        'matplotlib',
        'scipy',
        'PIL',
        'cv2',
        'sklearn',
        'tensorflow',
        'torch',
        'jupyter',
        'notebook',
        'IPython',
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
    name='PPLX_Handeler',
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
