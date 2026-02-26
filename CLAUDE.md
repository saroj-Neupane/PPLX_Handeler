# PPLX Handler

Python GUI application for batch editing PPLX (Pole Line Engineering XML) files with Excel integration. Manages electrical pole structure data, auxiliary data tagging, and make-ready work coordination.

## Project Structure

```
app.py                    # Entry point (imports src.main)
src/
  main.py                 # GUI initialization
  config/manager.py       # Config/state persistence, PPLXConfigManager
  core/
    handler.py            # PPLXHandler - XML parsing/manipulation
    logic.py              # Business logic (aux data, keyword matching)
    utils.py              # Shared utilities
  excel/
    loader.py             # Excel data loading (nodes sheet)
    fill_details.py       # PPLX fill details Excel export
  gui/
    app.py                # Main window (PPLXGUIApp), theme application
    constants.py          # THEME dict, AUX_AUTO_FILL_CONFIG
    frames/
      file_list.py        # PPLXFileListFrame (folder/ZIP selection)
      aux_data.py         # AuxDataEditFrame (8 aux fields + Excel)
      processing.py       # ProcessingFrame (batch processing, threading)
config/                   # Configuration profiles
  _active.json            # Active profile tracker
  OPPD.json               # OPPD config (keywords, power labels)
  state.json              # Session state (gitignored)
assets/handler.ico        # App icon
data/                     # Shapefiles (gitignored)
output/                   # Output log Excel files
```

## Architecture & Conventions

- **GUI Framework:** tkinter with custom purple/white theme (THEME dict in constants.py)
- **Project root resolution:** `_get_project_root()` in config/manager.py — uses `sys.executable` dir when frozen (PyInstaller), otherwise `Path(__file__).resolve().parents[2]`
- **Config system:** Dual-file — `state.json` for session data, named profiles (e.g. `OPPD.json`) for keywords/settings
- **Threading:** Batch processing runs in daemon threads via `ProcessingFrame`
- **Dependencies:** openpyxl, pandas, Pillow (see requirements.txt)
- **Build:** PyInstaller via handler.spec → `PPLX_Handler.exe`

## Key Patterns

- PPLX files are XML; parsed with `xml.etree.ElementTree`
- SCID extracted from filenames to match Excel data
- Aux Data 1-5 assigned via keyword matching against MR notes
- Categories: EXISTING and PROPOSED (processed separately)
- Output change logs saved to `output/` folder at project root
- Wire spec comparison uses shapefiles in `data/OPPD/shape/`

## Commands

```bash
python app.py              # Run the GUI
pip install -r requirements.txt  # Install dependencies
pyinstaller handler.spec   # Build executable
```
