"""Entry point for PPLX GUI application."""


def main():
    """Run the GUI application."""
    from tkinter import messagebox
    from src.gui.app import PPLXGUIApp

    try:
        app = PPLXGUIApp()
        app.run()
    except Exception as e:
        print(f"Error starting application: {e}")
        messagebox.showerror("Startup Error", f"Failed to start application:\n{str(e)}")


def headless_main(args) -> int:
    """Run batch processing without the GUI. Returns exit code (0 = success)."""
    import os
    import glob
    from datetime import datetime
    from pathlib import Path

    from src.config.manager import PPLXConfigManager, _get_project_root
    from src.core.processor import process_single_file
    from src.core.utils import parse_keywords
    from src.excel.loader import load_excel_data
    from src.excel.changelog import write_change_log

    config_manager = PPLXConfigManager(config_name=args.config)

    # --- Collect PPLX files ---
    def find_pplx(folder):
        if not folder or not os.path.isdir(folder):
            return []
        return sorted(glob.glob(os.path.join(folder, "**/*.pplx"), recursive=True))

    existing_files = find_pplx(getattr(args, "existing", None))
    proposed_files = find_pplx(getattr(args, "proposed", None))

    if not existing_files and not proposed_files:
        print("ERROR: No PPLX files found. Provide --existing or --proposed folder.")
        return 1

    # --- Excel ---
    excel_path = getattr(args, "excel", None) or config_manager.get("excel_file_path", "")
    if not excel_path or not os.path.exists(excel_path):
        print(f"ERROR: Excel file not found: {excel_path!r}. Use --excel to specify.")
        return 1

    excel_data = load_excel_data(excel_path, log_callback=print)
    valid_scids = set(excel_data.keys()) if excel_data else set()

    # --- Settings from config/state ---
    keyword_payload = {
        "comm_keywords": parse_keywords(config_manager.get("comm_keywords", "")),
        "power_keywords": parse_keywords(config_manager.get("power_keywords", "")),
        "pco_keywords": parse_keywords(config_manager.get("pco_keywords", "")),
        "aux5_keywords": parse_keywords(config_manager.get("aux5_keywords", "")),
        "power_label": config_manager.get("power_label", "POWER"),
    }
    # --- Output directory ---
    timestamp = datetime.now().strftime("%Y_%m_%d_%H%M%S")
    if getattr(args, "output", None):
        output_root = args.output
    else:
        prefix = os.path.basename(excel_path).split(" ")[0].strip() or Path(excel_path).stem
        output_root = str(Path(_get_project_root()) / "output" / f"{prefix}_PPLX_{timestamp}")
    os.makedirs(output_root, exist_ok=True)
    print(f"Output root: {output_root}")

    # --- Shared processing args ---
    proc_kwargs = dict(
        excel_data=excel_data,
        valid_scids=valid_scids,
        auto_fill_aux1=config_manager.get("auto_fill_aux1", False),
        auto_fill_aux2=config_manager.get("auto_fill_aux2", False),
        keyword_payload=keyword_payload,
    )

    # --- Process categories ---
    for name, files in [("EXISTING", existing_files), ("PROPOSED", proposed_files)]:
        if not files:
            continue

        output_dir = os.path.join(output_root, name)
        os.makedirs(output_dir, exist_ok=True)
        successful = failed = skipped = 0
        csv_data = []

        print(f"\nCategory: {name} ({len(files)} file{'s' if len(files) != 1 else ''})")

        for file_path in files:
            result = process_single_file(
                file_path,
                condition_value=name,
                output_dir=output_dir,
                **proc_kwargs,
            )
            for entry in result["logs"]:
                print(f"  {entry}" if not entry.startswith("  ") else entry)
            if result["status"] == "success":
                successful += 1
                if result["csv_row"]:
                    csv_data.append(result["csv_row"])
            elif result["status"] == "skipped":
                skipped += 1
            else:
                failed += 1

        print(f"  Result - OK: {successful}, Skipped: {skipped}, Failed: {failed}")

        # Wire spec comparison (OPPD config, shapefiles present)
        wire_spec_data = []
        if config_manager.config_name == "OPPD":
            try:
                from src.core.wire_spec_from_excel import build_wire_spec_comparison
                from src.core.logic import extract_scid_from_filename
                shape_base = Path(_get_project_root()) / "data" / "OPPD" / "shape"
                if shape_base.exists():
                    wire_spec_data = build_wire_spec_comparison(
                        Path(excel_path), files, shape_base, extract_scid_from_filename,
                        log_callback=print,
                    )
                    print(f"  Wire spec rows: {len(wire_spec_data)}")
            except Exception as e:
                print(f"  Wire spec comparison failed: {e}")

        log_path = os.path.join(output_root, f"{name}_change_log_{timestamp}.xlsx")
        wire_spec_mapping = config_manager.get("wire_spec_mapping", {})
        try:
            if write_change_log(log_path, csv_data, wire_spec_data, wire_spec_mapping):
                print(f"  Change log: {log_path}")
        except Exception as e:
            print(f"  WARNING: Could not write change log: {e}")

    print(f"\nDone. Output: {output_root}")
    return 0


if __name__ == "__main__":
    main()
