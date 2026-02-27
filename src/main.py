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
    import time
    from datetime import datetime
    from pathlib import Path
    from concurrent.futures import ThreadPoolExecutor
    from functools import partial

    t0 = time.perf_counter()
    from src.config.manager import PPLXConfigManager, _get_project_root
    from src.core.processor import process_single_file
    from src.core.utils import parse_keywords
    from src.excel.loader import load_excel_data
    from src.excel.changelog import write_change_log
    t_imports = time.perf_counter() - t0

    t0 = time.perf_counter()
    config_manager = PPLXConfigManager(config_name=args.config)
    t_config = time.perf_counter() - t0

    # OPPD: eager-import heavy module once (numpy, pyproj, etc.) so cost is paid once, not in loop
    t_oppd_import_once = 0.0
    if getattr(config_manager, "config_name", None) == "OPPD":
        t0 = time.perf_counter()
        import src.core.wire_spec_from_excel  # noqa: F401
        from src.core.logic import extract_scid_from_filename  # noqa: F401
        t_oppd_import_once = time.perf_counter() - t0

    # --- Collect PPLX files ---
    t0 = time.perf_counter()
    def find_pplx(folder):
        if not folder or not os.path.isdir(folder):
            return []
        return sorted(glob.glob(os.path.join(folder, "**/*.pplx"), recursive=True))

    existing_files = find_pplx(getattr(args, "existing", None))
    proposed_files = find_pplx(getattr(args, "proposed", None))
    t_find_files = time.perf_counter() - t0

    if not existing_files and not proposed_files:
        print("ERROR: No PPLX files found. Provide --existing or --proposed folder.")
        return 1

    # --- Excel ---
    excel_path = getattr(args, "excel", None) or config_manager.get("excel_file_path", "")
    if not excel_path or not os.path.exists(excel_path):
        print(f"ERROR: Excel file not found: {excel_path!r}. Use --excel to specify.")
        return 1

    t0 = time.perf_counter()
    excel_data = load_excel_data(excel_path, log_callback=print)
    t_excel = time.perf_counter() - t0
    valid_scids = set(excel_data.keys()) if excel_data else set()

    # --- Settings from config/state ---
    t0 = time.perf_counter()
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
    t_setup = time.perf_counter() - t0

    timing_breakdown = [
        ("headless_main imports", t_imports),
        ("config load", t_config),
        ("OPPD module import (numpy, pyproj, etc.)", t_oppd_import_once),
        ("find PPLX files", t_find_files),
        ("Excel load", t_excel),
        ("setup (keywords, output dir)", t_setup),
    ]

    # --- Process categories ---
    for name, files in [("EXISTING", existing_files), ("PROPOSED", proposed_files)]:
        if not files:
            continue

        output_dir = os.path.join(output_root, name)
        os.makedirs(output_dir, exist_ok=True)
        successful = failed = skipped = 0
        csv_data = []

        print(f"\nCategory: {name} ({len(files)} file{'s' if len(files) != 1 else ''})")

        max_workers = min(8, os.cpu_count() or 4)
        category_start = time.perf_counter()

        process_fn = partial(
            process_single_file,
            condition_value=name,
            output_dir=output_dir,
            **proc_kwargs,
        )

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            for file_path, result in zip(files, executor.map(process_fn, files)):
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

        category_elapsed = time.perf_counter() - category_start
        timing_breakdown.append((f"{name} process files", category_elapsed))
        print(
            f"  Result - OK: {successful}, Skipped: {skipped}, Failed: {failed} "
            f"(in {category_elapsed:.1f}s using {max_workers} workers)"
        )

        # Wire spec and spans comparison (OPPD config)
        wire_spec_data = []
        spans_data = []
        t_oppd_import = t_pplx_preload = t_wire_spec = t_spans = t_changelog = 0.0
        if config_manager.config_name == "OPPD":
            try:
                t0 = time.perf_counter()
                from src.core.wire_spec_from_excel import (
                    build_wire_spec_comparison, build_spans_comparison_data,
                    load_all_excel_data,
                )
                t_oppd_import = time.perf_counter() - t0
                timing_breakdown.append((f"{name} wire_spec_from_excel import (cached)", t_oppd_import))
                # Preload all PPLX handlers once (shared cache for wire_spec + spans)
                from src.core.handler import PPLXHandler
                def _load_one_pplx(path):
                    try:
                        return (path, PPLXHandler(path))
                    except Exception:
                        return (path, None)
                t0 = time.perf_counter()
                shared_pplx_cache = {}
                with ThreadPoolExecutor(max_workers=max_workers) as ex:
                    for fp, handler in ex.map(_load_one_pplx, files):
                        shared_pplx_cache[fp] = handler
                t_pplx_preload = time.perf_counter() - t0
                timing_breakdown.append((f"{name} PPLX preload (shared cache)", t_pplx_preload))
                if shared_pplx_cache:
                    print(f"  PPLX preloaded: {sum(1 for h in shared_pplx_cache.values() if h is not None)}/{len(shared_pplx_cache)} handlers")

                # Pre-load Excel data once (single workbook open) and share between wire_spec and spans builds
                t0 = time.perf_counter()
                shared_nodes, shared_conns, shared_sections = load_all_excel_data(Path(excel_path))
                t_excel_shared = time.perf_counter() - t0
                timing_breakdown.append((f"{name} shared Excel load (nodes+conns+sections)", t_excel_shared))

                shape_base = Path(_get_project_root()) / "data" / "OPPD" / "shape"
                t0 = time.perf_counter()
                if shape_base.exists():
                    wire_spec_data = build_wire_spec_comparison(
                        Path(excel_path), files, shape_base, extract_scid_from_filename,
                        log_callback=print,
                        pplx_cache=shared_pplx_cache,
                        nodes=shared_nodes,
                        conns=shared_conns,
                    )
                    print(f"  Wire spec rows: {len(wire_spec_data)}")
                t_wire_spec = time.perf_counter() - t0
                t0 = time.perf_counter()
                span_mapping = config_manager.get("span_type_mapping", {})
                midspan_path = config_manager.get("midspan_heights_file_path", "")
                spans_data = build_spans_comparison_data(
                    Path(excel_path), files, extract_scid_from_filename,
                    log_callback=print,
                    span_type_mapping=span_mapping,
                    pplx_cache=shared_pplx_cache,
                    midspan_heights_path=Path(midspan_path) if midspan_path else None,
                    nodes=shared_nodes,
                    conns=shared_conns,
                    sections=shared_sections,
                )
                print(f"  Spans comparison rows: {len(spans_data)}")
                t_spans = time.perf_counter() - t0
            except Exception as e:
                print(f"  Wire spec/spans comparison failed: {e}")
        timing_breakdown.append((f"{name} wire_spec build", t_wire_spec))
        timing_breakdown.append((f"{name} spans build", t_spans))

        log_path = os.path.join(output_root, f"{name}_change_log_{timestamp}.xlsx")
        wire_spec_mapping = config_manager.get("wire_spec_mapping", {})
        t0 = time.perf_counter()
        try:
            if write_change_log(log_path, csv_data, wire_spec_data, wire_spec_mapping, spans_data=spans_data):
                print(f"  Change log: {log_path}")
        except Exception as e:
            print(f"  WARNING: Could not write change log: {e}")
        t_changelog = time.perf_counter() - t0
        timing_breakdown.append((f"{name} write_change_log", t_changelog))

    total_inner = sum(t for _, t in timing_breakdown)
    print(f"\nDone. Output: {output_root}")
    print("\n[TIMING] Breakdown (root cause):")
    for label, sec in timing_breakdown:
        pct = (100 * sec / total_inner) if total_inner else 0
        print(f"  {label}: {sec:.2f}s ({pct:.0f}%)")
    slow = [(l, s) for l, s in timing_breakdown if s >= 1.0]
    if slow:
        top = max(slow, key=lambda x: x[1])
        print(f"\n  ROOT CAUSE: '{top[0]}' dominates ({top[1]:.1f}s).")
    return 0


if __name__ == "__main__":
    main()
