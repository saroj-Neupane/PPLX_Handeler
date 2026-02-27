#!/usr/bin/env python3
"""
PPLX File Editor - Entry point.

Run: python app.py
Run headless: python app.py --headless --existing <folder> --proposed <folder> --excel <file>
"""

if __name__ == "__main__":
    import argparse
    import sys
    import time as _time

    _t_process_start = _time.perf_counter()

    parser = argparse.ArgumentParser(description="PPLX Handler")
    parser.add_argument("--headless", action="store_true", help="Run without GUI")
    parser.add_argument("--existing", help="Path to EXISTING PPLX folder")
    parser.add_argument("--proposed", help="Path to PROPOSED PPLX folder")
    parser.add_argument("--excel", help="Path to Excel file (overrides saved state)")
    parser.add_argument("--output", help="Output directory (default: ~/Downloads/Processed PPLX/)")
    parser.add_argument("--config", help="Config profile name (default: active config)")
    args = parser.parse_args()

    if args.headless:
        _t_before_import = _time.perf_counter()
        from src.main import headless_main
        _t_after_import = _time.perf_counter()
        exit_code = headless_main(args)
        _t_after_headless = _time.perf_counter()
        print(
            f"\n[TIMING] process start->import: {_t_before_import - _t_process_start:.2f}s | "
            f"import headless_main: {_t_after_import - _t_before_import:.2f}s | "
            f"headless_main(): {_t_after_headless - _t_after_import:.2f}s | "
            f"total: {_t_after_headless - _t_process_start:.2f}s"
        )
        sys.exit(exit_code)
    else:
        from src.main import main
        main()
