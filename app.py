#!/usr/bin/env python3
"""
PPLX File Editor - Entry point.

Run: python app.py
Run headless: python app.py --headless --existing <folder> --proposed <folder> --excel <file>
"""

if __name__ == "__main__":
    import argparse
    import sys

    parser = argparse.ArgumentParser(description="PPLX Handler")
    parser.add_argument("--headless", action="store_true", help="Run without GUI")
    parser.add_argument("--existing", help="Path to EXISTING PPLX folder")
    parser.add_argument("--proposed", help="Path to PROPOSED PPLX folder")
    parser.add_argument("--excel", help="Path to Excel file (overrides saved state)")
    parser.add_argument("--output", help="Output directory (default: ~/Downloads/Processed PPLX/)")
    parser.add_argument("--config", help="Config profile name (default: active config)")
    args = parser.parse_args()

    if args.headless:
        from src.main import headless_main
        sys.exit(headless_main(args))
    else:
        from src.main import main
        main()
