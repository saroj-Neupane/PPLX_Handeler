#!/usr/bin/env python3
"""
Build script for creating PPLX GUI executable using PyInstaller.

This script automates the build process and provides helpful feedback.

Usage:
    python build_exe.py

Requirements:
    - PyInstaller: pip install pyinstaller
    - All project dependencies installed
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def check_pyinstaller():
    """Check if PyInstaller is installed."""
    try:
        import PyInstaller
        print(f"[OK] PyInstaller found: {PyInstaller.__version__}")
        return True
    except ImportError:
        print("[ERROR] PyInstaller not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("[OK] PyInstaller installed successfully")
            return True
        except subprocess.CalledProcessError:
            print("[ERROR] Failed to install PyInstaller")
            return False

def check_dependencies():
    """Check if required dependencies are available."""
    dependencies = ['tkinter', 'openpyxl', 'pandas']
    missing = []
    
    for dep in dependencies:
        try:
            __import__(dep)
            print(f"[OK] {dep} found")
        except ImportError:
            print(f"[ERROR] {dep} missing")
            missing.append(dep)
    
    if missing:
        print(f"\nMissing dependencies: {', '.join(missing)}")
        print("Install with: pip install " + " ".join(missing))
        return False
    
    return True

def clean_build_dirs():
    """Clean previous build directories."""
    dirs_to_clean = ['build', 'dist', '__pycache__']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"Cleaning {dir_name}/...")
            shutil.rmtree(dir_name)

def build_executable():
    """Build the executable using PyInstaller."""
    spec_file = "pplx_gui.spec"
    
    if not os.path.exists(spec_file):
        print(f"[ERROR] Spec file not found: {spec_file}")
        return False
    
    print(f"Building executable using {spec_file}...")
    
    try:
        # Run PyInstaller with the spec file
        cmd = [sys.executable, "-m", "PyInstaller", "--clean", spec_file]
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("[OK] Build completed successfully!")
            
            # Check if executable was created
            exe_name = "pplx_gui.exe" if os.name == 'nt' else "pplx_gui"
            exe_path = os.path.join("dist", exe_name)
            
            if os.path.exists(exe_path):
                size_mb = os.path.getsize(exe_path) / (1024 * 1024)
                print(f"[OK] Executable created: {exe_path} ({size_mb:.1f} MB)")
                return True
            else:
                print(f"[ERROR] Executable not found at expected location: {exe_path}")
                return False
        else:
            print("[ERROR] Build failed!")
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            return False
            
    except Exception as e:
        print(f"[ERROR] Build error: {e}")
        return False

def main():
    """Main build process."""
    print("PPLX GUI Build Script")
    print("=" * 50)
    
    # Check if we're in the right directory
    if not os.path.exists("pplx_gui.py"):
        print("[ERROR] Please run this script from the project root directory")
        sys.exit(1)
    
    # Check dependencies
    print("\nChecking dependencies...")
    if not check_pyinstaller():
        sys.exit(1)
    
    if not check_dependencies():
        sys.exit(1)
    
    # Clean previous builds
    print("\nCleaning previous builds...")
    clean_build_dirs()
    
    # Build executable
    print("\nBuilding executable...")
    if build_executable():
        print("\n[OK] Build completed successfully!")
        print("\nTo run the executable:")
        exe_name = "pplx_gui.exe" if os.name == 'nt' else "pplx_gui"
        print(f"  dist/{exe_name}")
    else:
        print("\n[ERROR] Build failed!")
        sys.exit(1)

if __name__ == "__main__":
    main()
