#!/usr/bin/env python3
"""
Build script for creating eksel executable
"""
import os
import subprocess
import sys
from pathlib import Path

def build_exe():
    """Build the executable using PyInstaller"""
    print("Building eksel executable...")

    # Check if PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # Navigate to exe directory
    exe_dir = Path(__file__).parent
    os.chdir(exe_dir)

    # Build command using spec file with python -m to ensure venv is used
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "eksel.spec",
        "--clean"
    ]

    # Run PyInstaller
    try:
        subprocess.check_call(cmd)
        print("\nBuild successful!")
        print("Executable location: exe/dist/eksel.exe")
        print("Clean up build files with: python exe/build_exe.py --clean")
    except subprocess.CalledProcessError as e:
        print(f"Build failed: {e}")
        return False

    return True

def clean_build():
    """Clean up build artifacts"""
    import shutil

    # Navigate to exe directory
    exe_dir = Path(__file__).parent
    os.chdir(exe_dir)

    dirs_to_remove = ["build", "dist", "__pycache__"]

    print("Cleaning build artifacts...")

    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"Removed: {dir_name}/")

    print("✅ Cleanup complete!")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--clean":
        clean_build()
    else:
        if build_exe():
            print("\nTo distribute:")
            print("   - Share the exe/dist/eksel.exe file")
            print("   - No Python installation required on target machine")
            print("   - Excel must be installed on target machine")
            print("   - Assets folder is embedded in the executable")
