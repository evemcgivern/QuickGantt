"""
Build script for creating an executable version of QuickGantt using PyInstaller.

This script first checks if PyInstaller is installed and attempts to install it
if missing. It then configures and runs PyInstaller to create a standalone executable.
"""

import os
import sys
import subprocess
import platform
from pathlib import Path

def check_pyinstaller() -> bool:
    """
    Check if PyInstaller is installed in the current environment.
    
    Returns:
        bool: True if PyInstaller is installed, False otherwise
    """
    try:
        import PyInstaller
        return True
    except ImportError:
        return False

def install_pyinstaller() -> bool:
    """
    Attempt to install PyInstaller using pip.
    
    Returns:
        bool: True if installation was successful, False otherwise
    """
    print("PyInstaller not found. Attempting to install...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        return True
    except subprocess.CalledProcessError:
        print("Failed to install PyInstaller. Please install manually:")
        print("pip install pyinstaller")
        return False

def build_app() -> None:
    """
    Build the QuickGantt application using PyInstaller.
    
    This creates a standalone executable with all necessary dependencies.
    """
    # Determine the appropriate icon file based on platform
    if platform.system() == "Windows":
        icon_file = "assets/icon.ico"  # Updated to match your actual file
    elif platform.system() == "Darwin":  # macOS
        icon_file = "assets/icon.icns"  # Updated to be consistent
    else:  # Linux
        icon_file = "assets/icon.png"   # Updated to be consistent
    
    # Create assets directory if it doesn't exist
    assets_dir = Path("assets")
    if not assets_dir.exists():
        assets_dir.mkdir()
        print("Created assets directory. Please add appropriate icon files.")
    
    # Default to no icon if the file doesn't exist
    icon_path = Path(icon_file)
    icon_param = []
    if icon_path.exists():
        print(f"Found icon: {icon_path.absolute()}")
        icon_param = ["--icon", str(icon_path)]
    else:
        print(f"Icon not found at: {icon_path.absolute()}")
        print("Building without an icon.")
    
    # Import PyInstaller and run
    from PyInstaller.__main__ import run

    # Define PyInstaller arguments
    pyinstaller_args = [
        "app.py",                         # Main script
        "--name", "QuickGantt",           # Name of the executable
        "--onefile",                      # Create a single file executable
        "--windowed",                     # Don't show console window on Windows/macOS
        "--clean",                        # Clean PyInstaller cache
        "--add-data", "color_selector.py;.",  # Include additional modules
        "--add-data", "chart_engine.py;.",
        *icon_param,                      # Add icon if available
        # Add hooks for required packages
        "--hidden-import", "matplotlib",
        "--hidden-import", "pandas",
        "--hidden-import", "tkinter",
        "--hidden-import", "openpyxl"
    ]
    
    # Change path separator for non-Windows platforms
    if platform.system() != "Windows":
        pyinstaller_args = [arg.replace(';', ':') if ';' in arg else arg for arg in pyinstaller_args]
    
    print(f"Building QuickGantt for {platform.system()}...")
    print(f"Command: {' '.join(pyinstaller_args)}")
    
    # Run PyInstaller with our arguments
    run(pyinstaller_args)
    

def main() -> None:
    """
    Main function to check dependencies and build the application.
    """
    if not check_pyinstaller():
        if not install_pyinstaller():
            sys.exit(1)
    
    # Import should work now
    build_app()

if __name__ == "__main__":
    main()