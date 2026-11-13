"""
00_setup_check.py  --  RUN THIS FIRST (click the Run Python File button in VS Code)

Checks that pandas and openpyxl are installed and installs them if they are missing,
then checks that the sample data files are where the scripts expect them.
No typing in a terminal needed.
"""
import importlib.util
import subprocess
import sys
from pathlib import Path

REQUIRED = {"pandas": "pandas", "openpyxl": "openpyxl"}  # import name -> pip name
DATA = Path(__file__).resolve().parent.parent / "data"


def main():
    print(f"Python {sys.version.split()[0]}")
    print(f"Running from: {sys.executable}")

    missing = [pip for mod, pip in REQUIRED.items() if importlib.util.find_spec(mod) is None]
    if missing:
        print(f"Installing missing libraries: {', '.join(missing)} ...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
        except subprocess.CalledProcessError:
            print("\nCould not install automatically.")
            print("In VS Code: press Ctrl+Shift+P (Cmd+Shift+P on Mac), run")
            print("'Python: Create Environment', choose Venv, tick requirements.txt,")
            print("then run this script again.")
            sys.exit(1)

    import openpyxl
    import pandas

    print(f"pandas {pandas.__version__} and openpyxl {openpyxl.__version__} are ready.")

    files = sorted(DATA.glob("*.xlsx")) if DATA.exists() else []
    if len(files) >= 6:
        print(f"Found {len(files)} sample files in the data folder. You are all set!")
    else:
        print("\nThe data folder looks incomplete.")
        print("Make sure you opened the whole vibe-coding-kit folder in VS Code")
        print("(File > Open Folder), and that the data folder has the .xlsx files.")


if __name__ == "__main__":
    main()
