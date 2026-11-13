# Vibe Coding Kit: Python & Excel (VS Code edition)

## Set up (once)

1. Install [VS Code](https://code.visualstudio.com/) and [Python](https://www.python.org/downloads/).
2. Create a folder called `VibeCodingWorkshop`. In VS Code choose **File > Open Folder** and open it.
3. Optional: open the Extensions panel (Ctrl+Shift+X, or Cmd+Shift+X on Mac), search for **Python** (by Microsoft) and click Install. This adds a **Run Python File** button to the editor.
4. In the Explorer panel, right-click and choose **New Folder** to create these three folders inside `VibeCodingWorkshop`:

```
VibeCodingWorkshop/
  data/       the Excel files (.xlsx)
  scripts/    the Python files (.py)
  output/     leave empty; results are saved here
```

5. Go to https://github.com/charanyaram/VibeCodingDemo and download the files. Put the Python files (.py) in `scripts/` and the Excel files (.xlsx) in `data/`.
6. Open `scripts/00_setup_check.py` and run it. It checks that Python is installed, installs pandas and openpyxl if they are missing, and confirms the data files are in the right place.

To run a script, click the **Run Python File** button (the triangle at the top right of the editor). If you skipped step 3, open a terminal with **Terminal > New Terminal** and type `python scripts/00_setup_check.py` (use `python3` on Mac).

If you are asked to select an interpreter, choose a recent Python 3. If the Run button is missing, press Ctrl+Shift+P (Cmd+Shift+P on Mac) and run **Python: Select Interpreter**.

## Notes

- The sample data is fake and was generated with a fixed random seed, so everyone gets the same results.
- Week 3 of the attendance data uses different column names (`Student`, `Status`) on purpose.
- The scripts were tested with Python 3.13, pandas 3.0 and openpyxl 3.1.
