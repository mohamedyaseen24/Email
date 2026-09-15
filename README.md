# Split Workbook by Unit — Setup & Usage Guide

This tool takes one Excel workbook with multiple sheets, splits it into
separate files per "Unit," and generates email drafts from any
Summary sheets. This guide walks you through running it in VS Code,
even if you've never used it before.

---

## 1. Install Python

1. Go to [python.org/downloads](https://www.python.org/downloads/) and download Python 3.10 or newer.
2. Run the installer.
   - **Windows:** check the box that says **"Add python.exe to PATH"** before clicking Install.
   - **Mac:** the default installer settings are fine.
3. To confirm it worked, open a terminal (see step 3) and type:
   ```
   python --version
   ```
   You should see something like `Python 3.12.4`. If that command isn't found, try `py --version` instead — Windows sometimes uses `py`.

## 2. Install VS Code

1. Download it from [code.visualstudio.com](https://code.visualstudio.com/).
2. Install and open it.
3. Click the **Extensions** icon in the left sidebar (looks like four squares).
4. Search for **"Python"** and install the official one published by Microsoft.

## 3. Set up your project folder

1. Create a folder anywhere on your computer, e.g. `Desktop/unit-splitter`.
2. Put these two files in it:
   - `split_by_unit_with_summary.py` (the script)
   - your Excel file (e.g. `trackers.xlsx`)
3. In VS Code: **File → Open Folder…** and select that folder.
4. Open a terminal inside VS Code: **Terminal → New Terminal** (or press `` Ctrl+` `` on Windows, `` Cmd+` `` on Mac). It opens already pointed at your project folder — no need to navigate manually.

## 4. Install the required packages

In the terminal that just opened, run:

```
pip install pandas openpyxl
```

If you get a permissions error, run this instead:

```
pip install --user pandas openpyxl
```

You only need to do this once per computer (not every time you run the script).

## 5. Run the script

In the same terminal, run:

```
python split_by_unit_with_summary.py --input trackers.xlsx --unit-column "Unit" --output ./split_output
```

Replace:
- `trackers.xlsx` → your actual Excel filename
- `"Unit"` → the actual column name your data uses, if different
- `./split_output` → leave as-is, or change to any folder name you like

If `python` isn't recognized, use `py` instead:

```
py split_by_unit_with_summary.py --input trackers.xlsx --unit-column "Unit" --output ./split_output
```

## 6. Check your results

A new folder (`split_output` by default) appears in your project folder, visible in VS Code's file explorer on the left. Inside you'll find:

- One `.xlsx` file per unique unit value
- An `email_drafts` subfolder with a `.txt` file (subject + body) for every unit that had summary data — ready to copy into Outlook or Gmail

## Troubleshooting

| Problem | Fix |
|---|---|
| `'python' is not recognized` | Use `py` instead of `python` |
| `ModuleNotFoundError: No module named 'pandas'` | Re-run `pip install pandas openpyxl` |
| `ERROR: no values found in a 'Unit' column` | Check the exact column name in your file and pass it with `--unit-column "Your Column Name"` |
| Summary sheets aren't picked up | Pass a custom pattern, e.g. `--summary-pattern "summary|recap|sum"` |

## Re-running later

Every time you want to run it again (e.g. on an updated file), you only need step 5 — open the terminal in that folder and run the command again. Steps 1–4 are one-time setup.
