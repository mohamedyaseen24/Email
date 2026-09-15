"""
split_by_unit_with_summary.py

GENERAL-PURPOSE tool. Works on any workbook that has:
  - Multiple sheets
  - A "Unit" column (name configurable) inside one or more of those sheets
  - Optionally, one or more "Summary"/"Sum" sheets

WHAT IT DOES
------------
1. Scans every sheet in the source workbook.
2. Sheets whose NAME matches the summary pattern (default: contains
   "summary" or is/contains "sum") are treated as summary sheets.
3. For every OTHER sheet, it looks for the unit column (default name:
   "Unit", case-insensitive, configurable). It collects every unique
   unit value found across all such sheets.
4. For each unit value, it builds one new .xlsx file containing:
     - Every non-summary sheet, filtered down to only that unit's rows
       (sheets that don't have the unit column are copied through
       unchanged, since there's no way to know which rows belong to
       which unit).
     - Every summary sheet: filtered by unit if the summary sheet also
       has the unit column, otherwise copied through unchanged.
5. If any summary sheet content applies to a unit, an email draft
   (.txt file: subject + body) is generated per unit under
   <output>/email_drafts/, with the summary content as the body.
   These are plain drafts to copy into Outlook/Gmail — the script
   does not send anything.

This is deliberately generic — column name, summary pattern, input
file and output folder are all CLI arguments, so it works on
different trackers without code changes.

USAGE
-----
    python split_by_unit_with_summary.py \\
        --input trackers.xlsx \\
        --unit-column "Unit" \\
        --output ./split_output

    # If your summary sheets are named something unusual:
    python split_by_unit_with_summary.py --input trackers.xlsx \\
        --summary-pattern "summary|recap|sum"
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd


def safe_filename(value: str) -> str:
    text = str(value).strip()
    text = re.sub(r'[\\/*?:"<>|]', "_", text)
    text = re.sub(r"\s+", "_", text)
    return text or "blank"


def find_unit_column(columns, unit_column_name: str):
    """Case-insensitive match for the unit column in a sheet's columns."""
    target = unit_column_name.strip().lower()
    for col in columns:
        if str(col).strip().lower() == target:
            return col
    return None


def is_summary_sheet(sheet_name: str, pattern: str) -> bool:
    return re.search(pattern, sheet_name, re.IGNORECASE) is not None


def load_workbook_sheets(input_path: Path):
    """Return dict of sheet_name -> DataFrame, preserving sheet order."""
    xls = pd.ExcelFile(input_path)
    return {name: xls.parse(name) for name in xls.sheet_names}


def collect_units(sheets: dict, unit_column_name: str, summary_pattern: str):
    units = set()
    for name, df in sheets.items():
        if is_summary_sheet(name, summary_pattern):
            continue
        col = find_unit_column(df.columns, unit_column_name)
        if col is not None:
            units.update(df[col].dropna().unique().tolist())
    return sorted(units, key=str)


def build_unit_workbook(unit_value, sheets: dict, unit_column_name: str,
                         summary_pattern: str, output_dir: Path):
    out_path = output_dir / f"{safe_filename(unit_value)}.xlsx"
    summary_texts = []

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        wrote_any_sheet = False
        for name, df in sheets.items():
            col = find_unit_column(df.columns, unit_column_name)

            if is_summary_sheet(name, summary_pattern):
                if col is not None:
                    filtered = df[df[col] == unit_value]
                else:
                    filtered = df  # no unit column on summary sheet -> copy through
                if not filtered.empty:
                    filtered.to_excel(writer, sheet_name=name[:31], index=False)
                    wrote_any_sheet = True
                    summary_texts.append(f"--- {name} ---\n{filtered.to_string(index=False)}")
                continue

            if col is not None:
                filtered = df[df[col] == unit_value]
                if filtered.empty:
                    continue  # this unit has no rows on this sheet -> skip it
                filtered.to_excel(writer, sheet_name=name[:31], index=False)
                wrote_any_sheet = True
            else:
                # No unit column on this sheet -> copy through unchanged
                df.to_excel(writer, sheet_name=name[:31], index=False)
                wrote_any_sheet = True

        if not wrote_any_sheet:
            # ExcelWriter can't save an empty workbook; add a placeholder.
            pd.DataFrame({"Note": [f"No data found for unit '{unit_value}'"]}).to_excel(
                writer, sheet_name="Note", index=False
            )

    return out_path, summary_texts


def write_email_draft(unit_value, summary_texts, email_dir: Path):
    if not summary_texts:
        return None
    email_dir.mkdir(parents=True, exist_ok=True)
    draft_path = email_dir / f"{safe_filename(unit_value)}_email.txt"
    subject = f"Subject: Summary - {unit_value}\n\n"
    body = (
        f"Hi,\n\nPlease find the summary for {unit_value} below.\n\n"
        + "\n\n".join(summary_texts)
        + f"\n\nRegards,\n"
    )
    draft_path.write_text(subject + body, encoding="utf-8")
    return draft_path


def main():
    parser = argparse.ArgumentParser(
        description="Split a multi-sheet workbook by unit and generate email drafts from summary sheets."
    )
    parser.add_argument("--input", required=True, help="Path to the source .xlsx file")
    parser.add_argument("--unit-column", default="Unit", help="Column name to split by (default: 'Unit')")
    parser.add_argument(
        "--summary-pattern",
        default=r"summary|sum",
        help="Regex to detect summary sheet names (default matches 'summary' or 'sum', case-insensitive)",
    )
    parser.add_argument("--output", default="./split_output", help="Folder to write output files to")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    email_dir = output_dir / "email_drafts"

    sheets = load_workbook_sheets(input_path)
    print(f"Loaded {len(sheets)} sheet(s): {', '.join(sheets.keys())}")

    units = collect_units(sheets, args.unit_column, args.summary_pattern)
    if not units:
        print(f"ERROR: no values found in a '{args.unit_column}' column on any non-summary sheet.")
        sys.exit(1)
    print(f"Found {len(units)} unit(s): {units}")

    for unit_value in units:
        out_path, summary_texts = build_unit_workbook(
            unit_value, sheets, args.unit_column, args.summary_pattern, output_dir
        )
        print(f"  wrote {out_path}")
        draft_path = write_email_draft(unit_value, summary_texts, email_dir)
        if draft_path:
            print(f"    + email draft: {draft_path}")

    print("Done.")


if __name__ == "__main__":
    main()
