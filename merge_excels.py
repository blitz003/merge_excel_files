#!/usr/bin/env python3
"""
merge_excels.py
Merge all Excel files in a folder into one workbook.

Usage
-----
python merge_excels.py /path/to/folder               # merged.xlsx in current dir
python merge_excels.py /path/to/folder -o all.xlsx   # custom output name
python merge_excels.py /path/to/folder -s Sheet1     # sheet name instead of index
python merge_excels.py /path/to/folder -p "*.xls"    # different glob pattern

Requirements
------------
pip install pandas openpyxl
"""

from pathlib import Path
import pandas as pd
import sys
import argparse


def merge_excels(src_dir: Path,
                 output_file: Path = Path("merged.xlsx"),
                 pattern: str = "*.xlsx",
                 sheet_name=0):
    """
    Parameters
    ----------
    src_dir      : Path to folder containing Excel files
    output_file  : Destination workbook path
    pattern      : Glob pattern to match (e.g. '*.xlsx' or '*.xls')
    sheet_name   : Sheet index (0-based) or name to read from each file
    """
    files = sorted(src_dir.glob(pattern))
    if not files:
        raise FileNotFoundError(
            f"No Excel files found in {src_dir} matching pattern '{pattern}'")

    dfs = []
    for f in files:
        try:
            df = pd.read_excel(f, sheet_name=sheet_name)
            df["__source_file"] = f.name        # keep provenance (optional)
            dfs.append(df)
            print(f"✔ Read {f.name}")
        except Exception as exc:
            print(f"✘ Skipped {f.name}: {exc}", file=sys.stderr)

    merged = pd.concat(dfs, ignore_index=True)
    merged.to_excel(output_file, index=False)
    print(
        f"\nMerged {len(dfs)} workbooks → {output_file} "
        f"({len(merged):,} rows)."
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Merge Excel workbooks found in a folder.")
    parser.add_argument("src_dir", help="Directory containing Excel files")
    parser.add_argument(
        "-o", "--output", default="merged.xlsx",
        help="Output workbook path (default: merged.xlsx)")
    parser.add_argument(
        "-p", "--pattern", default="*.xlsx",
        help="Glob pattern to match (default: *.xlsx)")
    parser.add_argument(
        "-s", "--sheet", default=0,
        help="Sheet index or name to merge (default: first sheet)")
    args = parser.parse_args()

    merge_excels(Path(args.src_dir).expanduser().resolve(),
                 Path(args.output).expanduser().resolve(),
                 pattern=args.pattern,
                 sheet_name=args.sheet)
