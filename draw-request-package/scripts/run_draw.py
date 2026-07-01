#!/usr/bin/env python3
"""Orchestrate the non-vision steps of a draw once invoices.json exists.

Given the tracker workbook, an invoices.json, and the invoices folder, this:
  1. creates the new 'Draw Request N' sheet if it isn't there yet,
  2. applies the invoices (Current Period + retainage, Exhibit D),
  3. builds the cover-sheet .docx,
  4. assembles the cover + Exhibit D + invoice PDFs into the package PDF.

The one step this does NOT do is read the scanned invoices — that is the agent's
job (see SKILL.md): the agent produces invoices.json from the rendered images,
then calls this.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import draw_lib as dl
import openpyxl


def _run(script, *args):
    cmd = [sys.executable, str(HERE / script), *map(str, args)]
    print(f"$ {' '.join(cmd)}")
    subprocess.run(cmd, check=True)


def main():
    ap = argparse.ArgumentParser(description="Run the mechanical steps of a draw from invoices.json.")
    ap.add_argument("workbook", help="Tracker .xlsx (edited in place)")
    ap.add_argument("invoices_json")
    ap.add_argument("invoices_dir", help="Folder of invoice PDFs")
    ap.add_argument("--out-dir", default=".", help="Where to write cover + package")
    args = ap.parse_args()

    data = dl.load_invoices(args.invoices_json)
    draw_number = data.get("draw_number")

    # Create the sheet if missing.
    wb = openpyxl.load_workbook(args.workbook)
    have = any(dl._norm(s) == dl._norm(f"draw request {draw_number}") for s in wb.sheetnames)
    _, latest = dl.latest_draw_sheet(wb)
    if not have:
        if draw_number != latest + 1:
            print(f"NOTE: latest sheet is Draw {latest}; requested Draw {draw_number}.")
        _run("new_draw_sheet.py", args.workbook)

    _run("apply_invoices.py", args.workbook, args.invoices_json, "--draw", draw_number)

    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    cover = out / f"Cover_Sheet_{draw_number}.docx"
    package = out / f"Draw_Request_{draw_number}_package.pdf"
    _run("build_cover.py", args.invoices_json, cover)
    _run("assemble_package.py", args.invoices_json, args.invoices_dir, package)

    print("\nDone.")
    print(f"  Tracker updated : {args.workbook}")
    print(f"  Cover sheet     : {cover}")
    print(f"  Package PDF     : {package}")


if __name__ == "__main__":
    main()
