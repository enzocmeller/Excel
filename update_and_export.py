#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automate the CSV update, refresh the Excel workbook, and export to PDF.
"""

import os
import sys
import time
import subprocess
from pathlib import Path

try:
    import win32com.client
except ImportError as exc:
    raise SystemExit("pywin32 is required. Install it with 'python -m pip install pywin32'.") from exc

BASE_DIR = Path(__file__).resolve().parent
CSV_SCRIPT = BASE_DIR / "usa_exports_monthly.py"
WORKBOOK = BASE_DIR / "USA_Exports.xlsx"
PDF_OUTPUT = BASE_DIR / "USA_Exports.pdf"


def run_csv_update():
    print("Updating CSV data...")
    result = subprocess.run([sys.executable, str(CSV_SCRIPT)], check=False)
    if result.returncode != 0:
        raise RuntimeError(f"CSV update failed with exit code {result.returncode}.")
    print("CSV update complete.")


def refresh_excel_workbook():
    if not WORKBOOK.exists():
        raise FileNotFoundError(f"Workbook not found: {WORKBOOK}")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False

    try:
        print(f"Opening workbook: {WORKBOOK.name}")
        wb = excel.Workbooks.Open(str(WORKBOOK), UpdateLinks=0, ReadOnly=False)
        try:
            print("Refreshing workbook queries and connections...")
            wb.RefreshAll()

            # Wait for background queries to finish.
            for attempt in range(120):
                status = getattr(excel, "BackgroundQueryStatus", 0)
                if status == 0:
                    break
                print(f"  waiting for query refresh... ({status} active)")
                time.sleep(1)
            else:
                raise RuntimeError("Excel query refresh did not finish within the timeout period.")

            if hasattr(excel, "CalculateUntilAsyncQueriesDone"):
                excel.CalculateUntilAsyncQueriesDone()

            print("Saving workbook...")
            wb.Save()

            print(f"Exporting workbook to PDF: {PDF_OUTPUT.name}")
            wb.ExportAsFixedFormat(
                Type=0,
                Filename=str(PDF_OUTPUT),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            print(f"PDF export complete: {PDF_OUTPUT}")
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


def main():
    try:
        run_csv_update()
        refresh_excel_workbook()
        print("All tasks complete.")
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
