#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Update the CSV from Census, refresh the workbook, and export the 8 commodity
sheets to ONE PDF (Soybean, Meal, Oil, Corn, Wheat, Pork, Chicken, Beef).

Why this is not just `wb.ExportAsFixedFormat(...)`:
  * A workbook-level export dumps EVERY sheet (incl. the big tExports data table),
    and each commodity sheet is saved at 100% scale with no "fit to page", so the
    print area is a hair larger than the page and spills onto 3 extra near-blank
    pages -> you get ~32 pages with big empty margins.
  * Instead we select ONLY the 8 report sheets, scale each one to fill exactly one
    A4 page, and export just that selection -> a clean 8-page A4 PDF.
None of this is saved back into the workbook (it is applied in memory and the
file is closed without saving), so the .xlsx itself is left untouched.
"""

import os
import sys
import time
import subprocess
from datetime import date
from pathlib import Path

try:
    import win32com.client
except ImportError as exc:
    raise SystemExit("pywin32 is required. Install it with 'python -m pip install pywin32'.") from exc

BASE_DIR = Path(__file__).resolve().parent
CSV_SCRIPT = BASE_DIR / "usa_exports_monthly.py"
WORKBOOK = BASE_DIR / "USA_Exports.xlsx"
PDF_OUTPUT = BASE_DIR / ("USA_Exports_Report_" + date.today().strftime("%Y_%m_%d") + ".pdf")

ORDER = ["Soybean", "Meal", "Oil", "Corn", "Wheat", "Pork", "Chicken", "Beef"]
PDF_PRINTER = "Microsoft Print to PDF"

# Excel enum constants
xlCalculationManual = -4135
xlCalculationAutomatic = -4105
xlTypePDF = 0

# A4 page size in points (used to compute the fill zoom)
A4_W, A4_H = 595.28, 841.89


# --------------------------------------------------------------------------- #
#  A4 paper: ExportAsFixedFormat uses the DEFAULT paper of "Microsoft Print to
#  PDF", not the sheet's PaperSize. On a US-defaulted Windows that is Letter, so
#  the report prints with wide side margins. We flip the printer's PER-USER
#  default to A4 just for the export (no admin needed) via a tiny PowerShell
#  call, and restore it afterwards. If anything is unavailable we just export on
#  whatever the default is.
# --------------------------------------------------------------------------- #
def _run_ps(script):
    return subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
                          capture_output=True, text=True)


def set_pdf_paper_a4():
    ps = (
        "Add-Type -AssemblyName System.Printing; Add-Type -AssemblyName ReachFramework;"
        "$q=(New-Object System.Printing.LocalPrintServer).GetPrintQueue('" + PDF_PRINTER + "');"
        "$tk=$q.UserPrintTicket; if(-not $tk){$tk=$q.DefaultPrintTicket};"
        "Write-Output $tk.PageMediaSize.PageMediaSizeName;"
        "$tk.PageMediaSize=New-Object System.Printing.PageMediaSize([System.Printing.PageMediaSizeName]::ISOA4);"
        "$q.UserPrintTicket=$tk;$q.Commit()"
    )
    try:
        out = _run_ps(ps)
        orig = (out.stdout or "").strip().splitlines()
        return orig[0] if orig else None
    except Exception:
        return None


def restore_pdf_paper(orig):
    if not orig:
        return
    ps = (
        "Add-Type -AssemblyName System.Printing; Add-Type -AssemblyName ReachFramework;"
        "$q=(New-Object System.Printing.LocalPrintServer).GetPrintQueue('" + PDF_PRINTER + "');"
        "$tk=$q.UserPrintTicket;"
        "$tk.PageMediaSize=New-Object System.Printing.PageMediaSize([System.Printing.PageMediaSizeName]'" + orig + "');"
        "$q.UserPrintTicket=$tk;$q.Commit()"
    )
    try:
        _run_ps(ps)
    except Exception:
        pass


# --------------------------------------------------------------------------- #
def run_csv_update():
    print("[1/3] Updating Census data...")
    result = subprocess.run([sys.executable, str(CSV_SCRIPT)], check=False)
    if result.returncode != 0:
        raise RuntimeError(f"CSV update failed with exit code {result.returncode}.")
    print("      CSV update complete.")


def refresh_and_export():
    if not WORKBOOK.exists():
        raise FileNotFoundError(f"Workbook not found: {WORKBOOK}")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AskToUpdateLinks = False
    except Exception:
        pass

    orig_paper = None
    wb = None
    try:
        print(f"[2/3] Opening and refreshing {WORKBOOK.name} ...")
        wb = excel.Workbooks.Open(str(WORKBOOK), UpdateLinks=0, ReadOnly=False)
        # Don't let OneDrive AutoSave persist the temporary export tweaks.
        try:
            wb.AutoSaveOn = False
        except Exception:
            pass
        # Manual calc so the IMAGE() flag formulas don't trigger a recalc storm
        # during the refresh.
        try:
            excel.Calculation = xlCalculationManual
        except Exception:
            pass

        wb.RefreshAll()
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        # One deliberate recalc, then wait until Excel is truly idle.
        try:
            excel.CalculateFull()
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        deadline = time.time() + 300
        while time.time() < deadline:
            try:
                if excel.CalculationState == 0:   # xlDone
                    break
            except Exception:
                break
            time.sleep(0.5)

        # Clear the dashed page-break preview lines before saving, then save the
        # refreshed data.
        for ws in wb.Worksheets:
            try:
                ws.DisplayPageBreaks = False
            except Exception:
                pass
        wb.Save()
        print("      data refreshed and saved.")

        print("[3/3] Building the 8-page PDF...")
        # Switch to automatic calc for the PDF stage: reordering tabs under manual
        # calc smears the table borders into the numbers; automatic repaints cleanly.
        try:
            excel.Calculation = xlCalculationAutomatic
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        # A4 paper for the export.
        orig_paper = set_pdf_paper_a4()

        # Order the tabs Soybean..Beef by moving each to the front in reverse.
        for i in range(len(ORDER) - 1, -1, -1):
            for _ in range(10):
                try:
                    wb.Worksheets(ORDER[i]).Move(Before=wb.Worksheets(1))
                    break
                except Exception:
                    time.sleep(0.2)

        # Scale each sheet so its print area fills exactly one A4 page. We set a
        # fixed Zoom (a clean operation) computed from the print-area size vs the
        # usable A4 area - NOT FitToPages, which never enlarges past 100%.
        for name in ORDER:
            try:
                ws = wb.Worksheets(name)
                ps = ws.PageSetup
                pa = ps.PrintArea
                if not pa:
                    continue
                rng = ws.Range(pa)
                cw, ch = float(rng.Width), float(rng.Height)
                if cw <= 0 or ch <= 0:
                    continue
                uw = A4_W - ps.LeftMargin - ps.RightMargin
                uh = A4_H - ps.TopMargin - ps.BottomMargin
                z = int(min(uw / cw, uh / ch) * 100)   # floor via int()
                z = max(10, min(400, z))
                ps.Zoom = z
            except Exception as exc:
                print(f"      (fill-zoom skipped for {name}: {exc})")

        # Group-select the 8 sheets so one export call emits all of them. These
        # selects must be DIRECT (a retry wrapper breaks 'extend selection'); we
        # verify the count and redo it if it didn't take.
        sel_ok = False
        for _ in range(12):
            try:
                wb.Worksheets(ORDER[0]).Select(True)
                for i in range(1, len(ORDER)):
                    wb.Worksheets(ORDER[i]).Select(False)
                if excel.ActiveWindow.SelectedSheets.Count == len(ORDER):
                    sel_ok = True
                    break
            except Exception:
                pass
            time.sleep(0.3)
        n_sel = 0
        try:
            n_sel = excel.ActiveWindow.SelectedSheets.Count
        except Exception:
            pass
        print(f"      grouped {n_sel} of {len(ORDER)} sheets.")

        wb.ActiveSheet.ExportAsFixedFormat(
            Type=xlTypePDF,
            Filename=str(PDF_OUTPUT),
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
        print(f"      PDF written: {PDF_OUTPUT}")
    finally:
        # Close WITHOUT saving so the export-only tweaks (zoom, tab order) are
        # discarded; the refreshed data was already saved above.
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()
        restore_pdf_paper(orig_paper)


def main():
    try:
        run_csv_update()
        refresh_and_export()
        print("Done.")
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
