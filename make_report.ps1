# =============================================================================
# make_report.ps1
#   1) updates the CSV from Census (runs usa_exports_monthly.py)
#   2) opens USA_Exports.xlsx, refreshes all Power Query connections, saves
#   3) exports the 8 commodity sheets to ONE PDF, in this exact order:
#        Soybean, Meal, Oil, Corn, Wheat, Pork, Chicken, Beef
# Run it from run_update.bat (double-click).
#
# Requirements: Excel installed; USA_Exports.xlsx in this same folder with sheet
# tabs named exactly Soybean / Meal / Oil / Corn / Wheat / Pork / Chicken / Beef.
# IMPORTANT: close USA_Exports.xlsx in Excel before running (so it isn't locked).
# =============================================================================

$ErrorActionPreference = "Stop"
$folder = $PSScriptRoot
$xlsx   = Join-Path $folder "USA_Exports.xlsx"
$pyfile = Join-Path $folder "usa_exports_monthly.py"
$pdf    = Join-Path $folder ("USA_Exports_Report_" + (Get-Date -Format "yyyy_MM_dd") + ".pdf")
$order  = @("Soybean","Meal","Oil","Corn","Wheat","Pork","Chicken","Beef")

# Excel COM occasionally answers "the message filter indicated the application
# is busy" (RPC_E_CALL_REJECTED) when it receives many property writes in a row,
# as the page-setup loop does. This retries a COM call a few times before giving
# up, which makes the script reliable instead of crashing mid-run.
function Try-COM([scriptblock]$b) {
    for ($i = 0; $i -lt 50; $i++) {
        try { return (& $b) } catch { Start-Sleep -Milliseconds 150 }
    }
    & $b   # last attempt: let a real failure surface
}

# ExportAsFixedFormat produces a PDF on the DEFAULT paper of the "Microsoft Print
# to PDF" printer, NOT the sheet's PageSetup.PaperSize. On a US-defaulted Windows
# that's Letter, so the A4-shaped report prints with wide side margins. We flip
# the printer's PER-USER default to A4 just for the export (no admin needed) and
# restore it in the finally block. If anything here is unavailable, we silently
# fall back to the existing default and still produce a valid (Letter) PDF.
$script:pdfPrinter = "Microsoft Print to PDF"
$script:origPaper  = $null
function Set-PdfPaperA4 {
    $script:origPaper = $null
    try {
        Add-Type -AssemblyName System.Printing -ErrorAction Stop
        Add-Type -AssemblyName ReachFramework -ErrorAction Stop
        $srv = New-Object System.Printing.LocalPrintServer
        $q   = $srv.GetPrintQueue($script:pdfPrinter)
        $tk  = $q.UserPrintTicket; if (-not $tk) { $tk = $q.DefaultPrintTicket }
        try { $script:origPaper = $tk.PageMediaSize.PageMediaSizeName } catch {}
        $tk.PageMediaSize = New-Object System.Printing.PageMediaSize ([System.Printing.PageMediaSizeName]::ISOA4)
        $q.UserPrintTicket = $tk; $q.Commit()
        return $true
    } catch {
        Write-Host "    (could not switch the PDF printer to A4: $($_.Exception.Message))"
        return $false
    }
}
function Restore-PdfPaper {
    if (-not $script:origPaper) { return }
    try {
        Add-Type -AssemblyName System.Printing -ErrorAction Stop
        Add-Type -AssemblyName ReachFramework -ErrorAction Stop
        $srv = New-Object System.Printing.LocalPrintServer
        $q   = $srv.GetPrintQueue($script:pdfPrinter)
        $tk  = $q.UserPrintTicket
        $tk.PageMediaSize = New-Object System.Printing.PageMediaSize ([System.Printing.PageMediaSizeName]$script:origPaper)
        $q.UserPrintTicket = $tk; $q.Commit()
    } catch {}
}

# ---- 1) update the CSV -------------------------------------------------------
Write-Host "[1/3] Updating Census data..."
$pyCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pyCmd) { $pyCmd = Get-Command py -ErrorAction SilentlyContinue }
if (-not $pyCmd) { Write-Host "  Python not found. Install Python 3 and try again."; exit 1 }
& $pyCmd.Source $pyfile
if ($LASTEXITCODE -ne 0) { Write-Host "  Data update failed - stopping."; exit 1 }

if (-not (Test-Path $xlsx)) {
    Write-Host "  USA_Exports.xlsx not found in this folder - data updated, skipping refresh/PDF."
    exit 0
}

# ---- 2) refresh queries, 3) export PDF --------------------------------------
$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AskToUpdateLinks = $false } catch {}
    $wb = $excel.Workbooks.Open($xlsx)

    # CRITICAL: switch off automatic recalculation. The commodity sheets use
    # IMAGE() formulas for the country flags, and every recalc re-downloads all
    # of them. Left on "automatic", refreshing the data triggers a recalc storm
    # that keeps Excel busy for many minutes - long enough that the later page
    # setup / export calls get "Call was rejected by callee" (RPC_E_CALL_REJECTED)
    # and the whole run dies. We recalc ONCE, on purpose, further down.
    try { $excel.Calculation = -4135 } catch {}   # xlCalculationManual

    Write-Host "[2/3] Refreshing queries (this can take a moment)..."
    # Force foreground refresh so each .Refresh() BLOCKS until it has actually
    # finished. Power Query ignores RefreshAll() timing, which is why a plain
    # RefreshAll + sleep often saved stale data. Refreshing each connection one
    # at a time and waiting is the reliable way.
    foreach ($c in $wb.Connections) {
        try { $c.OLEDBConnection.BackgroundQuery = $false } catch {}
        try { $c.ODBCConnection.BackgroundQuery  = $false } catch {}
    }
    $refreshed = 0
    foreach ($c in $wb.Connections) {
        try {
            $c.Refresh()
            $refreshed++
            Write-Host "    refreshed: $($c.Name)"
        } catch {
            Write-Host "    WARNING - could not refresh '$($c.Name)': $($_.Exception.Message)"
        }
    }
    if ($refreshed -eq 0) {
        Write-Host "    no named connections found - falling back to RefreshAll"
        try { $wb.RefreshAll(); $excel.CalculateUntilAsyncQueriesDone() } catch {}
    }
    # Now recalc the workbook ONCE with the fresh data (this is the single,
    # intentional flag re-download), then WAIT until Excel is fully idle before
    # we touch anything else. Doing page setup / export while Excel is still
    # mid-recalc is exactly what threw RPC_E_CALL_REJECTED before.
    Write-Host "    recalculating (one pass - flag images reload here)..."
    try { $excel.CalculateFull() } catch {}
    try { $excel.CalculateUntilAsyncQueriesDone() } catch {}
    # CalculationState: 0=xlDone, 1=xlCalculating, 2=xlPending. Wait up to 5 min.
    $deadline = (Get-Date).AddMinutes(5)
    while ((Get-Date) -lt $deadline) {
        $state = 1
        try { $state = $excel.CalculationState } catch {}
        if ($state -eq 0) { break }
        Start-Sleep -Milliseconds 500
    }
    try { $excel.CalculateUntilAsyncQueriesDone() } catch {}
    $wb.Save()                        # persist the refreshed data
    Write-Host "    queries refreshed, recalculated, and workbook saved."

    # Data is fully calculated and saved. Switch recalc back to AUTOMATIC for the
    # PDF stage: reordering the tabs (Move, below) under MANUAL calc leaves the
    # table borders mis-painted - stray horizontal lines through every number -
    # whereas under AUTOMATIC calc Excel repaints each move cleanly. Nothing is
    # dirty now, so this does NOT restart the flag-image recalc storm.
    try { $excel.Calculation = -4105 } catch {}   # xlCalculationAutomatic
    try { $excel.CalculateUntilAsyncQueriesDone() } catch {}

    Write-Host "[3/3] Building combined PDF..."
    # Put the 8 sheets in report order (in memory only - not saved). The PDF
    # follows TAB order, so we line the tabs up Soybean..Beef. We move each sheet
    # to the FRONT in REVERSE order: move Beef to front, then Chicken, ... ending
    # with Soybean, which leaves them as Soybean, Meal, ... , Beef.
    # IMPORTANT: these Moves are DIRECT, not wrapped in Try-COM. Wrapping a Move
    # (like a Select) in the retry helper disrupts its positioning side-effect,
    # which is what pushed Soybean to the last page. We use a plain retry instead.
    for ($i = $order.Count - 1; $i -ge 0; $i--) {
        for ($try = 0; $try -lt 10; $try++) {
            try { $wb.Worksheets.Item($order[$i]).Move($wb.Worksheets.Item(1)); break }
            catch { Start-Sleep -Milliseconds 200 }
        }
    }
    $tabs = @(); for ($i = 1; $i -le $order.Count; $i++) { try { $tabs += $wb.Worksheets.Item($i).Name } catch {} }
    Write-Host ("    tab order: " + ($tabs -join ', '))

    # Switch the PDF printer to A4 so the page comes out A4 (not Letter) and the
    # report fills the width. (See Set-PdfPaperA4 above; restored in finally.)
    $a4ok = Set-PdfPaperA4
    if ($a4ok) { Write-Host "    PDF printer set to A4 for this export."; $pageW = 595.28; $pageH = 841.89 }
    else       { Write-Host "    PDF printer left at its default (Letter).";  $pageW = 612.0;  $pageH = 792.0 }

    # FILL THE PAGE. We do NOT re-assign PrintArea / PaperSize / FitToPages here:
    # re-setting those pagination properties over COM is slow AND smears border
    # lines through every number. The only thing we touch is a FIXED Zoom %, which
    # renders cleanly. Excel's own "fit to 1 page" never scales UP past 100%, so an
    # A4-sized block that's a hair smaller than the sheet prints with a margin all
    # round; computing the exact fill zoom from the print-area size vs the usable
    # page area makes each sheet fill top-to-bottom.
    foreach ($name in $order) {
        try {
            $ws = $wb.Worksheets.Item($name)
            $ps = $ws.PageSetup
            $pa = $ps.PrintArea
            if ([string]::IsNullOrEmpty($pa)) { continue }    # no print area -> leave saved setup
            $rng = $ws.Range($pa)
            $cw = [double]$rng.Width; $ch = [double]$rng.Height
            if ($cw -le 0 -or $ch -le 0) { continue }
            $uw = $pageW - $ps.LeftMargin - $ps.RightMargin
            $uh = $pageH - $ps.TopMargin  - $ps.BottomMargin
            # floor so we never tip just over a page boundary into a 2nd page
            $z = [Math]::Floor([Math]::Min($uw / $cw, $uh / $ch) * 100)
            if ($z -lt 10)  { $z = 10 }
            if ($z -gt 400) { $z = 400 }     # Excel's zoom limits
            $ps.Zoom = [int]$z
        } catch { Write-Host "    (fill-zoom skipped for ${name}: $($_.Exception.Message))" }
    }

    # Make sure Excel is idle one more time before selecting/exporting.
    try { $excel.CalculateUntilAsyncQueriesDone() } catch {}

    # Group-select the 8 sheets so the single export call emits all of them.
    # IMPORTANT: these .Select() calls must be DIRECT. Wrapping them in Try-COM
    # silently breaks the "extend selection" (Select($false)) behaviour, so only
    # ONE sheet ends up selected and the PDF comes out a single page. Instead we
    # do them directly and VERIFY the count, redoing the whole selection if it
    # didn't take (handles a momentarily-busy Excel without losing accumulation).
    $selOK = $false
    for ($attempt = 0; $attempt -lt 12 -and -not $selOK; $attempt++) {
        try {
            $wb.Worksheets.Item($order[0]).Select($true)
            for ($i = 1; $i -lt $order.Count; $i++) { $wb.Worksheets.Item($order[$i]).Select($false) }
            if ($excel.ActiveWindow.SelectedSheets.Count -eq $order.Count) { $selOK = $true }
            else { Start-Sleep -Milliseconds 300 }
        } catch { Start-Sleep -Milliseconds 300 }
    }
    $selCount = try { $excel.ActiveWindow.SelectedSheets.Count } catch { 0 }
    Write-Host "    grouped $selCount of $($order.Count) sheets for export."

    # 0 = xlTypePDF. Respects each sheet's print area / page setup. The export
    # itself is a single call with no accumulation, so Try-COM is safe here.
    Try-COM { $wb.ActiveSheet.ExportAsFixedFormat(0, $pdf) }

    Write-Host ""
    Write-Host "Done."
    Write-Host "  PDF : $pdf"
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    exit 1
}
finally {
    # Put the PDF printer's paper size back the way we found it.
    Restore-PdfPaper
    # Close WITHOUT saving so the temporary sheet re-order (and the export zoom)
    # is discarded (the refreshed data was already saved above).
    if ($wb)    { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() }     catch {} }
    if ($wb)    { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }    catch {} }
    if ($excel) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
