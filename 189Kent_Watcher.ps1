# ============================================================
# 189Kent_Watcher.ps1
# Monitors 189Kent_CoSTracker.xlsx and exports live data to
# 189Kent_data.js so the dashboard HTML stays up to date.
#
# HOW TO USE:
#   Double-click 189Kent_StartWatcher.bat  (leave the window open)
#   Every time you save the Excel file the JS is regenerated.
#   Then just refresh the browser tab showing the dashboard.
# ============================================================

$FolderPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$XlsxPath   = Join-Path $FolderPath "189Kent_CoSTracker.xlsx"
$JsPath     = Join-Path $FolderPath "189Kent_data.js"

function ColNum([string]$col) {
    $n = 0
    foreach ($c in $col.ToCharArray()) { $n = $n * 26 + ([int][char]$c - 64) }
    return $n
}
function XlDate($raw) {
    if ($null -eq $raw -or "$raw".Trim() -eq "") { return "null" }
    try { $d = [double]$raw; if ($d -gt 1000) { return ('"' + [DateTime]::FromOADate($d).ToString('yyyy-MM-dd') + '"') } } catch {}
    return "null"
}
function XlBool($raw) {
    if ($null -eq $raw) { return "false" }
    if ("$raw".Trim().ToUpper().StartsWith("YES")) { return "true" }
    return "false"
}
function XlStr($raw) {
    if ($null -eq $raw -or "$raw".Trim() -eq "") { return "null" }
    $s = "$raw".Trim().Replace('\','\\').Replace('"','\"').Replace("`n",' ').Replace("`r",'')
    return ('"' + $s + '"')
}
function XlNum($raw) {
    if ($null -eq $raw -or "$raw".Trim() -eq "") { return "null" }
    try { return [long]$raw } catch { return "null" }
}

function Export-Data {
    Write-Host "$(Get-Date -Format 'HH:mm:ss')  Reading Excel data..."

    $xl       = $null
    $wb       = $null
    $ownExcel = $false

    try {
        # Try to grab the already-open workbook first
        try {
            $runXl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            foreach ($w in $runXl.Workbooks) {
                if ($w.FullName -like "*189Kent_CoSTracker*") { $xl = $runXl; $wb = $w; break }
            }
        } catch {}

        # Otherwise open a silent read-only Excel instance
        if ($null -eq $wb) {
            $xl               = New-Object -ComObject Excel.Application
            $xl.Visible       = $false
            $xl.DisplayAlerts = $false
            $wb               = $xl.Workbooks.Open($XlsxPath, 0, $true)
            $ownExcel         = $true
        }

        $ws = $wb.Sheets.Item("CoS Tracker")
        if ($null -eq $ws) { throw "Sheet 'CoS Tracker' not found" }

        $rows = @()
        $r = 8
        while ($true) {
            $unitVal = $ws.Cells($r, (ColNum "B")).Value2
            if ($null -eq $unitVal -or "$unitVal".Trim() -eq "") { break }
            $rows += [PSCustomObject]@{
                unit=$unitVal; name=$ws.Cells($r,(ColNum "D")).Value2
                agent=$ws.Cells($r,(ColNum "F")).Value2; status="$($ws.Cells($r,(ColNum 'G')).Value2)".Trim()
                dateIssued=$ws.Cells($r,(ColNum "H")).Value2; exchanged=$ws.Cells($r,(ColNum "I")).Value2
                price=$ws.Cells($r,(ColNum "J")).Value2; spec=$ws.Cells($r,(ColNum "Q")).Value2
                colour=$ws.Cells($r,(ColNum "R")).Value2; amalgamation=$ws.Cells($r,(ColNum "S")).Value2
                bespokeLot=$ws.Cells($r,(ColNum "T")).Value2; friendsFamily=$ws.Cells($r,(ColNum "U")).Value2
                curStatus=$ws.Cells($r,(ColNum "Y")).Value2; brief=$ws.Cells($r,(ColNum "Z")).Value2
                sketch=$ws.Cells($r,(ColNum "AA")).Value2
                layoutApproved=$ws.Cells($r,(ColNum "AB")).Value2
                costsIssued=$ws.Cells($r,(ColNum "AC")).Value2   # AC: 1.4 KEY MILESTONE: Layout and Cost Range Letter #1 Signed
                feasibility=$ws.Cells($r,(ColNum "AD")).Value2   # AD: 1.2 Buildability Review
                cad=$ws.Cells($r,(ColNum "AE")).Value2           # AE: 1.5 Sketch Layout Converted to CAD by Architect
                qsEst=$ws.Cells($r,(ColNum "AF")).Value2         # AF: 1.7 Letter #2 Issued: Updated Cost Range and Interiors Confirmed
                confirm=$ws.Cells($r,(ColNum "AG")).Value2       # AG: 1.8 KEY MILESTONE: Interiors and Cost Range Letter #2 Signed
                designEnd=$ws.Cells($r,(ColNum "AH")).Value2     # AH: 1.9 Move to Phase 2 - COST (150 days after COS date)
                builder=$ws.Cells($r,(ColNum "AI")).Value2       # AI: 2.1 Builder Pricing
                commercial=$ws.Cells($r,(ColNum "AJ")).Value2    # AJ: 2.3 Internal Commercial Review
                qsCert=$ws.Cells($r,(ColNum "AK")).Value2        # AK: 2.2 QS Cost Certification
                dovApproval=$ws.Cells($r,(ColNum "AL")).Value2   # AL: 3.1 Internal DoV Approval
                dovIssue=$ws.Cells($r,(ColNum "AM")).Value2      # AM: 3.2 DoV and SA Invoice Issued to Client
                dovDeadline=$ws.Cells($r,(ColNum "AN")).Value2   # AN: 3.3 Client Executes DoV
                hickory=$ws.Cells($r,(ColNum "AO")).Value2       # AO: 4.1 Instructed to Builder
                planning=$ws.Cells($r,(ColNum "AP")).Value2      # AP: 4.2 Planning / DA Modification Submission
                modApproval=$ws.Cells($r,(ColNum "AQ")).Value2   # AQ: 4.3 Modification Approval
                completionDeadline=$ws.Cells($r,(ColNum "AR")).Value2 # AR: 4.4 Project Completion Deadline / Registration of Strata Plan
                lead=$ws.Cells($r,(ColNum "BF")).Value2          # BF: COMMS LEAD
                bespokeLink=$(
                    $auCol = ColNum "AU"                         # AU: BESPOKE PLAN
                    $hlUrl = $null
                    foreach ($hl in $ws.Hyperlinks) {
                        if ($hl.Range.Row -eq $r -and $hl.Range.Column -eq $auCol) {
                            $hlUrl = $hl.Address
                            # Excel stores SharePoint links as relative paths (../../...) when the
                            # file is synced via OneDrive. Strip the leading ../ traversals and
                            # prepend the SharePoint base URL to get a working absolute URL.
                            if ($hlUrl -match '^(\.\.[\\/])+(.+)$') {
                                $hlUrl = 'https://uiservicesptyltd.sharepoint.com/' + $Matches[2]
                            }
                            break
                        }
                    }
                    if ($hlUrl) { $hlUrl } else { $null }
                )
                bic=$ws.Cells($r,(ColNum "BD")).Value2           # BD: BALL IN COURT
                nextSteps=$ws.Cells($r,(ColNum "BE")).Value2     # BE: NEXT STEPS
                correspondence=$(
                    $avCol = ColNum "AV"                         # AV: CORRESPONDENCE
                    $hlUrl = $null
                    foreach ($hl in $ws.Hyperlinks) {
                        if ($hl.Range.Row -eq $r -and $hl.Range.Column -eq $avCol) {
                            $hlUrl = $hl.Address
                            if ($hlUrl -match '^(\.\.[\\/])+(.+)$') {
                                $hlUrl = 'https://uiservicesptyltd.sharepoint.com/' + $Matches[2]
                            }
                            break
                        }
                    }
                    if ($hlUrl) { $hlUrl } else { $ws.Cells($r, $avCol).Value2 }
                )
            }
            $r++
        }

        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add("// Auto-generated from 189Kent_CoSTracker.xlsx")
        $lines.Add("// Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        $lines.Add("// Do NOT edit manually - overwritten by 189Kent_StartWatcher.bat on each Excel save.")
        $lines.Add("window.UNITS_FROM_EXCEL = [")

        for ($i = 0; $i -lt $rows.Count; $i++) {
            $u = $rows[$i]
            $comma = if ($i -lt $rows.Count-1) {","} else {""}
            $bespoke = if ("$($u.bespokeLot)".Trim().ToUpper() -eq "YES") {"true"} else {"false"}
            $exchDate = if ($u.status -eq "EXCHANGED") {XlDate $u.exchanged} else {XlDate $u.dateIssued}
            $line = "  {unit:$(XlStr $u.unit),name:$(XlStr $u.name),agent:$(XlStr $u.agent)," +
                    "status:$(XlStr $u.status),price:$(XlNum $u.price),bespoke:$bespoke," +
                    "exchanged:$exchDate,spec:$(XlStr $u.spec),colour:$(XlStr $u.colour)," +
                    "amalgamation:$(XlStr $u.amalgamation),bespokeLot:$(XlStr $u.bespokeLot)," +
                    "friendsFamily:$(XlStr $u.friendsFamily),curStatus:$(XlStr $u.curStatus)," +
                    "brief:$(XlDate $u.brief),designEnd:$(XlDate $u.designEnd)," +
                    "qsCert:$(XlDate $u.qsCert),dovIssue:$(XlDate $u.dovIssue)," +
                    "dovDeadline:$(XlDate $u.dovDeadline),hickory:$(XlDate $u.hickory)," +
                    "planning:$(XlDate $u.planning),modApproval:$(XlDate $u.modApproval)," +
                    "completionDeadline:$(XlDate $u.completionDeadline)," +
                    "sketch:$(XlBool $u.sketch),layoutApproved:$(XlBool $u.layoutApproved)," +
                    "costsIssued:$(XlBool $u.costsIssued),feasibility:$(XlBool $u.feasibility)," +
                    "cad:$(XlBool $u.cad),qsEst:$(XlBool $u.qsEst),confirm:$(XlBool $u.confirm)," +
                    "builder:$(XlBool $u.builder),commercial:$(XlBool $u.commercial)," +
                    "dovApproval:$(XlBool $u.dovApproval),bic:$(XlStr $u.bic)," +
                    "lead:$(XlStr $u.lead),nextSteps:$(XlStr $u.nextSteps)," +
                    "correspondence:$(XlStr $u.correspondence)," +
                    "bespokeLink:$(XlStr $u.bespokeLink)}$comma"
            $lines.Add($line)
        }
        $lines.Add("];")
        $ts = Get-Date -Format 'dd MMM yyyy HH:mm'
        $lines.Add('window.EXCEL_LAST_UPDATED = "' + $ts + '";')

        [System.IO.File]::WriteAllLines($JsPath, $lines, [System.Text.Encoding]::UTF8)
        Write-Host "$(Get-Date -Format 'HH:mm:ss')  Done - $($rows.Count) units exported to 189Kent_data.js"
        Write-Host "          Refresh the dashboard in your browser to see updates."
        Write-Host ""

    } catch {
        Write-Host "$(Get-Date -Format 'HH:mm:ss')  ERROR: $_" -ForegroundColor Red
    } finally {
        if ($ownExcel -and $null -ne $wb)  { try { $wb.Close($false) } catch {} }
        if ($ownExcel -and $null -ne $xl)  { try { $xl.Quit() } catch {}; try { [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {} }
    }
}

# ── Startup ────────────────────────────────────────────────
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  189 Kent -- CoS Dashboard Data Watcher" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Watching: $XlsxPath"
Write-Host "Output:   $JsPath"
Write-Host ""
Write-Host "Leave this window open. Every time you save the Excel"
Write-Host "file, the dashboard data will refresh automatically."
Write-Host "Press Ctrl+C to stop."
Write-Host ""

Export-Data

# ── FileSystemWatcher ───────────────────────────────────────
$watcher                     = New-Object System.IO.FileSystemWatcher
$watcher.Path                = $FolderPath
$watcher.Filter              = "189Kent_CoSTracker.xlsx"
$watcher.NotifyFilter        = [System.IO.NotifyFilters]::LastWrite
$watcher.EnableRaisingEvents = $true

$lastFired = [DateTime]::MinValue

$action = {
    $now = [DateTime]::Now
    if (($now - $script:lastFired).TotalSeconds -lt 3) { return }
    $script:lastFired = $now
    Start-Sleep -Milliseconds 1500
    Export-Data
}

Register-ObjectEvent $watcher "Changed" -Action $action | Out-Null

try {
    while ($true) { Start-Sleep -Seconds 5 }
} finally {
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
}
