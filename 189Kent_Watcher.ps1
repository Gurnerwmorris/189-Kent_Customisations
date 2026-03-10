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
                agent=$ws.Cells($r,(ColNum "E")).Value2; status="$($ws.Cells($r,(ColNum 'F')).Value2)".Trim()
                dateIssued=$ws.Cells($r,(ColNum "G")).Value2; exchanged=$ws.Cells($r,(ColNum "H")).Value2
                price=$ws.Cells($r,(ColNum "I")).Value2; spec=$ws.Cells($r,(ColNum "P")).Value2
                colour=$ws.Cells($r,(ColNum "Q")).Value2; amalgamation=$ws.Cells($r,(ColNum "R")).Value2
                bespokeLot=$ws.Cells($r,(ColNum "S")).Value2; friendsFamily=$ws.Cells($r,(ColNum "T")).Value2
                curStatus=$ws.Cells($r,(ColNum "X")).Value2; brief=$ws.Cells($r,(ColNum "Y")).Value2
                sketch=$ws.Cells($r,(ColNum "Z")).Value2; feasibility=$ws.Cells($r,(ColNum "AA")).Value2
                cad=$ws.Cells($r,(ColNum "AC")).Value2; qsEst=$ws.Cells($r,(ColNum "AD")).Value2
                confirm=$ws.Cells($r,(ColNum "AE")).Value2; designEnd=$ws.Cells($r,(ColNum "AF")).Value2
                builder=$ws.Cells($r,(ColNum "AG")).Value2; commercial=$ws.Cells($r,(ColNum "AH")).Value2
                qsCert=$ws.Cells($r,(ColNum "AI")).Value2; dovApproval=$ws.Cells($r,(ColNum "AJ")).Value2
                dovIssue=$ws.Cells($r,(ColNum "AK")).Value2; dovDeadline=$ws.Cells($r,(ColNum "AL")).Value2
                hickory=$ws.Cells($r,(ColNum "AM")).Value2; planning=$ws.Cells($r,(ColNum "AN")).Value2
                modApproval=$ws.Cells($r,(ColNum "AO")).Value2; bic=$ws.Cells($r,(ColNum "AZ")).Value2
                nextSteps=$ws.Cells($r,(ColNum "BA")).Value2; lead=$ws.Cells($r,(ColNum "BB")).Value2
                bespokeLink=$(
                    $arCol = ColNum "AR"
                    $hlUrl = $null
                    foreach ($hl in $ws.Hyperlinks) {
                        if ($hl.Range.Row -eq $r -and $hl.Range.Column -eq $arCol) {
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
                    if ($hlUrl) { $hlUrl } else { $ws.Cells($r, $arCol).Value2 }
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
                    "sketch:$(XlBool $u.sketch),feasibility:$(XlBool $u.feasibility)," +
                    "cad:$(XlBool $u.cad),qsEst:$(XlBool $u.qsEst),confirm:$(XlBool $u.confirm)," +
                    "builder:$(XlBool $u.builder),commercial:$(XlBool $u.commercial)," +
                    "dovApproval:$(XlBool $u.dovApproval),bic:$(XlStr $u.bic)," +
                    "lead:$(XlStr $u.lead),nextSteps:$(XlStr $u.nextSteps)," +
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
