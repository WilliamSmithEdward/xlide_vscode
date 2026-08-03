param(
    [Parameter(Mandatory = $true)][string]$Manifest,
    [Parameter(Mandatory = $true)][string]$Module,
    [Parameter(Mandatory = $true)][string]$Procedure
)

# Worker for scripts/verify-excel-formats.mjs: open each file in Excel, run the
# sentinel macro, and report what Excel says about it.
#
# Only the Excel this script starts is ever touched. Snapshot the PIDs that
# already exist so an instance the developer has open is never counted, and
# never killed during cleanup.
$ErrorActionPreference = 'Stop'
$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

$excel = New-Object -ComObject Excel.Application
$after = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$ownPid = ($after | Where-Object { $before -notcontains $_ }) | Select-Object -First 1

$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AutomationSecurity = 1   # msoAutomationSecurityLow: run macros unattended

$results = @()
foreach ($file in (Get-Content $Manifest)) {
    $name = Split-Path $file -Leaf
    $row = [ordered]@{ name = $name; isAddin = $null; fileFormat = $null; macro = $null; error = $null }
    try {
        $wb = $excel.Workbooks.Open($file)
        $row.isAddin = [bool]$wb.IsAddin
        $row.fileFormat = [int]$wb.FileFormat
        # Fully qualified Project.Module.Procedure: the short forms are
        # ambiguous once more than one project is loaded.
        $row.macro = $excel.Run(("{0}.{1}.{2}" -f $wb.VBProject.Name, $Module, $Procedure))
        $wb.Close($false)
    } catch {
        $row.error = $_.Exception.Message
        if ($wb) { try { $wb.Close($false) } catch { } }
    }
    $results += [pscustomobject]$row
}

$out = Join-Path (Split-Path $Manifest -Parent) 'results.json'
ConvertTo-Json @($results) -Depth 4 | Out-File -FilePath $out -Encoding utf8

try { $excel.Quit() } catch { }
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
if ($ownPid) {
    Start-Sleep -Milliseconds 400
    $still = Get-Process -Id $ownPid -ErrorAction SilentlyContinue
    if ($still) { Stop-Process -Id $ownPid -Force }
}
