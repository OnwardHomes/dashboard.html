# Watches the raw data folders and rebuilds raw_data_compiled.js when files change.
# Usage (from z:\Dashboards\TV):
#   powershell -ExecutionPolicy Bypass -File .\watch_build_raw_data.ps1

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$targets = @(
    Join-Path -Path $root -ChildPath "Arrears_Data"
    Join-Path -Path $root -ChildPath "Vacancy_Data"
)

Write-Host "Watching data folders for changes..." -ForegroundColor Cyan
Write-Host "Targets:" $targets -join ", "

function Build-RawData {
    try {
        Write-Host "`n[build] Starting build_raw_data.py..." -ForegroundColor Yellow
        $result = & python "$root\build_raw_data.py" 2>&1
        Write-Host $result
        Write-Host "[build] Done at $(Get-Date -Format 'HH:mm:ss')." -ForegroundColor Green
    } catch {
        Write-Host "[build] Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Debounce timer (milliseconds)
$debounceMs = 1500
$timer = New-Object Timers.Timer
$timer.Interval = $debounceMs
$timer.AutoReset = $false
$timer.add_Elapsed({ Build-RawData })

$watchers = @()
foreach ($path in $targets) {
    if (-not (Test-Path $path)) { continue }
    $fsw = New-Object IO.FileSystemWatcher $path -Property @{
        IncludeSubdirectories = $true
        Filter = "*.xlsx"
        EnableRaisingEvents = $true
    }
    $handlers = @("Changed","Created","Deleted","Renamed") | ForEach-Object {
        Register-ObjectEvent $fsw $_ -Action {
            $script:timer.Stop()
            $script:timer.Start()
        }
    }
    $watchers += @{ fsw = $fsw; handlers = $handlers }
}

Write-Host "Watcher armed. Press Ctrl+C to stop." -ForegroundColor Cyan

try {
    Build-RawData  # initial build
    while ($true) { Start-Sleep -Seconds 1 }
} finally {
    foreach ($w in $watchers) {
        foreach ($h in $w.handlers) { Unregister-Event -SubscriptionId $h.Id }
        $w.fsw.EnableRaisingEvents = $false
        $w.fsw.Dispose()
    }
    $timer.Stop(); $timer.Dispose()
    Write-Host "`nWatcher stopped."
}
