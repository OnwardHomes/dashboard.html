# Monitors Vacancy_Data/ and Arrears_Data/ for XLSX changes and rebuilds raw_data_compiled.js.
# Usage: powershell -ExecutionPolicy Bypass -File build_watch.ps1
# Optional: add -NoBuildOnStart to skip the initial build.
[CmdletBinding()]
param(
  [switch]$NoBuildOnStart
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$debounceSeconds = 5
$pending = $false
$lastEvent = Get-Date

function Write-Note($msg) {
  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Write-Host "[$ts] $msg"
}

function Run-Build {
  Write-Note "Running build_raw_data.py..."
  try {
    python "$root\build_raw_data.py" | Write-Host
    Write-Note "Build complete."
  } catch {
    Write-Note "Build failed: $($_.Exception.Message)"
  }
}

function Queue-Build {
  $script:pending = $true
  $script:lastEvent = Get-Date
}

Write-Note "Watching for XLSX changes under Vacancy_Data/ and Arrears_Data/ (debounce ${debounceSeconds}s). Press Ctrl+C to stop."

$watchers = @()
foreach ($pair in @(
  @{ Path = Join-Path $root "Vacancy_Data"; Filter = "*.xlsx" },
  @{ Path = Join-Path $root "Arrears_Data"; Filter = "*.xlsx" }
)) {
  $fsw = New-Object IO.FileSystemWatcher $pair.Path, $pair.Filter
  $fsw.IncludeSubdirectories = $false
  $fsw.EnableRaisingEvents = $true
  $watchers += $fsw
  foreach ($evt in @("Created","Changed","Renamed")) {
    Register-ObjectEvent -InputObject $fsw -EventName $evt -Action { Queue-Build } | Out-Null
  }
}

if (-not $NoBuildOnStart) {
  Run-Build
}

try {
  while ($true) {
    Wait-Event -Timeout 2 | Out-Null
    if ($pending -and ((Get-Date) - $lastEvent).TotalSeconds -ge $debounceSeconds) {
      $pending = $false
      Run-Build
    }
  }
} finally {
  Get-EventSubscriber | Unregister-Event
  $watchers | ForEach-Object { $_.Dispose() }
  Write-Note "Watcher stopped."
}
