$folders = @(
    'briqlabAnimations-MS',
    'briqlabBarChart-MS',
    'briqlabBulletPro-MS',
    'briqlabCalendarHeat-MS',
    'briqlabDonutChart-MS',
    'briqlabDotMatrix-MS',
    'briqlabDrillBubble-MS',
    'briqlabDrillPie-MS',
    'briqlabFlowSankey-MS',
    'briqlabGauge-MS',
    'briqlabKPICard-MS',
    'briqlabKPISparkline-MS',
    'briqlabMekkoChart-MS',
    'briqlabPieChart-MS',
    'briqlabProgressRing-MS',
    'briqlabPulseKPI-MS',
    'briqlabRadarPro-MS',
    'briqlabScroller-MS',
    'briqlabSearch-MS',
    'briqlabSlopeChart-MS',
    'briqlabViolinPlot-MS',
    'briqlabWordCloud-MS'
)

$base = 'C:\Users\Abhijit\BriqlabVisuals'
$success = @()
$failed = @()

foreach ($folder in $folders) {
    $path = Join-Path $base $folder
    Write-Host "
=== Building $folder ===" -ForegroundColor Cyan
    Set-Location $path
    npm install --silent 2>$null
    npx pbiviz package 2>$null
    $pbiviz = Get-ChildItem "$path\dist" -Filter "*.pbiviz" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($pbiviz -and ($pbiviz.LastWriteTime -gt (Get-Date).AddMinutes(-5))) {
        Write-Host "OK: $($pbiviz.Name)" -ForegroundColor Green
        $success += $folder
    } else {
        Write-Host "FAILED: $folder" -ForegroundColor Red
        $failed += $folder
    }
}

Write-Host "
=== RESULTS ===" -ForegroundColor Yellow
Write-Host "Success: $($success.Count)/$($folders.Count)" -ForegroundColor Green
if ($failed.Count -gt 0) { Write-Host "Failed: $($failed -join ', ')" -ForegroundColor Red }
Set-Location $base
