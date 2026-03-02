# Quick Copy Power Query to Clipboard
# This script copies the main Power Query to your clipboard for easy pasting into Power BI/Excel

param(
    [ValidateSet('Appointments', 'Services', 'Staff')]
    [string]$QueryType = 'Appointments'
)

Write-Host "`n=== Power Query Copy Helper ===" -ForegroundColor Cyan

switch ($QueryType) {
    'Appointments' {
        $file = "PowerQuery-BookingsData.m"
        $description = "Main Appointments Query (with Business info, dates, customers)"
    }
    'Services' {
        $file = "PowerQuery-Services.m"
        $description = "Services Lookup Table"
    }
    'Staff' {
        $file = "PowerQuery-Staff.m"
        $description = "Staff Lookup Table"
    }
}

$filePath = Join-Path $PSScriptRoot $file

if (Test-Path $filePath) {
    $content = Get-Content $filePath -Raw
    Set-Clipboard -Value $content
    
    Write-Host "[OK] Copied to clipboard!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Query Type: $QueryType" -ForegroundColor Yellow
    Write-Host "Description: $description" -ForegroundColor Gray
    Write-Host "File: $file" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Next Steps:" -ForegroundColor Cyan
    Write-Host "1. Open Power BI Desktop or Excel" -ForegroundColor White
    Write-Host "2. Go to: Get Data > Blank Query" -ForegroundColor White
    Write-Host "3. Click: Advanced Editor" -ForegroundColor White
    Write-Host "4. Press Ctrl+A to select all, then Ctrl+V to paste" -ForegroundColor White
    Write-Host "5. Click Done" -ForegroundColor White
    Write-Host ""
    
    # Show preview of first few lines
    $lines = $content -split "`n" | Select-Object -First 10
    Write-Host "Preview (first 10 lines):" -ForegroundColor Gray
    Write-Host "------------------------" -ForegroundColor Gray
    $lines | ForEach-Object { Write-Host $_ -ForegroundColor DarkGray }
    Write-Host "..." -ForegroundColor DarkGray
    Write-Host ""
    
    $totalLines = ($content -split "`n").Count
    Write-Host "Total lines: $totalLines" -ForegroundColor Gray
    Write-Host ""
}
else {
    Write-Host "[ERROR] File not found: $file" -ForegroundColor Red
}

Write-Host "To copy a different query, run:" -ForegroundColor Yellow
Write-Host "  .\Copy-PowerQuery.ps1 -QueryType Appointments" -ForegroundColor Cyan
Write-Host "  .\Copy-PowerQuery.ps1 -QueryType Services" -ForegroundColor Cyan
Write-Host "  .\Copy-PowerQuery.ps1 -QueryType Staff" -ForegroundColor Cyan
