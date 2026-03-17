# Zhihuishu-Pro-PPT-Downloader Full Process Script
# Auto-detect manifest -> Download PPTs -> Merge to PDF -> Reconcile

$ErrorActionPreference = "Stop"

# Use UTF8 encoding for output to prevent mojibake
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "   Zhihuishu-Pro-PPT-Downloader Startup" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan

# 1. Sync environment
Write-Host "[1/5] Syncing Python environment..." -ForegroundColor Yellow
uv sync

# 2. Find Manifest
Write-Host "[2/5] Searching for course manifest in tampermonkey/ ..." -ForegroundColor Yellow
$manifest = Get-ChildItem "tampermonkey/*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($null -eq $manifest) {
    Write-Host "Error: No .json manifest found in tampermonkey/ directory!" -ForegroundColor Red
    Write-Host "Please export it from browser first."
    Pause
    exit
}

$manifestPath = $manifest.FullName
Write-Host "Found manifest: $($manifest.Name)" -ForegroundColor Green

# Create output directory
$outputFolder = "output"
if (!(Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

# 3. Batch Download
Write-Host "`n[3/5] Starting batch download..." -ForegroundColor Yellow
uv run zhihuishu-pro-ppt-downloader manifest-download "$manifestPath" --downloads-dir "./downloads"

# 4. Export PDF
Write-Host "`n[4/5] Exporting PDFs (Requires PowerPoint)..." -ForegroundColor Yellow

# 4a. Merged PDF
Write-Host "   -> Creating merged PDF with bookmarks..." -ForegroundColor Cyan
$outputPdf = "$outputFolder\Course_Merged_$(Get-Date -Format 'yyyyMMdd_HHmm').pdf"
try {
    uv run zhihuishu-pro-ppt-downloader manifest-export-pdf "$manifestPath" --downloads-dir "./downloads" --output "$outputPdf"
    Write-Host "   [OK] Merged PDF saved as $outputPdf" -ForegroundColor Green
} catch {
    Write-Host "   [FAIL] Merged PDF export failed." -ForegroundColor Red
}

# 4b. Individual PDFs
Write-Host "   -> Exporting individual PDFs..." -ForegroundColor Cyan
$outputIndividualDir = "$outputFolder\pdfs_$(Get-Date -Format 'yyyyMMdd_HHmm')"
try {
    uv run zhihuishu-pro-ppt-downloader manifest-export-pdf "$manifestPath" --downloads-dir "./downloads" --individual --output-dir "$outputIndividualDir"
    Write-Host "   [OK] Individual PDFs saved in folder: $outputIndividualDir" -ForegroundColor Green
} catch {
    Write-Host "   [FAIL] Individual PDF export failed." -ForegroundColor Red
}

# 5. Reconcile
Write-Host "`n[5/5] Performing final data reconciliation..." -ForegroundColor Yellow
uv run zhihuishu-pro-ppt-downloader manifest-reconcile "$manifestPath" --downloads-dir "./downloads"

Write-Host "`n===============================================" -ForegroundColor Cyan
Write-Host "   Process completed successfully!" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan

Pause
