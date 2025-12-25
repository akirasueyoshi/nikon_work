# Setup script for Windows PowerShell

Write-Host "================================" -ForegroundColor Cyan
Write-Host "Document Relevance Matrix Setup" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# Check if uv is installed
try {
    $uvVersion = uv --version 2>$null
    Write-Host "✓ uv is already installed: $uvVersion" -ForegroundColor Green
} catch {
    Write-Host "Installing uv..." -ForegroundColor Yellow
    Invoke-RestMethod https://astral.sh/uv/install.ps1 | Invoke-Expression
    # Refresh PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

# Sync dependencies
Write-Host ""
Write-Host "Installing dependencies..." -ForegroundColor Yellow
uv sync

Write-Host ""
Write-Host "================================" -ForegroundColor Green
Write-Host "✓ Setup completed!" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green
Write-Host ""
Write-Host "Try running:" -ForegroundColor Cyan
Write-Host "  uv run extract-links examples/test_files" -ForegroundColor White
Write-Host "  uv run build-matrix extraction_results/document_graph_*.json" -ForegroundColor White
