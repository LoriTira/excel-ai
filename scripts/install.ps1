param(
    [string]$BaseUrl = "https://loritira.github.io/excel-ai"
)

$ErrorActionPreference = "Stop"

$AddinId = "5cc82bad-f723-4c10-82ba-f135772ad04f"
$ManifestDir = Join-Path $env:LOCALAPPDATA "ExcelAI"
$ManifestPath = Join-Path $ManifestDir "manifest.xml"
$RegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"

Write-Host "Installing Excel AI add-in..."
Write-Host "Source: $BaseUrl"

# Create manifest directory
if (-not (Test-Path $ManifestDir)) {
    New-Item -ItemType Directory -Path $ManifestDir -Force | Out-Null
}

# Download the production manifest
$ManifestUrl = "$BaseUrl/manifest.xml"
Invoke-WebRequest -Uri $ManifestUrl -OutFile $ManifestPath -UseBasicParsing

# Create registry key for sideloading
if (-not (Test-Path $RegPath)) {
    New-Item -Path $RegPath -Force | Out-Null
}
New-ItemProperty -Path $RegPath -Name $AddinId -Value $ManifestPath -PropertyType String -Force | Out-Null

Write-Host ""
Write-Host "Excel AI installed successfully!" -ForegroundColor Green
Write-Host "  Manifest: $ManifestPath"
Write-Host "  Registry: $RegPath\$AddinId"
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Close and reopen Microsoft Excel"
Write-Host "  2. Open the Excel AI taskpane (Home tab > Excel AI)"
Write-Host "  3. Configure your AI provider (LM Studio or external API)"
Write-Host '  4. Use =EXCELAI.AI("your prompt") in any cell'
