$ErrorActionPreference = "Stop"

$AddinId = "5cc82bad-f723-4c10-82ba-f135772ad04f"
$ManifestDir = Join-Path $env:LOCALAPPDATA "ExcelAI"
$RegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"

Write-Host "Uninstalling Excel AI add-in..."

# Remove registry entry
if (Test-Path $RegPath) {
    $prop = Get-ItemProperty -Path $RegPath -Name $AddinId -ErrorAction SilentlyContinue
    if ($prop) {
        Remove-ItemProperty -Path $RegPath -Name $AddinId -Force
        Write-Host "Removed registry entry: $RegPath\$AddinId"
    } else {
        Write-Host "Registry entry not found (already removed?)"
    }
}

# Remove manifest directory
if (Test-Path $ManifestDir) {
    Remove-Item -Path $ManifestDir -Recurse -Force
    Write-Host "Removed: $ManifestDir"
} else {
    Write-Host "Manifest directory not found (already removed?)"
}

Write-Host ""
Write-Host "Excel AI has been uninstalled." -ForegroundColor Green
Write-Host "Restart Excel for the change to take effect."
Write-Host ""
Write-Host "Note: Ollama was not removed. To uninstall Ollama, use Add/Remove Programs or:"
Write-Host "  winget uninstall Ollama.Ollama"
