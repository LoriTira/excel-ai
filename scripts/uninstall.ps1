$ErrorActionPreference = "Stop"

$AddinId = "5cc82bad-f723-4c10-82ba-f135772ad04f"
$InstallDir = Join-Path $env:LOCALAPPDATA "ExcelAI"
$RegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
$TaskName = "ExcelAI-Server"

Write-Host "Uninstalling Excel AI..."

# Stop and remove the server
$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Stopped Excel AI server."
}
Stop-Process -Name "excelai-server*" -Force -ErrorAction SilentlyContinue 2>$null

# Remove registry entry
if (Test-Path $RegPath) {
    $prop = Get-ItemProperty -Path $RegPath -Name $AddinId -ErrorAction SilentlyContinue
    if ($prop) {
        Remove-ItemProperty -Path $RegPath -Name $AddinId -Force
        Write-Host "Removed registry entry."
    }
}

# Remove install directory
if (Test-Path $InstallDir) {
    Remove-Item -Path $InstallDir -Recurse -Force
    Write-Host "Removed: $InstallDir"
}

# Clean up environment variables
[System.Environment]::SetEnvironmentVariable("OLLAMA_HOST", $null, "User")
[System.Environment]::SetEnvironmentVariable("OLLAMA_ORIGINS", $null, "User")

Write-Host ""
Write-Host "Excel AI has been uninstalled." -ForegroundColor Green
Write-Host "Restart Excel for the change to take effect."
Write-Host ""
Write-Host "Note: Ollama and its models were not removed. To uninstall Ollama:"
Write-Host "  winget uninstall Ollama.Ollama"
