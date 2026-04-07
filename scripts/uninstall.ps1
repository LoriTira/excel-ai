$ErrorActionPreference = "Stop"

$AddinId = "5cc82bad-f723-4c10-82ba-f135772ad04f"
$InstallDir = Join-Path $env:LOCALAPPDATA "ExcelAI"
$OllamaDir = Join-Path $env:USERPROFILE ".ollama"
$RegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
$TaskName = "ExcelAI-Server"

Write-Host "Uninstalling Excel AI (everything)..."
Write-Host ""

# --- 1. Stop and remove the Excel AI server ---

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}
Stop-Process -Name "excelai-server" -Force -ErrorAction SilentlyContinue

if (Test-Path $InstallDir) {
    Remove-Item -Path $InstallDir -Recurse -Force
}
Write-Host "[1/5] Removed Excel AI server."

# --- 2. Remove add-in manifest ---

if (Test-Path $RegPath) {
    $prop = Get-ItemProperty -Path $RegPath -Name $AddinId -ErrorAction SilentlyContinue
    if ($prop) {
        Remove-ItemProperty -Path $RegPath -Name $AddinId -Force
    }
}
Write-Host "[2/5] Removed add-in manifest."

# --- 3. Stop and remove Ollama ---

Write-Host "[3/5] Removing Ollama..."
Stop-Process -Name "ollama" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# Uninstall via winget (silent)
$ollamaInstalled = winget list --id Ollama.Ollama --accept-source-agreements 2>$null | Select-String "Ollama"
if ($ollamaInstalled) {
    winget uninstall --id Ollama.Ollama --silent --accept-source-agreements 2>$null
    Write-Host "  Uninstalled Ollama via winget."
} else {
    # Fallback: try the Ollama uninstaller directly
    $uninstaller = Join-Path $env:LOCALAPPDATA "Programs\Ollama\unins000.exe"
    if (Test-Path $uninstaller) {
        Start-Process -FilePath $uninstaller -ArgumentList "/VERYSILENT" -Wait
        Write-Host "  Uninstalled Ollama via uninstaller."
    } else {
        Write-Host "  Ollama not found (already removed?)."
    }
}

# --- 4. Remove all models, certs, and data ---

Write-Host "[4/5] Removing models and data..."
if (Test-Path $OllamaDir) {
    Remove-Item -Path $OllamaDir -Recurse -Force
    Write-Host "  Removed $OllamaDir (models, certs, config)"
}

# Remove leftover Ollama program files
$ollamaProgDir = Join-Path $env:LOCALAPPDATA "Programs\Ollama"
if (Test-Path $ollamaProgDir) {
    Remove-Item -Path $ollamaProgDir -Recurse -Force
    Write-Host "  Removed $ollamaProgDir"
}

# --- 5. Remove trusted certificate and env vars ---

Write-Host "[5/5] Cleaning up..."

# Remove the trusted cert from the user root store
$certs = Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -match "Excel AI Local" }
foreach ($cert in $certs) {
    $rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
    $rootStore.Open("ReadWrite")
    $rootStore.Remove($cert)
    $rootStore.Close()
}

[System.Environment]::SetEnvironmentVariable("OLLAMA_HOST", $null, "User")
[System.Environment]::SetEnvironmentVariable("OLLAMA_ORIGINS", $null, "User")

Write-Host ""
Write-Host "Excel AI has been completely uninstalled." -ForegroundColor Green
Write-Host "Restart Excel for the change to take effect."
