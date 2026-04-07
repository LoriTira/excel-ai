param(
    [string]$BaseUrl = "https://loritira.github.io/excel-ai",
    [string]$Model = "qwen2.5:1.5b"
)

$ErrorActionPreference = "Stop"

$AddinId = "5cc82bad-f723-4c10-82ba-f135772ad04f"
$ManifestDir = Join-Path $env:LOCALAPPDATA "ExcelAI"
$ManifestPath = Join-Path $ManifestDir "manifest.xml"
$RegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"

Write-Host "Installing Excel AI add-in..."
Write-Host "Source: $BaseUrl"
Write-Host ""

# --- 1. Install the Excel add-in manifest ---

if (-not (Test-Path $ManifestDir)) {
    New-Item -ItemType Directory -Path $ManifestDir -Force | Out-Null
}

$ManifestUrl = "$BaseUrl/manifest.xml"
Invoke-WebRequest -Uri $ManifestUrl -OutFile $ManifestPath -UseBasicParsing

if (-not (Test-Path $RegPath)) {
    New-Item -Path $RegPath -Force | Out-Null
}
New-ItemProperty -Path $RegPath -Name $AddinId -Value $ManifestPath -PropertyType String -Force | Out-Null

Write-Host "[1/3] Excel add-in manifest installed."

# --- 2. Install Ollama ---

$ollamaCmd = Get-Command ollama -ErrorAction SilentlyContinue
if ($ollamaCmd) {
    Write-Host "[2/3] Ollama is already installed."
} else {
    Write-Host "[2/3] Installing Ollama..."
    $installerPath = Join-Path $env:TEMP "OllamaSetup.exe"
    Invoke-WebRequest -Uri "https://ollama.com/download/OllamaSetup.exe" -OutFile $installerPath -UseBasicParsing
    Start-Process -FilePath $installerPath -ArgumentList "/VERYSILENT","/NORESTART" -Wait
    Remove-Item $installerPath -Force -ErrorAction SilentlyContinue

    # Refresh PATH so ollama is available in this session
    $machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machinePath;$userPath"
}

# --- 3. Pull the default model ---

Write-Host "[3/3] Pulling model '$Model' (this may take a minute)..."
& ollama pull $Model

Write-Host ""
Write-Host "Excel AI installed successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Close and reopen Microsoft Excel"
Write-Host '  2. Use =EXCELAI.AI("your prompt") in any cell'
Write-Host ""
Write-Host "Ollama is running locally - your data never leaves your machine."
