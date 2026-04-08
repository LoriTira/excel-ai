param(
    [string]$BaseUrl = "https://loritira.github.io/excel-ai",
    [string]$Model = "qwen2.5:1.5b"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"  # Makes Invoke-WebRequest 10-50x faster

$AddinId = "5cc82bad-f723-4c10-82ba-f135772ad04f"
$InstallDir = Join-Path $env:LOCALAPPDATA "ExcelAI"
$OllamaDir = Join-Path $env:USERPROFILE ".ollama"
$CertPath = Join-Path $OllamaDir "cert.pem"
$KeyPath = Join-Path $OllamaDir "key.pem"
$ServerBinary = Join-Path $InstallDir "excelai-server.exe"
$ManifestPath = Join-Path $InstallDir "manifest.xml"
$RegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
$TaskName = "ExcelAI-Server"

Write-Host "Installing Excel AI..."
Write-Host "Source: $BaseUrl"
Write-Host ""

# --- 1. Install Ollama ---

$ollamaCmd = Get-Command ollama -ErrorAction SilentlyContinue
if ($ollamaCmd) {
    Write-Host "[1/5] Ollama is already installed."
} else {
    Write-Host "[1/5] Installing Ollama..."
    $installerPath = Join-Path $env:TEMP "OllamaSetup.exe"
    & curl.exe -# -L -o $installerPath "https://ollama.com/download/OllamaSetup.exe"
    Start-Process -FilePath $installerPath -ArgumentList "/VERYSILENT","/NORESTART"
    # Installer launches Ollama and doesn't exit — poll for ollama on PATH instead
    do {
        Start-Sleep -Seconds 2
        $machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
        $userPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
        $env:Path = "$machinePath;$userPath"
    } until (Get-Command ollama -ErrorAction SilentlyContinue)
    Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
}

# Clean up any old HTTPS config from previous installs
[System.Environment]::SetEnvironmentVariable("OLLAMA_HOST", $null, "User")
$env:OLLAMA_HOST = $null

# --- 2. Generate TLS certificate ---

Write-Host "[2/5] Setting up TLS certificate..."

if (-not (Test-Path $OllamaDir)) {
    New-Item -ItemType Directory -Path $OllamaDir -Force | Out-Null
}

if (-not (Test-Path $CertPath)) {
    $cert = New-SelfSignedCertificate `
        -Subject "CN=Excel AI Local" `
        -TextExtension @("2.5.29.17={text}IPAddress=127.0.0.1") `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -KeyExportPolicy Exportable `
        -NotAfter (Get-Date).AddYears(10)

    # Trust it in the current user root store (no admin needed)
    $rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
    $rootStore.Open("ReadWrite")
    $rootStore.Add($cert)
    $rootStore.Close()

    # Export certificate as PEM for the Go server
    $certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    $certBase64 = [Convert]::ToBase64String($certBytes, 'InsertLineBreaks')
    "-----BEGIN CERTIFICATE-----`n$certBase64`n-----END CERTIFICATE-----" | Out-File -FilePath $CertPath -Encoding ascii -NoNewline

    # Export private key as PEM
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
    $keyBytes = $rsa.Key.Export([System.Security.Cryptography.CngKeyBlobFormat]::Pkcs8PrivateBlob)
    $keyBase64 = [Convert]::ToBase64String($keyBytes, 'InsertLineBreaks')
    "-----BEGIN PRIVATE KEY-----`n$keyBase64`n-----END PRIVATE KEY-----" | Out-File -FilePath $KeyPath -Encoding ascii -NoNewline

    # Remove from personal store (only needed in root)
    $personalStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $personalStore.Open("ReadWrite")
    $personalStore.Remove($cert)
    $personalStore.Close()
}

# --- 3. Download server binary ---

Write-Host "[3/5] Downloading Excel AI server..."

if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
}

# Detect architecture
switch ($env:PROCESSOR_ARCHITECTURE) {
    "ARM64" { $binaryName = "excelai-server-windows-arm64.exe" }
    default { $binaryName = "excelai-server-windows-amd64.exe" }
}

# Stop existing server before overwriting the binary
Get-Process -Name "excelai-server" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 1

Invoke-WebRequest -Uri "$BaseUrl/$binaryName" -OutFile $ServerBinary -UseBasicParsing
Invoke-WebRequest -Uri "$BaseUrl/manifest-local.xml" -OutFile $ManifestPath -UseBasicParsing

# --- 4. Register manifest and configure auto-start ---

Write-Host "[4/5] Configuring add-in and auto-start..."

if (-not (Test-Path $RegPath)) {
    New-Item -Path $RegPath -Force | Out-Null
}
New-ItemProperty -Path $RegPath -Name $AddinId -Value $ManifestPath -PropertyType String -Force | Out-Null

# Create/replace scheduled task to start server at login
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
$action = New-ScheduledTaskAction -Execute $ServerBinary
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero)
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null

# Start the server now
Start-Process -FilePath $ServerBinary -WindowStyle Hidden

# Ensure Ollama is running (Ollama auto-starts after install; only start if not running)
if (-not (Get-Process -Name "ollama" -ErrorAction SilentlyContinue)) {
    Start-Process ollama -ArgumentList "serve" -WindowStyle Hidden
    Start-Sleep -Seconds 3
}

# --- 5. Pull the default model ---

Write-Host "[5/5] Pulling model '$Model' (this may take a minute)..."
& ollama pull $Model

Write-Host ""
Write-Host "Excel AI installed successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Close and reopen Microsoft Excel"
Write-Host '  2. Use =EXCELAI.AI("your prompt") in any cell'
Write-Host ""
Write-Host "Everything runs locally — your data never leaves your machine."
Write-Host "Works offline after this initial setup."
