# Excel AI

A Microsoft 365 Excel Add-in that adds a custom function `=AI()` powered by a local AI model via [Ollama](https://ollama.com/) or any OpenAI-compatible API. Your data never leaves your machine.

## Usage

```
=EXCELAI.AI("What is the capital of France?")
=EXCELAI.AI("Classify sentiment", A1)
=EXCELAI.AI("Summarize these items", A1:A10)
=EXCELAI.AI("Translate to Spanish", A1, "qwen2.5-7b")
```

## Install

One command installs everything — the Excel add-in, Ollama, a default AI model, and a local HTTPS server. No coding required. Works offline after setup.

### macOS

Open **Terminal** and paste:

```bash
curl -fsSL https://loritira.github.io/excel-ai/install.sh | bash
```

### Windows

Open **PowerShell** and paste:

```powershell
irm https://loritira.github.io/excel-ai/install.ps1 | iex
```

### After installation

1. Quit and reopen Excel
2. Use `=EXCELAI.AI("your prompt")` in any cell — it works immediately with the bundled local model

To change settings, click **Excel AI** in the Home tab:
- **Local (Ollama)**: Change the model name
- **External API**: Use OpenAI, Anthropic, or any OpenAI-compatible service instead

### Stop / restart

**macOS:**
```bash
# Stop
launchctl unload ~/Library/LaunchAgents/com.excelai.server.plist
pkill ollama

# Restart
launchctl load ~/Library/LaunchAgents/com.excelai.server.plist
ollama serve &
```

**Windows (PowerShell):**
```powershell
# Stop
Stop-ScheduledTask -TaskName "ExcelAI-Server"; Stop-Process -Name "excelai-server" -Force
Stop-Process -Name "ollama" -Force

# Restart
Start-ScheduledTask -TaskName "ExcelAI-Server"
Start-Process ollama -ArgumentList "serve" -WindowStyle Hidden
```

### Uninstall

**macOS:** `curl -fsSL https://loritira.github.io/excel-ai/uninstall.sh | bash`

**Windows:** `irm https://loritira.github.io/excel-ai/uninstall.ps1 | iex`

## How it works

The install script sets up three components that run locally:

1. **Ollama** — runs AI models on your machine (HTTP on port 11434)
2. **Excel AI server** — a tiny Go binary (~8MB) that serves the add-in UI and proxies API requests to Ollama over HTTPS (port 11435)
3. **The add-in manifest** — tells Excel where to load the add-in from

```
Excel Add-in (HTTPS) ←→ Excel AI server (HTTPS :11435) ←→ Ollama (HTTP :11434)
```

The HTTPS server is needed because Office Add-ins require HTTPS, but Ollama only speaks HTTP. The server bridges this gap while keeping everything local.

All AI processing happens on your machine — no data is sent to any external service unless you explicitly configure an external API provider.

## Development

```bash
# Quick start
./start.sh

# Or manually
nvm use 22
npm install
npm start
```

The webpack dev server proxies `/v1/*` to Ollama on `localhost:11434` automatically, so the same code works in both dev and production.

### Building the server binary locally

```bash
npm run build
rm -rf server/static && cp -r dist server/static
cd server && go build -ldflags="-s -w" -o excelai-server .
```
