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
- **Local (Ollama)**: Auto-detects whichever model is loaded. Override per cell with `=EXCELAI.AI("prompt", , "gemma3:4b")`
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

## Model comparison

We benchmarked 9 local models on 10 realistic Excel tasks (factual Q&A, sentiment classification, unit conversion, translation, summarization, email extraction, Excel formula generation, prime check, math, expense categorization). All tests used `num_ctx=2048`, which is plenty for cell prompts and keeps RAM low.

| Model | Disk | RAM | Avg latency | Score | Notes |
|---|---|---|---|---|---|
| **llama3.2:1b** | 1.3 GB | ~1.5 GB | 116 ms | **9.5/10** | **Default.** Best accuracy/size ratio. Nails `=SUM(A:A)`. |
| gemma3:4b | 3.3 GB | ~4.2 GB | 270 ms | 9/10 | Best raw accuracy. Only misses prime check. |
| granite3-moe:3b | 2.1 GB | ~2.2 GB | 103 ms | 9/10 | Fastest overall. One math error (0.345 vs 34.5). |
| phi4-mini | 2.5 GB | ~2.8 GB | 198 ms | 8/10 | Strong reasoning. Misses conversion + prime. |
| llama3.2:3b | 2.0 GB | ~2.3 GB | 166 ms | 8/10 | Solid all-round. Misses conversion + formula. |
| qwen2.5:1.5b | 986 MB | ~1.1 GB | 113 ms | 7/10 | Smallest + fastest. Weaker on formulas + math. |
| gemma3:1b | 815 MB | ~1.3 GB | 158 ms | 7/10 | Tiny. Catastrophic math error (217 for 34.5). |
| smollm2:1.7b | 1.8 GB | — | 120 ms | 5/10 | Wrong on sentiment, conversion, prime, math. |
| granite3-moe:1b | 821 MB | — | 189 ms | 3/10 | Refuses questions, verbose, wrong math. |

> **RAM note:** Without `num_ctx=2048`, models with 128K default context allocate 5–18 GB of RAM. Excel AI sets this automatically so even 8 GB machines work fine.

To switch models, change the model name in the Excel AI settings panel, or pass `--model` to the install script:
```bash
curl -fsSL https://loritira.github.io/excel-ai/install.sh | bash -s -- --model gemma3:4b
```

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
