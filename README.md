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

One command installs the Excel add-in, Ollama, and a default AI model — no coding required.

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
- **Local (Ollama)**: Change the model or server address
- **External API**: Use OpenAI, Anthropic, or any OpenAI-compatible service instead

### Manual install

If you prefer not to use the script, download [`manifest.xml`](https://loritira.github.io/excel-ai/manifest.xml) and sideload it in Excel via **Insert** > **My Add-ins** > **Upload My Add-in**. You'll need to [install Ollama](https://ollama.com/download) separately.

See the full [install page](https://loritira.github.io/excel-ai/install.html) for more details.

### Uninstall

**macOS:** `curl -fsSL https://loritira.github.io/excel-ai/uninstall.sh | bash`

**Windows:** `irm https://loritira.github.io/excel-ai/uninstall.ps1 | iex`

## Development

For local development with Ollama proxy support:

```bash
# Quick start (handles nvm automatically)
./start.sh

# Or manually
nvm use 22
npm install
npm start
```

This builds the add-in and sideloads it into Excel with a local dev server. The dev server proxies `/ollama` to `localhost:11434` automatically.

## How it works

The `=AI()` custom function sends prompts to Ollama's OpenAI-compatible API (or an external API endpoint) and returns the response text into the cell. Results are cached in memory to avoid re-querying on recalculation.

All AI processing happens locally on your machine by default — no data is sent to any external service unless you explicitly configure an external API provider.
