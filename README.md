# Excel AI

A Microsoft 365 Excel Add-in that adds a custom function `=AI()` powered by a local AI model (LM Studio) or any OpenAI-compatible API.

## Usage

```
=EXCELAI.AI("What is the capital of France?")
=EXCELAI.AI("Classify sentiment", A1)
=EXCELAI.AI("Summarize these items", A1:A10)
=EXCELAI.AI("Translate to Spanish", A1, "qwen2.5-7b")
```

## Install

The add-in is hosted on GitHub Pages. Install with a single command — no coding required.

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
2. Click **Excel AI** in the Home tab to open settings
3. Configure your AI provider:
   - **Local (LM Studio)**: Have [LM Studio](https://lmstudio.ai/) running with a model loaded, server started, and CORS enabled
   - **External API**: Enter your API endpoint, key, and model name (works with OpenAI, Anthropic, or any OpenAI-compatible service)
4. Use `=EXCELAI.AI("your prompt")` in any cell

### Manual install

If you prefer not to use the script, download [`manifest.xml`](https://loritira.github.io/excel-ai/manifest.xml) and sideload it in Excel via **Insert** > **My Add-ins** > **Upload My Add-in**.

See the full [install page](https://loritira.github.io/excel-ai/install.html) for more details.

### Uninstall

**macOS:** `curl -fsSL https://loritira.github.io/excel-ai/uninstall.sh | bash`

**Windows:** `irm https://loritira.github.io/excel-ai/uninstall.ps1 | iex`

## Development

For local development with full LM Studio proxy support:

```bash
# Quick start (handles nvm automatically)
./start.sh

# Or manually
nvm use 22
npm install
npm start
```

This builds the add-in and sideloads it into Excel with a local dev server.

## How it works

The `=AI()` custom function sends prompts to LM Studio's OpenAI-compatible API or an external API endpoint and returns the response text into the cell. Results are cached in memory to avoid re-querying on recalculation.

When running from GitHub Pages, the add-in calls LM Studio directly. When running from the local dev server, requests are proxied through webpack to avoid CORS issues.
