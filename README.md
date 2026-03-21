# Excel AI

A Microsoft 365 Excel Add-in that adds a custom function `=AI()` powered by a local AI model (LM Studio) or any OpenAI-compatible API.

## Usage

```
=EXCELAI.AI("What is the capital of France?")
=EXCELAI.AI("Classify sentiment", A1)
=EXCELAI.AI("Summarize these items", A1:A10)
=EXCELAI.AI("Translate to Spanish", A1, "qwen2.5-7b")
```

## Install (no coding required)

The add-in is hosted on GitHub Pages. To use it:

1. Download [`manifest.xml`](https://github.com/LoriTira/excel-ai/releases/latest) from the latest release
2. Open Excel > **Insert** > **My Add-ins** > **Upload My Add-in** > browse to `manifest.xml`
3. Open the **Excel AI** taskpane and configure your AI provider:
   - **Local (LM Studio)**: Have [LM Studio](https://lmstudio.ai/) running with a model loaded, server started, and CORS enabled
   - **External API**: Enter your API endpoint, key, and model name (works with OpenAI, Anthropic, or any OpenAI-compatible service)
4. Use `=EXCELAI.AI("your prompt")` in any cell

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
