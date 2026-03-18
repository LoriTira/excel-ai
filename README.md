# Excel AI

A Microsoft 365 Excel Add-in that adds a custom function `=AI()` powered by a local AI model via LM Studio.

## Usage

```
=EXCELAI.AI("What is the capital of France?")
=EXCELAI.AI("Classify sentiment", A1)
=EXCELAI.AI("Summarize these items", A1:A10)
=EXCELAI.AI("Translate to Spanish", A1, "qwen2.5-7b")
```

## Requirements

- Node.js 22 LTS
- Microsoft 365 Excel (desktop)
- [LM Studio](https://lmstudio.ai/) running locally with:
  - A model loaded (recommended: Nemotron-3-Nano 4B)
  - Local server started on `localhost:1234`
  - CORS enabled in server settings

## Setup

```bash
nvm use 22
npm install
npm start
```

This builds the add-in and sideloads it into Excel. Use `=EXCELAI.AI("your prompt")` in any cell.

## How it works

The `=AI()` custom function sends prompts to LM Studio's OpenAI-compatible API (`localhost:1234/v1/chat/completions`) and returns the response text into the cell. Results are cached in memory to avoid re-querying on recalculation.
