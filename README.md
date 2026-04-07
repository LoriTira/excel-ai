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

Removes everything: the add-in, the server binary, Ollama, all models, TLS certificates, and environment variables.

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
Excel Add-in (HTTPS) <-> Excel AI server (HTTPS :11435) <-> Ollama (HTTP :11434)
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

---

## Technical reference

> Everything below is written for developers (or AI agents) who will maintain or extend this project. It documents every component, design decision, gotcha, and lesson learned.

### Architecture overview

```
+------------------+     HTTPS :11435     +------------------+     HTTP :11434     +------------------+
|                  |  ================>   |                  |  ================>  |                  |
|  Excel Add-in   |      same-origin     |  Go server       |      reverse proxy  |  Ollama          |
|  (taskpane +    |  <================   |  (excelai-server) |  <================  |  (llama3.2:1b)   |
|   custom fns)   |     static files     |  ~8 MB binary    |      /v1/* /api/*   |  port 11434      |
+------------------+                     +------------------+                     +------------------+
```

**Why a Go server?** Office Add-ins must be served over HTTPS. Ollama only speaks HTTP. A Python proxy was tried first but required a Python runtime. The Go binary is ~8 MB, uses ~14 MB RSS at runtime, has zero dependencies, and embeds all static files via `go:embed`. It serves both the add-in UI and proxies API requests on the same origin, eliminating mixed-content and CORS issues entirely.

**Why same-origin?** When the add-in and API share the same `https://127.0.0.1:11435` origin, the browser allows `fetch()` calls without CORS headers. This means Ollama doesn't need any CORS configuration, and the system works even when Ollama doesn't set CORS headers.

### File structure

```
excel-ai/
├── .github/workflows/deploy.yml   # CI: build + cross-compile + deploy to GitHub Pages
├── assets/                        # Add-in icons (16, 32, 64, 80px)
├── scripts/
│   ├── install.sh                 # macOS installer (Ollama + cert + server + manifest)
│   ├── install.ps1                # Windows installer
│   ├── uninstall.sh               # macOS full removal
│   ├── uninstall.ps1              # Windows full removal
│   └── bench.sh                   # Model benchmarking script
├── server/
│   ├── main.go                    # Go HTTPS reverse proxy + static file server
│   ├── go.mod                     # Go module (no dependencies)
│   └── static/                    # (gitignored) webpack output copied here before Go build
├── src/
│   ├── functions/
│   │   ├── functions.ts           # =EXCELAI.AI() custom function implementation
│   │   └── functions.html         # Custom functions runtime HTML loader
│   ├── services/
│   │   ├── ai.ts                  # AI completion logic (local + API modes)
│   │   ├── settings.ts            # Settings persistence (localStorage)
│   │   ├── cache.ts               # In-memory response cache (Map)
│   │   └── queue.ts               # Concurrent request limiter (max 3)
│   ├── taskpane/
│   │   ├── taskpane.html          # Settings panel UI
│   │   └── taskpane.ts            # Settings form logic + model detection
│   ├── install/
│   │   └── install.html           # Platform-specific install instructions page
│   └── legal/
│       ├── privacy.html
│       └── terms.html
├── manifest.xml                   # Office Add-in XML manifest
├── webpack.config.js              # Webpack bundler config (dev proxy, manifest transforms)
├── package.json                   # npm scripts and dependencies
├── tsconfig.json                  # TypeScript config (target: ES5)
├── babel.config.json              # Babel config (ES5 transpilation)
├── .eslintrc.json                 # ESLint with office-addins plugin
├── .nvmrc                         # Node.js version: 22
└── start.sh                       # Quick-start dev script
```

### Component details

#### Custom function (`src/functions/functions.ts`)

The `=EXCELAI.AI(prompt, [context], [model])` function:

- **prompt** (string): The instruction to send to the AI
- **context** (range, optional): Cell values serialized as tab-separated rows, joined by newlines
- **model** (string, optional): Override the auto-detected Ollama model per cell

Flow: `serialize context -> build cache key -> check cache -> enqueue request -> call AI -> truncate to 32,767 chars (Excel limit) -> cache and return`

Errors are returned as `#ERROR: <message>` strings (not thrown), so they appear in the cell.

#### AI service (`src/services/ai.ts`)

Two modes:
- **Local mode** (`provider === "local"`): Uses relative URLs (e.g., `/v1/chat/completions`), which resolve against the same origin in production (Go server) or hit the webpack dev proxy in development
- **API mode** (`provider === "api"`): Uses the full endpoint URL with Bearer token auth

Key parameters sent to Ollama:
- `temperature: 0.3` — low for deterministic, cell-appropriate answers
- `max_tokens: 512` — keeps responses concise
- `num_ctx: 2048` — **critical for RAM**: limits Ollama's KV cache. Without this, models with 128K default context allocate 5–18 GB of RAM. With 2048, even large models fit in ~1.5 GB. This is only sent in local mode.
- `model`: **required by Ollama** — the `/v1/chat/completions` endpoint returns `{"error":{"message":"model is required"}}` if omitted. The model is auto-detected via `/v1/models`.

The system prompt is: `"You are a helpful assistant embedded in Excel. Give concise answers suitable for spreadsheet cells."`

Auto-detection: `detectLocalModel()` calls `GET /v1/models`, takes the first model's `id`, and caches it in a module-level variable. This avoids a round-trip on every cell evaluation.

#### Settings (`src/services/settings.ts`)

Stored in `localStorage` under key `excelai_settings`. Schema:

```typescript
interface ProviderSettings {
  provider: "local" | "api";
  localAddress: string;   // empty = same-origin (default), or a custom URL
  apiEndpoint: string;    // e.g., "https://api.openai.com"
  apiKey: string;         // e.g., "sk-..."
  apiModel: string;       // e.g., "gpt-4o-mini"
}
```

There is no `localModel` field — the model is auto-detected from Ollama at runtime. This was a deliberate design choice: users who change models via `ollama pull`/`ollama run` should not also need to update a setting.

#### Queue (`src/services/queue.ts`)

- Max 3 concurrent requests, 100 pending queue depth
- Prevents Ollama overload when many cells recalculate simultaneously
- FIFO execution with promise-based waiting

#### Cache (`src/services/cache.ts`)

- Simple `Map<string, string>` keyed by `"prompt||serialized_context"`
- Cleared when settings are saved (to pick up model/endpoint changes)
- Lives in memory only — does not survive Excel restarts

#### Go server (`server/main.go`)

A minimal ~60-line Go program:
- Embeds `server/static/` into the binary via `//go:embed static`
- Routes `/v1/*` and `/api/*` to Ollama at `http://127.0.0.1:11434` via `httputil.NewSingleHostReverseProxy`
- Serves all other paths from embedded static files
- Loads TLS cert/key from `~/.ollama/cert.pem` and `~/.ollama/key.pem`
- Listens on `127.0.0.1:11435` (localhost only)
- Zero external Go dependencies (only stdlib)

**Building**: `go build -ldflags="-s -w"` strips debug symbols, producing an ~8 MB binary. The `server/static/` directory must be populated before building (the CI pipeline does this).

#### Manifest (`manifest.xml`)

- **Add-in ID**: `5cc82bad-f723-4c10-82ba-f135772ad04f`
- **Namespace**: `EXCELAI` (so functions are `=EXCELAI.AI(...)`)
- **Runtime**: SharedRuntime with `lifetime="long"` — the taskpane and custom functions share the same JavaScript runtime, which means localStorage settings are shared
- **Host**: Workbook only (Excel desktop)
- **Ribbon**: "Excel AI" button in the Home tab that opens the settings taskpane

Two manifest variants are built:
- `manifest.xml` — production URLs pointing to GitHub Pages
- `manifest-local.xml` — all URLs point to `https://127.0.0.1:11435/`

The local manifest is what gets installed by the install scripts.

#### Taskpane UI (`src/taskpane/`)

- Radio toggle between Local and API modes
- On load, queries `/v1/models` to display the active Ollama model name
- Advanced section (collapsed by default): custom server address
- Troubleshooting section with copy-able commands
- Dark mode via `prefers-color-scheme: dark` media query
- Save button clears cache and runs a connection test

### Webpack config (`webpack.config.js`)

Key behaviors:

- **Dev proxy**: `/v1` and `/api` requests are proxied to `http://127.0.0.1:11434` (Ollama), configurable via `OLLAMA_URL` env var
- **Manifest transform**: In production builds, replaces `https://localhost:3000/` URLs with `https://loritira.github.io/excel-ai/`
- **Local manifest**: An additional copy transforms the manifest to use `https://127.0.0.1:11435/` URLs for the installed server
- **Static copies**: Install scripts, uninstall scripts, legal pages, and icons are copied to `dist/`

### CI/CD (`.github/workflows/deploy.yml`)

On push to `main`:
1. `npm ci && npm run build` — webpack production build
2. Copy `dist/` to `server/static/` — embed UI in Go binary
3. Cross-compile Go binary for 4 targets:
   - `darwin-arm64` (macOS Apple Silicon)
   - `darwin-amd64` (macOS Intel)
   - `windows-amd64` (Windows x64)
   - `windows-arm64` (Windows ARM)
4. All binaries placed in `dist/`
5. Deploy `dist/` to GitHub Pages

Everything is served from `https://loritira.github.io/excel-ai/` — the install scripts download binaries and manifests from there.

### Install scripts

#### macOS (`scripts/install.sh`)

Steps:
1. Verify `~/Library/Containers/com.microsoft.Excel` exists (Excel must have been opened once)
2. Install Ollama via `curl -fsSL https://ollama.com/install.sh | sh` (if not present)
3. Generate TLS cert with OpenSSL (10-year, CN=Excel AI Local, SAN=IP:127.0.0.1), add to system keychain via `sudo security add-trusted-cert` (requires password)
4. Download arch-specific Go binary (`uname -m` -> arm64 or x86_64)
5. Download `manifest-local.xml` to wef directory: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/<addin-id>.manifest.xml`
6. Create launchd plist at `~/Library/LaunchAgents/com.excelai.server.plist` (RunAtLoad + KeepAlive)
7. Start Ollama if not running
8. Pull default model (`llama3.2:1b`, or custom via `--model` flag)

Flags: `--url <base>` (custom download URL), `--model <name>` (custom model)

#### Windows (`scripts/install.ps1`)

Steps:
1. Install Ollama via winget or direct download of `OllamaSetup.exe` (`/VERYSILENT /NORESTART`)
2. Generate TLS cert with `New-SelfSignedCertificate` (10-year, CN=Excel AI Local, SAN=IP:127.0.0.1), add to `CurrentUser\Root` store (no admin needed)
3. Export cert and private key as PEM files to `~\.ollama\` (Go server reads these)
4. Download arch-specific binary (`$env:PROCESSOR_ARCHITECTURE` -> ARM64 or default AMD64)
5. Register manifest via registry: `HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\<addin-id>` = path to `manifest.xml`
6. Create Scheduled Task `ExcelAI-Server` (AtLogOn trigger, no battery/time restrictions)
7. Start server process immediately
8. Pull default model

**`$ProgressPreference = "SilentlyContinue"`**: This is critical on Windows. Without it, `Invoke-WebRequest` renders a text progress bar that causes 10–50x slowdown on downloads. Both install and uninstall scripts set this at the top.

#### Uninstall scripts

Both scripts remove **everything**:
1. Stop and remove the Excel AI server (launchd plist / scheduled task)
2. Remove add-in manifest (wef file / registry key)
3. Stop and uninstall Ollama (app, binary, via winget or direct uninstaller)
4. Remove `~/.ollama` directory (models, certs, config)
5. Remove trusted TLS certificate from keychain/cert store
6. Clean up environment variables (`OLLAMA_HOST`, `OLLAMA_ORIGINS`)

### Platform-specific details

#### macOS sideloading

Excel on macOS discovers add-ins from the wef directory:
```
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

The manifest file must be named `<addin-id>.manifest.xml`. The container directory only exists after Excel has been opened at least once — the install script checks for this.

#### Windows sideloading

Excel on Windows discovers developer add-ins from the registry:
```
HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer
```

Each entry is a string property where the name is the add-in ID and the value is the full path to the manifest XML file.

#### TLS certificates

**macOS**: Generated with `openssl`, trusted via `sudo security add-trusted-cert` into the system keychain (`/Library/Keychains/System.keychain`). Requires the user's password during install. Removed via `sudo security delete-certificate -c "Excel AI Local"`.

**Windows**: Generated with `New-SelfSignedCertificate`, added to `Cert:\CurrentUser\Root` (no admin needed). The cert is created in `Cert:\CurrentUser\My`, exported as PEM, then moved to Root and removed from My. The PEM export is necessary because Go's `tls.LoadX509KeyPair` reads PEM files, not Windows cert stores.

**Both platforms**: Cert is stored as `~/.ollama/cert.pem` and `~/.ollama/key.pem`. The `.ollama` directory was chosen because it already exists from the Ollama install — no extra directories needed.

### Gotchas and lessons learned

1. **Ollama requires the `model` field** in `/v1/chat/completions` requests. Unlike OpenAI, omitting it returns `{"error":{"message":"model is required"}}`. The add-in auto-detects the model via `GET /v1/models` before making any completion request.

2. **`num_ctx` is critical for RAM**. Ollama defaults to the model's training context (often 128K tokens), which allocates a massive KV cache. A 1B model can use 18 GB of RAM at 128K context. Setting `num_ctx: 2048` drops this to ~1.5 GB. Cell prompts are short, so 2K tokens is more than enough.

3. **PowerShell `Invoke-WebRequest` is extremely slow by default**. The progress bar rendering causes 10–50x slowdown. Always set `$ProgressPreference = "SilentlyContinue"` at the top of any PowerShell script that downloads files.

4. **`manifest-local.xml` URL replacement must handle both dev and prod URLs**. The webpack `CopyWebpackPlugin` transform replaces `https://localhost:3000/` (dev) AND `https://loritira.github.io/excel-ai/` (prod) with `https://127.0.0.1:11435/`. In production builds, only prod URLs are in the source manifest, so replacing only dev URLs would leave prod URLs in place.

5. **SharedRuntime is required for custom functions + taskpane to share state**. The manifest uses `<Runtime lifetime="long"/>` so the taskpane and custom function runtime share the same JavaScript context (and therefore the same `localStorage`). Without this, settings changes in the taskpane wouldn't affect function behavior.

6. **qwen3.5 models are incompatible**. They return an empty `content` field and put their response in a non-standard `reasoning` field. The `/no_think` flag doesn't fix this. Do not use qwen3.5 models with this add-in.

7. **RAM measurements between models are unreliable if done sequentially**. Ollama doesn't fully unload a model's KV cache when loading a new one. To get accurate RAM measurements, each model must be tested in isolation with an explicit unload between tests.

8. **macOS requires Excel to have been opened once** before sideloading. The `~/Library/Containers/com.microsoft.Excel/` directory is created on first launch. The install script checks for this and exits with a helpful message if missing.

9. **Go `embed.FS` requires the directory to exist at build time**. The `//go:embed static` directive fails if `server/static/` doesn't exist. The CI pipeline creates it by copying webpack output before the Go build step.

10. **The Go binary is self-contained**: no runtime dependencies, no config files (besides TLS certs), no environment variables. It just needs `cert.pem`/`key.pem` to exist and Ollama to be running on port 11434.

11. **Windows cert export to PEM**: PowerShell can export `.pfx` natively but Go needs PEM. The install script exports the certificate bytes and private key separately, Base64-encodes them, and wraps them in `-----BEGIN CERTIFICATE-----` / `-----BEGIN PRIVATE KEY-----` headers. The private key is exported via `ExportPkcs8PrivateKey()`.

12. **Ollama auto-starts on Windows after install**. The install script doesn't need to explicitly start it. On macOS, the script checks `pgrep -x ollama` and starts it if needed.

13. **The Go server only listens on 127.0.0.1**, not `0.0.0.0`. This is intentional — the server should never be accessible from the network. The self-signed cert is only valid for `127.0.0.1`.
