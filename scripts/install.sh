#!/usr/bin/env bash
set -e

ADDIN_ID="5cc82bad-f723-4c10-82ba-f135772ad04f"
BASE_URL="https://loritira.github.io/excel-ai"
DEFAULT_MODEL="qwen2.5:1.5b"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
CONTAINER_DIR="$HOME/Library/Containers/com.microsoft.Excel"
OLLAMA_DIR="$HOME/.ollama"
INSTALL_DIR="$HOME/.excelai"
PLIST_LABEL="com.excelai.server"
PLIST_PATH="$HOME/Library/LaunchAgents/$PLIST_LABEL.plist"
SERVER_PORT=11435

# Parse arguments
while [[ $# -gt 0 ]]; do
  case "$1" in
    --url) BASE_URL="$2"; shift 2 ;;
    --model) DEFAULT_MODEL="$2"; shift 2 ;;
    *) echo "Unknown option: $1"; exit 1 ;;
  esac
done

echo "Installing Excel AI..."
echo "Source: $BASE_URL"
echo ""

# --- 1. Check Excel has been opened ---

if [ ! -d "$CONTAINER_DIR" ]; then
  echo "Error: Excel container directory not found."
  echo "Please open Microsoft Excel at least once, then run this script again."
  exit 1
fi

# --- 2. Install Ollama ---

if command -v ollama &>/dev/null; then
  echo "[1/5] Ollama is already installed."
else
  echo "[1/5] Installing Ollama..."
  # The Ollama installer may fail at its "Starting Ollama..." step (e.g. open -a Ollama
  # can't find the app). That's fine — we manage Ollama's lifecycle ourselves.
  curl -fsSL https://ollama.com/install.sh | sh || true
  if ! command -v ollama &>/dev/null; then
    echo "Error: Ollama installation failed — 'ollama' not found on PATH."
    exit 1
  fi
fi

# --- 3. Generate TLS certificate ---

echo "[2/5] Setting up TLS certificate..."
mkdir -p "$OLLAMA_DIR"

if [ ! -f "$OLLAMA_DIR/cert.pem" ]; then
  openssl req -x509 -newkey rsa:2048 \
    -keyout "$OLLAMA_DIR/key.pem" \
    -out "$OLLAMA_DIR/cert.pem" \
    -days 3650 -nodes \
    -subj "/CN=Excel AI Local" \
    -addext "subjectAltName=IP:127.0.0.1" 2>/dev/null

  echo "  Adding certificate to system trust store (may require your password)..."
  sudo security add-trusted-cert -d -r trustRoot \
    -k /Library/Keychains/System.keychain "$OLLAMA_DIR/cert.pem"
fi

# --- 4. Download server binary and manifest ---

echo "[3/5] Downloading Excel AI server..."
mkdir -p "$INSTALL_DIR"

# Detect architecture
ARCH=$(uname -m)
case "$ARCH" in
  arm64|aarch64) BINARY="excelai-server-darwin-arm64" ;;
  x86_64)        BINARY="excelai-server-darwin-amd64" ;;
  *) echo "Unsupported architecture: $ARCH"; exit 1 ;;
esac

curl -fsSL "$BASE_URL/$BINARY" -o "$INSTALL_DIR/excelai-server"
chmod +x "$INSTALL_DIR/excelai-server"

# Save the resolved ollama path so the server can find it under launchd's minimal PATH
OLLAMA_BIN="$(command -v ollama)"
echo "$OLLAMA_BIN" > "$INSTALL_DIR/ollama-path"

# --- 5. Install manifest and configure auto-start ---

echo "[4/5] Installing add-in manifest and auto-start..."
mkdir -p "$WEF_DIR"
curl -fsSL "$BASE_URL/manifest-local.xml" -o "$WEF_DIR/$ADDIN_ID.manifest.xml"

# Clean up any old Ollama HTTPS config
launchctl unsetenv OLLAMA_HOST 2>/dev/null || true

# Disable Ollama's own auto-start (our proxy manages Ollama lifecycle)
OLLAMA_PLIST="$HOME/Library/LaunchAgents/com.ollama.ollama.plist"
if [ -f "$OLLAMA_PLIST" ]; then
  launchctl unload "$OLLAMA_PLIST" 2>/dev/null || true
  rm -f "$OLLAMA_PLIST"
fi
pkill -x ollama 2>/dev/null || true
pkill -x Ollama 2>/dev/null || true

# Create launchd plist for auto-start
mkdir -p "$HOME/Library/LaunchAgents"
cat > "$PLIST_PATH" << PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>${PLIST_LABEL}</string>
    <key>ProgramArguments</key>
    <array>
        <string>${INSTALL_DIR}/excelai-server</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>${INSTALL_DIR}/server.log</string>
    <key>StandardErrorPath</key>
    <string>${INSTALL_DIR}/server.log</string>
</dict>
</plist>
PLIST

launchctl unload "$PLIST_PATH" 2>/dev/null || true
launchctl load "$PLIST_PATH"

# --- 6. Pull the default model ---

# Start Ollama temporarily for model pull (the proxy manages it at runtime)
ollama serve &>/dev/null &
OLLAMA_PID=$!
sleep 3

echo "[5/5] Pulling model '$DEFAULT_MODEL' (this may take a minute)..."
ollama pull "$DEFAULT_MODEL"

kill $OLLAMA_PID 2>/dev/null || true
wait $OLLAMA_PID 2>/dev/null || true

echo ""
echo "Excel AI installed successfully!"
echo ""
echo "Next steps:"
echo "  1. Quit and reopen Microsoft Excel"
echo "  2. Use =EXCELAI.AI(\"your prompt\") in any cell"
echo ""
echo "Everything runs locally — your data never leaves your machine."
echo "Works offline after this initial setup."
