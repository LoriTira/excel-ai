#!/usr/bin/env bash
set -e

ADDIN_ID="5cc82bad-f723-4c10-82ba-f135772ad04f"
BASE_URL="https://loritira.github.io/excel-ai"
DEFAULT_MODEL="qwen2.5:1.5b"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
CONTAINER_DIR="$HOME/Library/Containers/com.microsoft.Excel"

# Parse arguments
while [[ $# -gt 0 ]]; do
  case "$1" in
    --url) BASE_URL="$2"; shift 2 ;;
    --model) DEFAULT_MODEL="$2"; shift 2 ;;
    *) echo "Unknown option: $1"; exit 1 ;;
  esac
done

echo "Installing Excel AI add-in..."
echo "Source: $BASE_URL"
echo ""

# --- 1. Install the Excel add-in manifest ---

# Check that Excel has been opened at least once
if [ ! -d "$CONTAINER_DIR" ]; then
  echo "Error: Excel container directory not found."
  echo "Please open Microsoft Excel at least once, then run this script again."
  exit 1
fi

# Create wef directory if needed
mkdir -p "$WEF_DIR"

# Download the production manifest
curl -fsSL "$BASE_URL/manifest.xml" -o "$WEF_DIR/$ADDIN_ID.manifest.xml"
echo "[1/3] Excel add-in manifest installed."

# --- 2. Install Ollama ---

if command -v ollama &>/dev/null; then
  echo "[2/3] Ollama is already installed."
else
  echo "[2/3] Installing Ollama..."
  curl -fsSL https://ollama.com/install.sh | sh
fi

# Allow the hosted add-in to connect to Ollama (CORS)
launchctl setenv OLLAMA_ORIGINS "*"

# Restart Ollama so it picks up the new OLLAMA_ORIGINS setting
pkill -f ollama 2>/dev/null || true
sleep 1
ollama serve &>/dev/null &
sleep 2

# --- 3. Pull the default model ---

echo "[3/3] Pulling model '$DEFAULT_MODEL' (this may take a minute)..."
ollama pull "$DEFAULT_MODEL"

echo ""
echo "Excel AI installed successfully!"
echo ""
echo "Next steps:"
echo "  1. Quit and reopen Microsoft Excel"
echo "  2. Use =EXCELAI.AI(\"your prompt\") in any cell"
echo ""
echo "Ollama is running locally — your data never leaves your machine."
