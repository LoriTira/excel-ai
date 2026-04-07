#!/usr/bin/env bash
set -e

ADDIN_ID="5cc82bad-f723-4c10-82ba-f135772ad04f"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
MANIFEST="$WEF_DIR/$ADDIN_ID.manifest.xml"
INSTALL_DIR="$HOME/.excelai"
OLLAMA_DIR="$HOME/.ollama"
PLIST_LABEL="com.excelai.server"
PLIST_PATH="$HOME/Library/LaunchAgents/$PLIST_LABEL.plist"

echo "Uninstalling Excel AI (everything)..."
echo ""

# --- 1. Stop and remove the Excel AI server ---

if [ -f "$PLIST_PATH" ]; then
  launchctl unload "$PLIST_PATH" 2>/dev/null || true
  rm "$PLIST_PATH"
  echo "[1/5] Stopped Excel AI server."
else
  echo "[1/5] Excel AI server not running."
fi

if [ -d "$INSTALL_DIR" ]; then
  rm -rf "$INSTALL_DIR"
  echo "  Removed $INSTALL_DIR"
fi

# --- 2. Remove add-in manifest ---

if [ -f "$MANIFEST" ]; then
  rm "$MANIFEST"
  echo "[2/5] Removed add-in manifest."
else
  echo "[2/5] Manifest already removed."
fi

# --- 3. Stop and remove Ollama ---

echo "[3/5] Removing Ollama..."

# Quit the Ollama app if running
osascript -e 'quit app "Ollama"' 2>/dev/null || true
pkill -f ollama 2>/dev/null || true
sleep 1

# Remove the Ollama binary
if [ -f /usr/local/bin/ollama ]; then
  sudo rm -f /usr/local/bin/ollama
  echo "  Removed /usr/local/bin/ollama"
fi

# Remove the Ollama app if installed via .dmg
if [ -d "/Applications/Ollama.app" ]; then
  sudo rm -rf "/Applications/Ollama.app"
  echo "  Removed /Applications/Ollama.app"
fi

# --- 4. Remove all models, certs, and data ---

echo "[4/5] Removing models and data..."

if [ -d "$OLLAMA_DIR" ]; then
  rm -rf "$OLLAMA_DIR"
  echo "  Removed $OLLAMA_DIR (models, certs, config)"
fi

# --- 5. Remove trusted certificate and env vars ---

echo "[5/5] Cleaning up..."

# Remove the trusted cert from the system keychain
sudo security delete-certificate -c "Excel AI Local" /Library/Keychains/System.keychain 2>/dev/null || true

launchctl unsetenv OLLAMA_HOST 2>/dev/null || true
launchctl unsetenv OLLAMA_ORIGINS 2>/dev/null || true

echo ""
echo "Excel AI has been completely uninstalled."
echo "Restart Excel for the change to take effect."
