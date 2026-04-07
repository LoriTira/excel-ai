#!/usr/bin/env bash
set -e

ADDIN_ID="5cc82bad-f723-4c10-82ba-f135772ad04f"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
MANIFEST="$WEF_DIR/$ADDIN_ID.manifest.xml"
INSTALL_DIR="$HOME/.excelai"
PLIST_LABEL="com.excelai.server"
PLIST_PATH="$HOME/Library/LaunchAgents/$PLIST_LABEL.plist"

echo "Uninstalling Excel AI..."

# Stop and remove the server
if [ -f "$PLIST_PATH" ]; then
  launchctl unload "$PLIST_PATH" 2>/dev/null || true
  rm "$PLIST_PATH"
  echo "Stopped Excel AI server."
fi

if [ -d "$INSTALL_DIR" ]; then
  rm -rf "$INSTALL_DIR"
  echo "Removed: $INSTALL_DIR"
fi

# Remove add-in manifest
if [ -f "$MANIFEST" ]; then
  rm "$MANIFEST"
  echo "Removed add-in manifest."
else
  echo "Manifest not found (already uninstalled?)"
fi

# Clean up environment variables
launchctl unsetenv OLLAMA_HOST 2>/dev/null || true
launchctl unsetenv OLLAMA_ORIGINS 2>/dev/null || true

echo ""
echo "Excel AI has been uninstalled."
echo "Restart Excel for the change to take effect."
echo ""
echo "Note: Ollama and its models were not removed. To uninstall Ollama:"
echo "  sudo rm -f /usr/local/bin/ollama"
echo "  rm -rf ~/.ollama"
echo ""
echo "To remove the trusted certificate:"
echo "  Open Keychain Access > System > find 'Excel AI Local' > delete it"
