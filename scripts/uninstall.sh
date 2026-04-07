#!/usr/bin/env bash
set -e

ADDIN_ID="5cc82bad-f723-4c10-82ba-f135772ad04f"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
MANIFEST="$WEF_DIR/$ADDIN_ID.manifest.xml"

echo "Uninstalling Excel AI add-in..."

if [ -f "$MANIFEST" ]; then
  rm "$MANIFEST"
  echo "Removed: $MANIFEST"
else
  echo "Manifest not found at $MANIFEST (already uninstalled?)"
fi

echo ""
echo "Excel AI has been uninstalled."
echo "Restart Excel for the change to take effect."
echo ""
echo "Note: Ollama was not removed. To uninstall Ollama and its models:"
echo "  sudo rm -f /usr/local/bin/ollama"
echo "  rm -rf ~/.ollama"
