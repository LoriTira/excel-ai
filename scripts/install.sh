#!/usr/bin/env bash
set -e

ADDIN_ID="5cc82bad-f723-4c10-82ba-f135772ad04f"
BASE_URL="https://loritira.github.io/excel-ai"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
CONTAINER_DIR="$HOME/Library/Containers/com.microsoft.Excel"

# Parse arguments
while [[ $# -gt 0 ]]; do
  case "$1" in
    --url) BASE_URL="$2"; shift 2 ;;
    *) echo "Unknown option: $1"; exit 1 ;;
  esac
done

echo "Installing Excel AI add-in..."
echo "Source: $BASE_URL"

# Check that Excel has been opened at least once
if [ ! -d "$CONTAINER_DIR" ]; then
  echo ""
  echo "Error: Excel container directory not found."
  echo "Please open Microsoft Excel at least once, then run this script again."
  exit 1
fi

# Create wef directory if needed
mkdir -p "$WEF_DIR"

# Download the production manifest
curl -fsSL "$BASE_URL/manifest.xml" -o "$WEF_DIR/$ADDIN_ID.manifest.xml"

echo ""
echo "Excel AI installed successfully!"
echo "  Manifest: $WEF_DIR/$ADDIN_ID.manifest.xml"
echo ""
echo "Next steps:"
echo "  1. Quit and reopen Microsoft Excel"
echo "  2. Open the Excel AI taskpane (Home tab > Excel AI)"
echo "  3. Configure your AI provider (LM Studio or external API)"
echo "  4. Use =EXCELAI.AI(\"your prompt\") in any cell"
