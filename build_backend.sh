#!/bin/bash
# Build the zotellm backend binary using PyInstaller.
# Run from the zotellm project root directory.
#
# Usage:
#   chmod +x build_backend.sh
#   ./build_backend.sh
#
# Output: dist/zotellm_backend/zotellm_backend

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "Building zotellm backend with PyInstaller..."

python -m PyInstaller \
    --onedir \
    --name zotellm_backend \
    --noconfirm \
    bridge.py

echo ""
echo "Build complete: dist/zotellm_backend/zotellm_backend"
echo ""
echo "Next steps:"
echo "  cd desktop && npm install"
echo "  npm start          # dev mode"
echo "  npm run build      # package .dmg"
