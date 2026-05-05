#!/bin/zsh
# Build both MarkItDown Handy app variants.
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
"$SCRIPT_DIR/build_conda_wrapper_app.sh"
"$SCRIPT_DIR/build_portable_embedded_app.sh"
echo "Both builds finished."
