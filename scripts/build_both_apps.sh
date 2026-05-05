#!/bin/zsh
# Build local packages supported by the current platform.
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

if [ "$(uname -s)" = "Darwin" ]; then
  "$SCRIPT_DIR/build_conda_wrapper_app.sh"
fi

python3 "$SCRIPT_DIR/build_cross_platform_portable.py"
echo "Available builds finished."
