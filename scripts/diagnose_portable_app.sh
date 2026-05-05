#!/bin/zsh
# Diagnose a built portable app.
set -euo pipefail
APP_PATH="${1:-dist-portable/MarkItDown Handy Portable.app}"
if [ ! -d "$APP_PATH" ]; then
  echo "App not found: $APP_PATH"
  echo "Usage: ./scripts/diagnose_portable_app.sh \"dist-portable/MarkItDown Handy Portable.app\""
  exit 1
fi
RES="$APP_PATH/Contents/Resources"
ENV="$RES/env"
export MARKITDOWN_BUNDLED_ENV="$ENV"
export MARKITDOWN_RESOURCE_DIR="$RES"
export PYTHONNOUSERSITE=1
export PATH="$ENV/bin:/usr/bin:/bin:/usr/sbin:/sbin"
export DYLD_LIBRARY_PATH="$ENV/lib:${DYLD_LIBRARY_PATH:-}"
[ -d "$RES/tessdata" ] && export TESSDATA_PREFIX="$RES/tessdata"

echo "App: $APP_PATH"
echo "Env: $ENV"
ls -lh "$ENV/bin/python" "$ENV/bin/markitdown" "$ENV/bin/ocrmypdf" 2>/dev/null || true
"$ENV/bin/python" - <<'PYCHECK'
import importlib.util, sys
print("python:", sys.executable)
for mod in ["markitdown", "ocrmypdf", "tkinterdnd2"]:
    print(mod, "OK" if importlib.util.find_spec(mod) else "MISSING")
PYCHECK
"$ENV/bin/markitdown" --version || true
"$ENV/bin/ocrmypdf" --version || true
"$ENV/bin/tesseract" --list-langs | head -20 || true
