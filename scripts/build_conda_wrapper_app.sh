#!/bin/zsh
# Build a lightweight macOS app that uses an existing conda environment.

set -euo pipefail

APP_NAME="MarkItDown Handy"
APP_VERSION="0.1.0"
ENV_NAME="${ENV_NAME:-markitdown}"

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
SRC="$ROOT_DIR/src/markitdown_handy.py"
OUT_DIR="$ROOT_DIR/dist-conda"
APP_DIR="$OUT_DIR/$APP_NAME.app"
RES_DIR="$APP_DIR/Contents/Resources"
MACOS_DIR="$APP_DIR/Contents/MacOS"
ARCH="$(uname -m)"

rm -rf "$APP_DIR"
mkdir -p "$RES_DIR" "$MACOS_DIR" "$OUT_DIR"
cp "$SRC" "$RES_DIR/markitdown_handy.py"

cat > "$APP_DIR/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple Computer//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
  <dict>
    <key>CFBundleName</key><string>$APP_NAME</string>
    <key>CFBundleDisplayName</key><string>$APP_NAME</string>
    <key>CFBundleIdentifier</key><string>com.ziyuzhu.markitdown-handy.conda</string>
    <key>CFBundleVersion</key><string>$APP_VERSION</string>
    <key>CFBundleShortVersionString</key><string>$APP_VERSION</string>
    <key>CFBundleExecutable</key><string>MarkItDownHandy</string>
    <key>LSMinimumSystemVersion</key><string>10.15</string>
    <key>NSHighResolutionCapable</key><true/>
  </dict>
</plist>
PLIST

cat > "$MACOS_DIR/MarkItDownHandy" <<'LAUNCHER'
#!/bin/zsh
RESOURCE_DIR="$(cd "$(dirname "$0")/../Resources" && pwd)"
ENV_NAME="${ENV_NAME:-markitdown}"

export PATH="/opt/anaconda3/bin:$HOME/anaconda3/bin:$HOME/miniconda3/bin:/opt/homebrew/bin:/usr/local/bin:/opt/homebrew/Caskroom/miniforge/base/bin:/usr/local/Caskroom/miniforge/base/bin:$PATH"

find_conda_sh() {
  for file in \
    "/opt/anaconda3/etc/profile.d/conda.sh" \
    "$HOME/anaconda3/etc/profile.d/conda.sh" \
    "$HOME/miniconda3/etc/profile.d/conda.sh" \
    "/opt/homebrew/Caskroom/miniforge/base/etc/profile.d/conda.sh" \
    "/usr/local/Caskroom/miniforge/base/etc/profile.d/conda.sh" \
    "/opt/homebrew/anaconda3/etc/profile.d/conda.sh" \
    "/usr/local/anaconda3/etc/profile.d/conda.sh"
  do
    if [ -f "$file" ]; then
      echo "$file"
      return 0
    fi
  done
  return 1
}

CONDA_SH="$(find_conda_sh || true)"
if [ -z "$CONDA_SH" ]; then
  osascript -e 'display dialog "Cannot find conda.sh. Install conda or run the Python source manually." buttons {"OK"} default button "OK"'
  exit 1
fi

source "$CONDA_SH"
conda activate "$ENV_NAME"
exec python "$RESOURCE_DIR/markitdown_handy.py"
LAUNCHER

chmod +x "$MACOS_DIR/MarkItDownHandy"

ZIP_PATH="$OUT_DIR/MarkItDown_Handy_v${APP_VERSION}_conda_wrapper_macOS_${ARCH}.zip"
rm -f "$ZIP_PATH"
cd "$OUT_DIR"
/usr/bin/ditto -c -k --sequesterRsrc --keepParent "$APP_NAME.app" "$ZIP_PATH"

echo "Built conda-wrapper app: $APP_DIR"
echo "Zip: $ZIP_PATH"
