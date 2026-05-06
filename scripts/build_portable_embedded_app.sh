#!/bin/zsh
# Build a portable macOS app with embedded Python, MarkItDown, OCRmyPDF, and OCR tools.

set -euo pipefail

APP_NAME="MarkItDown Handy Portable"
APP_VERSION="0.1.0"
ENV_NAME="${ENV_NAME:-markitdown-handy-release}"
PY_VERSION="${PY_VERSION:-3.11}"

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
SRC="$ROOT_DIR/src/markitdown_handy.py"
BUILD_DIR="$ROOT_DIR/build-portable"
OUT_DIR="$ROOT_DIR/dist-portable"
APP_DIR="$OUT_DIR/$APP_NAME.app"
RES_DIR="$APP_DIR/Contents/Resources"
MACOS_DIR="$APP_DIR/Contents/MacOS"
ENV_DIR="$RES_DIR/env"
TESSDATA_DIR="$RES_DIR/tessdata"
ARCH="$(uname -m)"

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
  echo "Cannot find conda.sh. Portable builds require conda on the build machine."
  exit 1
fi
source "$CONDA_SH"

mkdir -p "$BUILD_DIR" "$OUT_DIR"

if conda env list | awk '{print $1}' | grep -qx "$ENV_NAME"; then
  echo "Reusing conda environment: $ENV_NAME"
else
  conda create -y -n "$ENV_NAME" -c conda-forge "python=$PY_VERSION" pip
fi

# setup-miniconda may pre-create the environment with only Python. Always make sure
# the native OCR/PDF tools and conda-pack are present before packaging.
# OCRmyPDF itself is installed with pip because conda-forge's osx-arm64 solve can
# require optional pngquant/unpaper packages that are not always available.
conda install -y -n "$ENV_NAME" -c conda-forge \
  conda-pack ffmpeg tesseract ghostscript qpdf

conda activate "$ENV_NAME"
# Keep pip itself conda-managed. Upgrading pip with pip clobbers conda-managed
# files and makes conda-pack fail before the app can be packaged.
python -m pip install --no-cache-dir 'markitdown[all]' tkinterdnd2 ocrmypdf

python - <<'PYCHECK'
import importlib.util
missing = [m for m in ["markitdown", "ocrmypdf", "tkinterdnd2"] if not importlib.util.find_spec(m)]
if missing:
    raise SystemExit(f"Missing required modules: {missing}")
print("Python module check OK")
PYCHECK

rm -rf "$APP_DIR"
mkdir -p "$RES_DIR" "$MACOS_DIR"
cp "$SRC" "$RES_DIR/markitdown_handy.py"

cat > "$APP_DIR/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple Computer//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
  <dict>
    <key>CFBundleName</key><string>$APP_NAME</string>
    <key>CFBundleDisplayName</key><string>$APP_NAME</string>
    <key>CFBundleIdentifier</key><string>com.ziyuzhu.markitdown-handy.portable</string>
    <key>CFBundleVersion</key><string>$APP_VERSION</string>
    <key>CFBundleShortVersionString</key><string>$APP_VERSION</string>
    <key>CFBundleExecutable</key><string>MarkItDownHandyPortable</string>
    <key>LSMinimumSystemVersion</key><string>11.0</string>
    <key>NSHighResolutionCapable</key><true/>
  </dict>
</plist>
PLIST

cat > "$MACOS_DIR/MarkItDownHandyPortable" <<'LAUNCHER'
#!/bin/zsh
APP_DIR="$(cd "$(dirname "$0")/../.." && pwd)"
RESOURCE_DIR="$APP_DIR/Contents/Resources"
ENV_DIR="$RESOURCE_DIR/env"

export MARKITDOWN_BUNDLED_ENV="$ENV_DIR"
export MARKITDOWN_RESOURCE_DIR="$RESOURCE_DIR"
export PYTHONNOUSERSITE=1
export PATH="$ENV_DIR/bin:/usr/bin:/bin:/usr/sbin:/sbin"
export DYLD_LIBRARY_PATH="$ENV_DIR/lib:${DYLD_LIBRARY_PATH:-}"

if [ -d "$RESOURCE_DIR/tessdata" ]; then
  export TESSDATA_PREFIX="$RESOURCE_DIR/tessdata"
fi

exec "$ENV_DIR/bin/python" "$RESOURCE_DIR/markitdown_handy.py"
LAUNCHER
chmod +x "$MACOS_DIR/MarkItDownHandyPortable"

ENV_TAR="$BUILD_DIR/${ENV_NAME}_${ARCH}.tar.gz"
rm -f "$ENV_TAR"
conda-pack -n "$ENV_NAME" -o "$ENV_TAR"

rm -rf "$ENV_DIR"
mkdir -p "$ENV_DIR"
tar -xzf "$ENV_TAR" -C "$ENV_DIR"
"$ENV_DIR/bin/conda-unpack" || true

if [ ! -x "$ENV_DIR/bin/markitdown" ]; then
  cat > "$ENV_DIR/bin/markitdown" <<'WRAP'
#!/bin/zsh
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec "$SCRIPT_DIR/python" -m markitdown "$@"
WRAP
  chmod +x "$ENV_DIR/bin/markitdown"
fi

if [ ! -x "$ENV_DIR/bin/ocrmypdf" ]; then
  cat > "$ENV_DIR/bin/ocrmypdf" <<'WRAP'
#!/bin/zsh
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec "$SCRIPT_DIR/python" -m ocrmypdf "$@"
WRAP
  chmod +x "$ENV_DIR/bin/ocrmypdf"
fi

mkdir -p "$TESSDATA_DIR"
copy_lang() {
  local base="$1"
  local lang="$2"
  if [ -f "$base/$lang.traineddata" ]; then
    cp "$base/$lang.traineddata" "$TESSDATA_DIR/"
  fi
}
for base in "$CONDA_PREFIX/share/tessdata" "/opt/homebrew/share/tessdata" "/usr/local/share/tessdata"; do
  if [ -d "$base" ]; then
    for lang in osd eng chi_sim chi_tra; do
      copy_lang "$base" "$lang"
    done
  fi
done

export MARKITDOWN_BUNDLED_ENV="$ENV_DIR"
export MARKITDOWN_RESOURCE_DIR="$RES_DIR"
export PYTHONNOUSERSITE=1
export PATH="$ENV_DIR/bin:/usr/bin:/bin:/usr/sbin:/sbin"
export DYLD_LIBRARY_PATH="$ENV_DIR/lib:${DYLD_LIBRARY_PATH:-}"
[ -d "$TESSDATA_DIR" ] && export TESSDATA_PREFIX="$TESSDATA_DIR"

"$ENV_DIR/bin/python" - <<'PYCHECK'
import importlib.util
for mod in ["markitdown", "ocrmypdf", "tkinterdnd2"]:
    if not importlib.util.find_spec(mod):
        raise SystemExit(f"Bundled module missing: {mod}")
print("Bundled module check OK")
PYCHECK
"$ENV_DIR/bin/markitdown" --version || true
"$ENV_DIR/bin/ocrmypdf" --version || true

ZIP_PATH="$OUT_DIR/MarkItDown_Handy_v${APP_VERSION}_portable_macOS_${ARCH}.zip"
rm -f "$ZIP_PATH"
cd "$OUT_DIR"
/usr/bin/ditto -c -k --sequesterRsrc --keepParent "$APP_NAME.app" "$ZIP_PATH"

echo "Built portable app: $APP_DIR"
echo "Zip: $ZIP_PATH"
