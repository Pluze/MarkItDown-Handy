#!/bin/zsh
# Run the app from source using the local conda environment.
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
ENV_NAME="${ENV_NAME:-markitdown}"

if [ -f "/opt/anaconda3/etc/profile.d/conda.sh" ]; then
  source "/opt/anaconda3/etc/profile.d/conda.sh"
elif [ -f "$HOME/anaconda3/etc/profile.d/conda.sh" ]; then
  source "$HOME/anaconda3/etc/profile.d/conda.sh"
elif [ -f "$HOME/miniconda3/etc/profile.d/conda.sh" ]; then
  source "$HOME/miniconda3/etc/profile.d/conda.sh"
else
  echo "Cannot find conda.sh."
  exit 1
fi
conda activate "$ENV_NAME"
python "$ROOT_DIR/src/markitdown_handy.py"
