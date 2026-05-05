#!/bin/zsh
# Prepare or update the local conda development environment.
set -euo pipefail
ENV_NAME="${ENV_NAME:-markitdown}"

find_conda_sh() {
  for file in \
    "/opt/anaconda3/etc/profile.d/conda.sh" \
    "$HOME/anaconda3/etc/profile.d/conda.sh" \
    "$HOME/miniconda3/etc/profile.d/conda.sh" \
    "/opt/homebrew/Caskroom/miniforge/base/etc/profile.d/conda.sh" \
    "/opt/homebrew/anaconda3/etc/profile.d/conda.sh"
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
  echo "Cannot find conda.sh."
  exit 1
fi
source "$CONDA_SH"
conda activate "$ENV_NAME"
python -m pip install -U pip
python -m pip install -U 'markitdown[all]' tkinterdnd2 conda-pack
markitdown --version || true
python - <<'PYCHECK'
import importlib.util
for mod in ["markitdown", "tkinterdnd2"]:
    print(mod, "OK" if importlib.util.find_spec(mod) else "MISSING")
PYCHECK
