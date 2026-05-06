#!/usr/bin/env python3
"""Dispatch the portable build to the platform-specific packager."""

from __future__ import annotations

import platform
import subprocess
import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
SCRIPT_DIR = ROOT_DIR / "scripts"


def run(cmd: list[str]) -> None:
    print("+", " ".join(cmd), flush=True)
    subprocess.run(cmd, cwd=ROOT_DIR, check=True)


def main() -> int:
    system = platform.system().lower()

    if system == "darwin":
        run(["zsh", str(SCRIPT_DIR / "build_portable_embedded_app.sh")])
        return 0

    if system == "windows":
        run([
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(SCRIPT_DIR / "build_windows_portable.ps1"),
        ])
        return 0

    print(
        "Portable packaging is currently implemented for macOS and Windows. "
        f"Unsupported platform: {platform.system()}",
        file=sys.stderr,
    )
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
