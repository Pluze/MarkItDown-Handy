# Build a portable Windows bundle with a private Python runtime.
# The bundle includes MarkItDown Handy, Python packages, and launchers.

param(
    [string]$PythonVersion = $env:PY_VERSION,
    [string]$AppVersion = "0.1.0"
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($PythonVersion)) {
    $PythonVersion = "3.11"
}

$RootDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$Src = Join-Path $RootDir "src\markitdown_handy.py"
$BuildDir = Join-Path $RootDir "build-windows"
$OutDir = Join-Path $RootDir "dist-windows"
$BundleName = "MarkItDown Handy Portable"
$BundleDir = Join-Path $OutDir $BundleName
$RuntimeDir = Join-Path $BundleDir "runtime"
$AppDir = Join-Path $BundleDir "app"
$ToolsDir = Join-Path $BundleDir "tools"
$TessdataDir = Join-Path $BundleDir "tessdata"
$Arch = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }

Remove-Item -Recurse -Force $BundleDir -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $BuildDir, $OutDir, $RuntimeDir, $AppDir, $ToolsDir, $TessdataDir | Out-Null
$AppSource = Join-Path $AppDir "markitdown_handy.py"
Copy-Item $Src $AppSource -Force

$VenvDir = Join-Path $BuildDir "venv"
if (-not (Test-Path $VenvDir)) {
    py -$PythonVersion -m venv $VenvDir
}

$Python = Join-Path $VenvDir "Scripts\python.exe"
if (-not (Test-Path $Python)) {
    $Python = "python"
}

& $Python -m pip install --upgrade pip
& $Python -m pip install --upgrade "markitdown[all]" tkinterdnd2 ocrmypdf

@'
import importlib.util
missing = [m for m in ["markitdown", "tkinterdnd2", "ocrmypdf"] if not importlib.util.find_spec(m)]
if missing:
    raise SystemExit(f"Missing required modules: {missing}")
print("Python module check OK")
'@ | & $Python -

# Patch only the bundled copy so the Windows portable app can use runtime\Scripts
# and Windows shell quoting without changing the source-tree default behavior.
@'
from __future__ import annotations

import sys
from pathlib import Path

path = Path(sys.argv[1])
text = path.read_text(encoding="utf-8")
start = text.index("def find_ocrmypdf_executable():")
end = text.index("def safe_stem(path):", start)
replacement = r'''def _shell_quote(value):
    if os.name == "nt":
        return subprocess.list2cmdline([str(value)])
    return shlex.quote(str(value))


def _bundled_bin_dir(env_dir):
    scripts_dir = env_dir / "Scripts"
    if scripts_dir.exists():
        return scripts_dir
    return env_dir / "bin"


def _bundled_executable(env_dir, name):
    bin_dir = _bundled_bin_dir(env_dir)
    names = [name]
    if os.name == "nt":
        names = [f"{name}.exe", f"{name}.bat", f"{name}.cmd", name]
    for item in names:
        candidate = bin_dir / item
        if candidate.exists():
            return candidate
    return None


def find_ocrmypdf_executable():
    env_dir = bundled_env_dir()
    if env_dir is not None:
        bundled_cli = _bundled_executable(env_dir, "ocrmypdf")
        bundled_python = _bundled_executable(env_dir, "python")
        if bundled_cli is not None:
            return _shell_quote(bundled_cli)
        if bundled_python is not None:
            return f"{_shell_quote(bundled_python)} -m ocrmypdf"

    candidates = [
        "/opt/homebrew/bin/ocrmypdf",
        "/usr/local/bin/ocrmypdf",
    ]
    for item in candidates:
        if Path(item).exists():
            return item
    return shutil.which("ocrmypdf") or "ocrmypdf"


def default_markitdown_command():
    env_dir = bundled_env_dir()
    if env_dir is not None:
        bundled_cli = _bundled_executable(env_dir, "markitdown")
        bundled_python = _bundled_executable(env_dir, "python")
        if bundled_cli is not None:
            return _shell_quote(bundled_cli)
        if bundled_python is not None:
            return f"{_shell_quote(bundled_python)} -m markitdown"
    return f"{_shell_quote(find_conda_executable())} run -n {CONDA_ENV_NAME} markitdown"


def add_gui_paths(env):
    env = env.copy()
    extra_paths = []
    path_sep = os.pathsep

    env_dir = bundled_env_dir()
    if env_dir is not None:
        bin_dir = _bundled_bin_dir(env_dir)
        extra_paths.append(str(bin_dir))
        tools_dir = resource_dir() / "tools"
        if tools_dir.exists():
            extra_paths.append(str(tools_dir))

        if os.name == "nt":
            dll_dirs = [str(env_dir), str(bin_dir)]
            old_path = env.get("PATH", "")
            env["PATH"] = path_sep.join(extra_paths + dll_dirs + ([old_path] if old_path else []))
        else:
            old_dyld = env.get("DYLD_LIBRARY_PATH", "")
            env["DYLD_LIBRARY_PATH"] = path_sep.join([str(env_dir / "lib")] + ([old_dyld] if old_dyld else []))

        tessdata = resource_dir() / "tessdata"
        if tessdata.exists():
            env["TESSDATA_PREFIX"] = str(tessdata)

    extra_paths.extend([
        "/opt/anaconda3/bin",
        str(Path.home() / "anaconda3/bin"),
        str(Path.home() / "miniconda3/bin"),
        "/opt/homebrew/bin",
        "/usr/local/bin",
        str(Path(sys.executable).parent),
    ])
    env["PATH"] = path_sep.join(extra_paths + [env.get("PATH", "")])
    return env


'''
path.write_text(text[:start] + replacement + text[end:], encoding="utf-8")
'@ | & $Python - $AppSource

Copy-Item -Recurse -Force (Join-Path $VenvDir "*") $RuntimeDir

$RuntimePython = Join-Path $RuntimeDir "Scripts\python.exe"
$RuntimeScripts = Join-Path $RuntimeDir "Scripts"

$LauncherCmd = Join-Path $BundleDir "MarkItDown Handy.cmd"
@"
@echo off
setlocal
set "APP_DIR=%~dp0"
set "MARKITDOWN_BUNDLED_ENV=%APP_DIR%runtime"
set "MARKITDOWN_RESOURCE_DIR=%APP_DIR%"
set "PYTHONNOUSERSITE=1"
set "PATH=%APP_DIR%tools;%APP_DIR%runtime\Scripts;%APP_DIR%runtime;%PATH%"
if exist "%APP_DIR%tessdata" set "TESSDATA_PREFIX=%APP_DIR%tessdata"
"%APP_DIR%runtime\Scripts\python.exe" "%APP_DIR%app\markitdown_handy.py"
endlocal
"@ | Set-Content -Encoding ASCII $LauncherCmd

$LauncherPs1 = Join-Path $BundleDir "MarkItDown Handy.ps1"
@"
`$AppDir = Split-Path -Parent `$MyInvocation.MyCommand.Path
`$env:MARKITDOWN_BUNDLED_ENV = Join-Path `$AppDir 'runtime'
`$env:MARKITDOWN_RESOURCE_DIR = `$AppDir
`$env:PYTHONNOUSERSITE = '1'
`$env:PATH = (Join-Path `$AppDir 'tools') + ';' + (Join-Path `$AppDir 'runtime\Scripts') + ';' + (Join-Path `$AppDir 'runtime') + ';' + `$env:PATH
`$Tessdata = Join-Path `$AppDir 'tessdata'
if (Test-Path `$Tessdata) { `$env:TESSDATA_PREFIX = `$Tessdata }
& (Join-Path `$AppDir 'runtime\Scripts\python.exe') (Join-Path `$AppDir 'app\markitdown_handy.py')
"@ | Set-Content -Encoding UTF8 $LauncherPs1

$RuntimeBin = Join-Path $RuntimeDir "bin"
New-Item -ItemType Directory -Force $RuntimeBin | Out-Null
@"
@echo off
"%~dp0..\Scripts\python.exe" -m markitdown %*
"@ | Set-Content -Encoding ASCII (Join-Path $RuntimeBin "markitdown.bat")
@"
@echo off
"%~dp0..\Scripts\python.exe" -m ocrmypdf %*
"@ | Set-Content -Encoding ASCII (Join-Path $RuntimeBin "ocrmypdf.bat")
Copy-Item -Force (Join-Path $RuntimeScripts "python.exe") (Join-Path $RuntimeBin "python.exe")

$ToolNames = @(
    "tesseract.exe",
    "gswin64c.exe",
    "gswin32c.exe",
    "qpdf.exe"
)
foreach ($tool in $ToolNames) {
    $cmd = Get-Command $tool -ErrorAction SilentlyContinue
    if ($cmd -and (Test-Path $cmd.Source)) {
        Copy-Item -Force $cmd.Source $ToolsDir
        Write-Host "Copied tool: $($cmd.Source)"
    } else {
        Write-Host "Tool not found, skipping: $tool"
    }
}

$TessCandidates = @(
    "$env:ProgramFiles\Tesseract-OCR\tessdata",
    "$env:ProgramFiles(x86)\Tesseract-OCR\tessdata",
    "$env:ChocolateyInstall\lib\tesseract\tools\tessdata"
)
foreach ($candidate in $TessCandidates) {
    if (Test-Path $candidate) {
        foreach ($lang in @("eng", "osd", "chi_sim", "chi_tra")) {
            $srcLang = Join-Path $candidate "$lang.traineddata"
            if (Test-Path $srcLang) {
                Copy-Item -Force $srcLang $TessdataDir
            }
        }
    }
}

@'
import importlib.util
for mod in ["markitdown", "tkinterdnd2", "ocrmypdf"]:
    if not importlib.util.find_spec(mod):
        raise SystemExit(f"Bundled module missing: {mod}")
print("Bundled runtime check OK")
'@ | & $RuntimePython -

$ZipPath = Join-Path $OutDir "MarkItDown_Handy_v${AppVersion}_portable_Windows_${Arch}.zip"
Remove-Item -Force $ZipPath -ErrorAction SilentlyContinue
Compress-Archive -Path $BundleDir -DestinationPath $ZipPath -Force

Write-Host "Built Windows portable bundle: $BundleDir"
Write-Host "Zip: $ZipPath"
