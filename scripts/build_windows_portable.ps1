# Build a portable Windows bundle with a private Python runtime.
# The bundle includes MarkItDown Handy, Python packages, and launchers.

param(
    [string]$PythonVersion = $env:PY_VERSION,
    [string]$PythonFullVersion = $env:PYTHON_FULL_VERSION,
    [string]$AppVersion = "0.1.0"
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($PythonVersion)) {
    $PythonVersion = "3.11"
}
if ([string]::IsNullOrWhiteSpace($PythonFullVersion)) {
    if ($PythonVersion -match '^\d+\.\d+\.\d+$') {
        $PythonFullVersion = $PythonVersion
    } else {
        $PythonFullVersion = "3.11.9"
    }
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

# Do not copy a venv from the CI runner. Windows venvs keep absolute references
# to the original base Python path, which breaks after the bundle is moved.
# Instead, install a private Python distribution directly inside runtime/.
$Installer = Join-Path $BuildDir "python-$PythonFullVersion-amd64.exe"
$InstallerUrl = "https://www.python.org/ftp/python/$PythonFullVersion/python-$PythonFullVersion-amd64.exe"
if (-not (Test-Path $Installer)) {
    Write-Host "Downloading Python runtime installer: $InstallerUrl"
    Invoke-WebRequest -Uri $InstallerUrl -OutFile $Installer
}

Write-Host "Installing private Python runtime into $RuntimeDir"
# Call the installer directly so PowerShell preserves TargetDir as one quoted
# argument. Start-Process -ArgumentList can flatten arrays in a way that lets
# spaces in `MarkItDown Handy Portable` truncate the target path.
& $Installer `
    /quiet `
    InstallAllUsers=0 `
    "TargetDir=$RuntimeDir" `
    Include_pip=1 `
    Include_tcltk=1 `
    Include_launcher=0 `
    InstallLauncherAllUsers=0 `
    PrependPath=0 `
    Shortcuts=0 `
    Include_test=0
if ($LASTEXITCODE -ne 0) {
    throw "Python installer failed with exit code $LASTEXITCODE"
}

$Python = Join-Path $RuntimeDir "python.exe"
if (-not (Test-Path $Python)) {
    Write-Host "Python runtime files found under bundle directory:"
    Get-ChildItem $BundleDir -Recurse -Filter python.exe -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Host "- $($_.FullName)"
    }
    throw "Bundled Python was not installed at expected path: $Python"
}

& $Python -m pip install --upgrade pip
& $Python -m pip install --upgrade "markitdown[all]" tkinterdnd2 ocrmypdf

@'
import importlib.util
import sys
from pathlib import Path

runtime = Path(sys.executable).resolve().parent
missing = [m for m in ["markitdown", "tkinterdnd2", "ocrmypdf", "tkinter"] if not importlib.util.find_spec(m)]
if missing:
    raise SystemExit(f"Missing required modules: {missing}")
if "hostedtoolcache" in str(runtime).lower():
    raise SystemExit(f"Runtime is still using the GitHub hosted toolcache: {runtime}")
print(f"Python module check OK: {sys.executable}")
'@ | & $Python -

# Patch only the bundled copy so the Windows portable app can use runtime\python.exe
# and Windows shell quoting without changing source-tree defaults. Prefer
# `python -m ...` over pip-generated console .exe files because console launchers
# can contain absolute CI build paths.
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


def _bundled_executable(env_dir, name):
    candidates = []
    if os.name == "nt":
        suffixes = [".exe", ".bat", ".cmd", ""]
        search_dirs = [env_dir, env_dir / "bin", env_dir / "Scripts"]
    else:
        suffixes = [""]
        search_dirs = [env_dir / "bin", env_dir]

    for search_dir in search_dirs:
        for suffix in suffixes:
            candidates.append(search_dir / f"{name}{suffix}")

    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def find_ocrmypdf_executable():
    env_dir = bundled_env_dir()
    if env_dir is not None:
        bundled_python = _bundled_executable(env_dir, "python")
        if bundled_python is not None:
            return f"{_shell_quote(bundled_python)} -m ocrmypdf"
        bundled_cli = _bundled_executable(env_dir, "ocrmypdf")
        if bundled_cli is not None:
            return _shell_quote(bundled_cli)

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
        bundled_python = _bundled_executable(env_dir, "python")
        if bundled_python is not None:
            return f"{_shell_quote(bundled_python)} -m markitdown"
        bundled_cli = _bundled_executable(env_dir, "markitdown")
        if bundled_cli is not None:
            return _shell_quote(bundled_cli)
    return f"{_shell_quote(find_conda_executable())} run -n {CONDA_ENV_NAME} markitdown"


def add_gui_paths(env):
    env = env.copy()
    extra_paths = []
    path_sep = os.pathsep

    env_dir = bundled_env_dir()
    if env_dir is not None:
        extra_paths.extend([str(env_dir), str(env_dir / "bin"), str(env_dir / "Scripts")])
        tools_dir = resource_dir() / "tools"
        if tools_dir.exists():
            extra_paths.append(str(tools_dir))

        if os.name != "nt":
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

$RuntimePython = Join-Path $RuntimeDir "python.exe"

$LauncherCmd = Join-Path $BundleDir "MarkItDown Handy.cmd"
@"
@echo off
setlocal
set "APP_DIR=%~dp0"
set "MARKITDOWN_BUNDLED_ENV=%APP_DIR%runtime"
set "MARKITDOWN_RESOURCE_DIR=%APP_DIR%"
set "PYTHONNOUSERSITE=1"
set "PYTHONHOME="
set "PYTHONPATH="
set "PATH=%APP_DIR%tools;%APP_DIR%runtime;%APP_DIR%runtime\bin;%APP_DIR%runtime\Scripts;%PATH%"
if exist "%APP_DIR%tessdata" set "TESSDATA_PREFIX=%APP_DIR%tessdata"
"%APP_DIR%runtime\python.exe" "%APP_DIR%app\markitdown_handy.py"
endlocal
"@ | Set-Content -Encoding ASCII $LauncherCmd

$LauncherPs1 = Join-Path $BundleDir "MarkItDown Handy.ps1"
@"
`$AppDir = Split-Path -Parent `$MyInvocation.MyCommand.Path
`$env:MARKITDOWN_BUNDLED_ENV = Join-Path `$AppDir 'runtime'
`$env:MARKITDOWN_RESOURCE_DIR = `$AppDir
`$env:PYTHONNOUSERSITE = '1'
Remove-Item Env:PYTHONHOME -ErrorAction SilentlyContinue
Remove-Item Env:PYTHONPATH -ErrorAction SilentlyContinue
`$env:PATH = (Join-Path `$AppDir 'tools') + ';' + (Join-Path `$AppDir 'runtime') + ';' + (Join-Path `$AppDir 'runtime\bin') + ';' + (Join-Path `$AppDir 'runtime\Scripts') + ';' + `$env:PATH
`$Tessdata = Join-Path `$AppDir 'tessdata'
if (Test-Path `$Tessdata) { `$env:TESSDATA_PREFIX = `$Tessdata }
& (Join-Path `$AppDir 'runtime\python.exe') (Join-Path `$AppDir 'app\markitdown_handy.py')
"@ | Set-Content -Encoding UTF8 $LauncherPs1

$RuntimeBin = Join-Path $RuntimeDir "bin"
New-Item -ItemType Directory -Force $RuntimeBin | Out-Null
@"
@echo off
"%~dp0..\python.exe" -m markitdown %*
"@ | Set-Content -Encoding ASCII (Join-Path $RuntimeBin "markitdown.bat")
@"
@echo off
"%~dp0..\python.exe" -m ocrmypdf %*
"@ | Set-Content -Encoding ASCII (Join-Path $RuntimeBin "ocrmypdf.bat")

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
import sys
from pathlib import Path
runtime = Path(sys.executable).resolve().parent
for mod in ["markitdown", "tkinterdnd2", "ocrmypdf", "tkinter"]:
    if not importlib.util.find_spec(mod):
        raise SystemExit(f"Bundled module missing: {mod}")
if "hostedtoolcache" in str(runtime).lower():
    raise SystemExit(f"Bundled runtime is not portable: {runtime}")
print(f"Bundled runtime check OK: {sys.executable}")
'@ | & $RuntimePython -

$ZipPath = Join-Path $OutDir "MarkItDown_Handy_v${AppVersion}_portable_Windows_${Arch}.zip"
Remove-Item -Force $ZipPath -ErrorAction SilentlyContinue
Compress-Archive -Path $BundleDir -DestinationPath $ZipPath -Force

Write-Host "Built Windows portable bundle: $BundleDir"
Write-Host "Zip: $ZipPath"
